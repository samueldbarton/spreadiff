import { execFileSync } from 'node:child_process';
import { mkdtempSync, readFileSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import ExcelJS from 'exceljs';
import { describe, expect, it } from 'vitest';
import {
  applyDecisions,
  applyDecisionsToXlsx,
  decideRange,
  decisionOf,
  diffTables,
  diffWorkbooks,
  noDecisions,
  parseCsv,
  parseXlsx,
  tableToCsv,
  translateFormula,
  type SheetDecisions,
} from '../src/core';

describe('accept and reject', () => {
  const a = parseCsv('id,name,price\n1,Lamp,10\n2,Chair,20\n3,Desk,30');
  const b = parseCsv('id,name,price\n1,Lamp,12\n3,Desk,35\n4,Rug,50');
  const d = diffTables(a, b);
  const row = (id: number) => [...Array(d.rowCount).keys()].find((r) => (d.newRows[r] >= 0 ? d.newTable.rows[d.newRows[r]][0] : d.oldTable.rows[d.oldRows[r]][0]) === id)!;

  it('keeps everything new when nothing is decided', () => {
    expect(applyDecisions(d, noDecisions()).rows).toEqual(b.rows);
  });

  it('rejects cells, drops added rows and restores removed ones where they were', () => {
    const dec = noDecisions();
    dec.cells.set(row(1) * 3 + 2, 'reject'); // Lamp keeps 10
    dec.rows.set(row(4), 'reject'); // no Rug
    dec.rows.set(row(2), 'reject'); // Chair stays
    const t = applyDecisions(d, dec);
    expect(t.rows).toEqual([[1, 'Lamp', 10], [2, 'Chair', 20], [3, 'Desk', 35]]);
    expect(tableToCsv(t)).toBe('id,name,price\r\n1,Lamp,10\r\n2,Chair,20\r\n3,Desk,35\r\n');
    expect(decisionOf(d, dec, row(1), 2)).toBe('reject');
    expect(decisionOf(d, dec, row(3), 2)).toBe('accept');
  });

  it('decides a selected range: its changed cells and its added or removed rows', () => {
    const all = [...Array(d.rowCount).keys()];
    const dec = decideRange(d, noDecisions(), all, [2], 'reject');
    expect(applyDecisions(d, dec).rows).toEqual([[1, 'Lamp', 10], [2, 'Chair', 20], [3, 'Desk', 30]]);
    expect(decideRange(d, dec, all, [2], undefined)).toEqual(noDecisions());
  });

  it('undoes column changes', () => {
    const c1 = parseCsv('id,price,note\n1,10,x');
    const c2 = parseCsv('id,cost,extra\n1,10,y');
    const cd = diffTables(c1, c2);
    const dec: SheetDecisions = noDecisions();
    cd.columns.forEach((c, i) => c.status !== 'equal' && dec.columns.set(i, 'reject'));
    const t = applyDecisions(cd, dec);
    expect(t.columns).toEqual(['id', 'price', 'note']);
    expect(t.rows).toEqual([[1, 10, 'x']]);
  });
});

/** A model: inputs in B2:B3, a total formula below. `insert` adds a row above the total. */
async function model(price: number, units: number, insert = false): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Model');
  ws.addRow(['item', 'value']);
  ws.addRow(['Price', price]);
  ws.addRow(['Units', units]);
  if (insert) ws.addRow(['Discount', 0]);
  const r = ws.rowCount + 1;
  ws.addRow(['Revenue', { formula: 'B2*B3', result: price * units }]);
  ws.getCell(`B${r}`).numFmt = '$#,##0';
  ws.addRow(['Note', 'draft']);
  return new Uint8Array(await wb.xlsx.writeBuffer());
}

describe('applyDecisionsToXlsx', () => {
  it('writes rejected cells back into the new file, keeping their style', async () => {
    const oldBytes = await model(10, 100);
    const newBytes = await model(12, 100);
    const wd = diffWorkbooks(parseXlsx(oldBytes), parseXlsx(newBytes));
    const s = wd.sheets[0].diff;
    const price = [...s.changedCells.keys()][0];
    const dec = noDecisions();
    dec.cells.set(price * s.columns.length + 1, 'reject');
    const { file, applied, recalculated, skipped } = applyDecisionsToXlsx(newBytes, wd, [dec]);
    expect(applied).toBe(1);
    expect(recalculated).toBe(1);
    expect(skipped).toEqual([]);
    const back = parseXlsx(file).sheets[0];
    expect(back.rows.map((r) => r[1])).toEqual([10, 100, 1000, 'draft']); // the total's saved result is recalculated too
    expect(back.formulas?.[2]?.[1]).toBe('=B2*B3');
    expect(back.numberFormats && back.numberFormats.codes[back.numberFormats.rows[2]![1]]).toBe('$#,##0');
    const wbXml = new TextDecoder().decode((await import('fflate')).unzipSync(file)['xl/workbook.xml']);
    expect(wbXml).toMatch(/fullCalcOnLoad="1"/);
  });

  it('moves an old formula’s references to where its cells are now', async () => {
    const oldBytes = await model(10, 100);
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load((await model(10, 100, true)) as never);
    wb.getWorksheet('Model')!.getCell('B5').value = { formula: 'B2*B3*(1-B4)', result: 1000 };
    const newBytes = new Uint8Array(await wb.xlsx.writeBuffer());
    const wd = diffWorkbooks(parseXlsx(oldBytes), parseXlsx(newBytes));
    expect(translateFormula('=B2*B3+SUM(B2:B3)', 'Model', wd)).toBe('=B2*B3+SUM(B2:B3)');
    expect(translateFormula('=$B$5&Model!B5', 'Model', wd)).toBe('=$B$6&Model!B6'); // the note moved down a row
    const s = wd.sheets[0].diff;
    const total = [...s.changedCells.keys()].find((r) => s.newTable.rows[s.newRows[r]][0] === 'Revenue')!;
    const dec = noDecisions();
    dec.cells.set(total * s.columns.length + 1, 'reject');
    dec.rows.set([...Array(s.rowCount).keys()].find((r) => s.newRows[r] >= 0 && s.oldRows[r] < 0)!, 'reject');
    const { file, applied, skipped } = applyDecisionsToXlsx(newBytes, wd, [dec]);
    expect(applied).toBe(1);
    expect(skipped.map((x) => x.target)).toEqual(['row 4']); // the Discount row, which can't be deleted in place
    expect(parseXlsx(file).sheets[0].formulas?.[3]?.[1]).toBe('=B2*B3');
  });

  it('skips a formula that reads a row the new file no longer has', () => {
    const wd = diffWorkbooks(
      { name: 'a', sheets: [{ name: 'S', columns: ['A'], rows: [[1], [2], [3]], formulas: [], origin: { row: 0, col: 0 } }] },
      { name: 'b', sheets: [{ name: 'S', columns: ['A'], rows: [[1], [3]], formulas: [], origin: { row: 0, col: 0 } }] },
      { keys: 'none' },
    );
    expect(translateFormula('=A2+A3', 'S', wd)).toBeNull();
    expect(translateFormula('=A1+A3', 'S', wd)).toBe('=A1+A2');
  });

  const soffice = (() => {
    try {
      execFileSync('soffice', ['--version'], { stdio: 'ignore', timeout: 30_000 });
      return true;
    } catch {
      return false;
    }
  })();
  it.skipIf(!soffice)('writes a file LibreOffice opens, with the results already recalculated', async () => {
    const oldBytes = await model(10, 100);
    const newBytes = await model(12, 150);
    const wd = diffWorkbooks(parseXlsx(oldBytes), parseXlsx(newBytes));
    const s = wd.sheets[0].diff;
    const price = [...s.changedCells.keys()].find((r) => s.newTable.rows[s.newRows[r]][0] === 'Price')!;
    const dec = noDecisions();
    dec.cells.set(price * s.columns.length + 1, 'reject');
    const { file } = applyDecisionsToXlsx(newBytes, wd, [dec]);
    const dir = mkdtempSync(join(tmpdir(), 'spreadiff-review-'));
    writeFileSync(join(dir, 'merged.xlsx'), file);
    execFileSync('soffice', ['--headless', '--convert-to', 'csv', '--outdir', dir, join(dir, 'merged.xlsx')], { stdio: 'ignore', timeout: 120_000 });
    const csv = readFileSync(join(dir, 'merged.csv'), 'utf8');
    expect(csv).toContain('Price,10');
    expect(csv).toContain('Units,150');
    expect(csv).toContain('Revenue,1500'); // 10 × 150, though LibreOffice didn't recalculate on open
  }, 180_000);
});
