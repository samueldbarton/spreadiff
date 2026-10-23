import { describe, expect, it } from 'vitest';
import { strToU8, zipSync } from 'fflate';
import { diffTables, diffWorkbooks, parseCsv, parseXlsx, diffText, toUnifiedCsv, ROW_STATUS_NAMES, hasChanges } from '../src/core';
import type { TableDiff } from '../src/core';

const csv = (s: string) => parseCsv(s.trim() + '\n');
const statuses = (d: TableDiff) => Array.from(d.status, (s) => ROW_STATUS_NAMES[s]);

describe('parseCsv', () => {
  it('handles quotes, escaped quotes, newlines in fields and CRLF', () => {
    const t = parseCsv('a,b,c\r\n1,"x, ""y""","multi\nline"\r\n2,,\r\n', { numbers: false });
    expect(t.columns).toEqual(['a', 'b', 'c']);
    expect(t.rows).toEqual([
      ['1', 'x, "y"', 'multi\nline'],
      ['2', null, null],
    ]);
  });
  it('detects tabs and semicolons', () => {
    expect(parseCsv('a\tb\n1\t2').columns).toEqual(['a', 'b']);
    expect(parseCsv('a;b\n1;2').rows).toEqual([[1, 2]]);
  });
  it('converts numbers only when lossless', async () => {
    const { isCanonicalNumber } = await import('../src/core/csv');
    const fields = ['0', '7', '-3', '1.5', '0.25', '0.000001', '123456789012345', '-0.5', '007', '1.50', '1.', '.5', '1e5', '-0', '0.0000001', '1234567890123456', '+1', '-', '1,000', '0x10', ' 1', '-0.0'];
    // Sound (never converts lossily); conservative past 15 significant digits.
    const accepted = fields.filter(isCanonicalNumber);
    expect(accepted).toEqual(['0', '7', '-3', '1.5', '0.25', '0.000001', '123456789012345', '-0.5']);
    let seed = 3;
    const rand = () => ((seed = (seed * 1103515245 + 12345) & 0x7fffffff) / 0x7fffffff);
    for (let k = 0; k < 50000; k++) {
      const f = (rand() < 0.3 ? '-' : '') + (rand() * 10 ** Math.floor(rand() * 18 - 8)).toFixed(Math.floor(rand() * 12));
      if (isCanonicalNumber(f)) expect(String(Number(f)), f).toBe(f);
      else if (f.replace(/[-.]/g, '').replace(/^0+/, '').length <= 15) expect(String(Number(f)) === f, f).toBe(false);
    }
  });
  it('makes duplicate and blank headers unique', () => {
    expect(parseCsv('x,x,\n1,2,3').columns).toEqual(['x', 'x (2)', 'C']);
  });
});

describe('diffTables', () => {
  it('aligns by sequence and pairs similar rows as modified', () => {
    const d = diffTables(
      csv(`name,qty,price\napple,1,0.5\nbanana,2,0.25\ncherry,3,4`),
      csv(`name,qty,price\napple,1,0.5\nbanana,5,0.25\ndate,9,9\ncherry,3,4`),
    );
    expect(statuses(d)).toEqual(['equal', 'modified', 'added', 'equal']);
    expect(Array.from(d.changedCells.get(1)!)).toEqual([0, 1, 0]);
    expect(d.stats).toMatchObject({ rowsAdded: 1, rowsModified: 1, cellsChanged: 1 });
  });

  it('matches by an id column regardless of order', () => {
    const d = diffTables(
      csv(`id,name\n1,a\n2,b\n3,c`),
      csv(`id,name\n3,c\n1,A\n4,d`),
    );
    expect(d.keyColumns).toEqual(['id']);
    const byNew = new Map<number, string>();
    for (let r = 0; r < d.rowCount; r++) byNew.set(d.newRows[r], ROW_STATUS_NAMES[d.status[r]]);
    expect(byNew.get(0)).toBe('equal');
    expect(byNew.get(1)).toBe('modified');
    expect(byNew.get(2)).toBe('added');
    expect(d.stats.rowsRemoved).toBe(1);
  });

  it('places removed rows where they were in keyed mode', () => {
    const d = diffTables(csv(`id,v\n1,a\n2,b\n3,c`), csv(`id,v\n1,a\n3,c`));
    expect(Array.from(d.oldRows)).toEqual([0, 1, 2]);
    expect(statuses(d)).toEqual(['equal', 'removed', 'equal']);
  });

  it('aligns added and removed columns', () => {
    const d = diffTables(csv(`a,b,c\n1,2,3`), csv(`a,c,d\n1,3,4`));
    expect(d.columns.map((c) => `${c.name}:${c.status}`)).toEqual(['a:equal', 'b:removed', 'c:equal', 'd:added']);
    expect(statuses(d)).toEqual(['equal']);
  });

  it('detects renamed columns from their contents', () => {
    const d = diffTables(
      csv(`id,qty\n1,10\n2,20\n3,30\n4,40\n5,50\n6,60`),
      csv(`id,quantity\n1,10\n2,20\n3,31\n4,40\n5,50\n6,60`),
    );
    const col = d.columns.find((c) => c.name === 'quantity')!;
    expect(col.status).toBe('renamed');
    expect(col.oldName).toBe('qty');
    expect(d.stats.cellsChanged).toBe(1);
  });

  it('respects numeric tolerance, whitespace and case options', () => {
    const a = csv(`k,v\nx,1.0000001\ny, Hello`);
    const b = csv(`k,v\nx,1\ny,hello`);
    expect(diffTables(a, b).stats.cellsChanged).toBe(2);
    expect(diffTables(a, b, { numericTolerance: 1e-3, ignoreCase: true }).stats.cellsChanged).toBe(0);
  });

  it('diffs workbooks sheet by sheet', () => {
    const t = (name: string, s: string) => ({ ...csv(s), name });
    const d = diffWorkbooks(
      { name: 'a', sheets: [t('One', 'a\n1'), t('Gone', 'a\n1')] },
      { name: 'b', sheets: [t('One', 'a\n2'), t('New', 'a\n1')] },
    );
    expect(d.sheets.map((s) => `${s.name}:${s.status}`)).toEqual(['One:modified', 'Gone:removed', 'New:added']);
  });

  it('exports a daff-style highlighter CSV', () => {
    const d = diffTables(csv(`id,v\n1,a\n2,b`), csv(`id,v\n1,z\n3,c`));
    expect(toUnifiedCsv(d)).toBe('@@,id,v\n->,1,a->z\n---,2,b\n+++,3,c\n');
  });
});

describe('diffText', () => {
  it('produces minimal char spans', () => {
    expect(diffText('kitten', 'sitting')).toEqual([
      { text: 'k', kind: 'removed' },
      { text: 's', kind: 'added' },
      { text: 'itt', kind: 'equal' },
      { text: 'e', kind: 'removed' },
      { text: 'i', kind: 'added' },
      { text: 'n', kind: 'equal' },
      { text: 'g', kind: 'added' },
    ]);
  });
});

describe('parseXlsx', () => {
  it('reads shared strings, numbers, dates, booleans and shared formulas', () => {
    const files = {
      '[Content_Types].xml': strToU8('<Types/>'),
      'xl/workbook.xml': strToU8(
        '<workbook><sheets><sheet name="Data &amp; More" sheetId="1" r:id="rId1"/></sheets></workbook>',
      ),
      'xl/_rels/workbook.xml.rels': strToU8(
        '<Relationships><Relationship Id="rId1" Type="ws" Target="worksheets/sheet1.xml"/></Relationships>',
      ),
      'xl/sharedStrings.xml': strToU8(
        '<sst><si><t>name</t></si><si><t>when</t></si><si><r><t>Ri</t></r><r><t>ch</t></r></si><si><t>n</t></si><si><t>x2</t></si></sst>',
      ),
      'xl/styles.xml': strToU8('<styleSheet><cellXfs><xf numFmtId="0"/><xf numFmtId="14"/></cellXfs></styleSheet>'),
      'xl/worksheets/sheet1.xml': strToU8(
        '<worksheet><sheetData>' +
          '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="s"><v>3</v></c><c r="D1" t="s"><v>4</v></c></row>' +
          '<row r="2"><c r="A2" t="s"><v>2</v></c><c r="B2" s="1"><v>45292</v></c><c r="C2"><v>1.5</v></c><c r="D2"><f t="shared" ref="D2:D3" si="0">C2*2</f><v>3</v></c></row>' +
          '<row r="3"><c r="A3" t="b"><v>1</v></c><c r="C3"><v>2</v></c><c r="D3"><f t="shared" si="0"/><v>4</v></c></row>' +
          '</sheetData></worksheet>',
      ),
    };
    const wb = parseXlsx(zipSync(files));
    expect(wb.sheets).toHaveLength(1);
    const s = wb.sheets[0];
    expect(s.name).toBe('Data & More');
    expect(s.columns).toEqual(['name', 'when', 'n', 'x2']);
    expect(s.rows).toEqual([
      ['Rich', '2024-01-01', 1.5, 3],
      [true, null, 2, 4],
    ]);
    expect(s.formulas![1][3]).toBe('=C3*2');
  });
});

describe('pack/unpack', () => {
  it('round-trips tables losslessly', async () => {
    const { packTable, unpackTable } = await import('../src/core');
    const t = { name: 's', columns: ['a', 'b', 'c'], rows: [['x', 1.5, true], [null, -0, false], ['y\u0000z', NaN, null]] as any, formulas: [[null, '=A1', null], [null, null, null], [null, null, null]] };
    const u = unpackTable(packTable(t));
    expect(u.rows[0]).toEqual(['x', 1.5, true]);
    expect(u.rows[1]).toEqual([null, -0, false]);
    expect(u.rows[2][0]).toBe('y�z');
    expect(u.rows[2][1]).toBeNaN();
    expect(u.formulas).toEqual(t.formulas);
    expect(unpackTable(packTable({ ...t, origin: { row: 4, col: 2 } })).origin).toEqual({ row: 4, col: 2 });
  });
});

describe('moved rows (keyed)', () => {
  const t = (ids: number[], edit?: number) => ({ name: 't', columns: ['id', 'v'], rows: ids.map((i) => [i, i === edit ? 'changed' : `v${i}`]) });
  const movedIds = (d: ReturnType<typeof diffTables>) => Array.from({ length: d.rowCount }, (_, r) => r).filter((r) => d.moved[r]).map((r) => d.newTable.rows[d.newRows[r]][0]);

  it('flags the one row that moved, not everything between', () => {
    const d = diffTables(t([1, 2, 3, 4, 5]), t([1, 4, 2, 3, 5]), { keys: ['id'] });
    expect(movedIds(d)).toEqual([4]);
    expect(d.stats).toMatchObject({ rowsMoved: 1, rowsEqual: 5, rowsModified: 0 });
    expect(hasChanges(d)).toBe(true);
  });

  it('flags the fewest rows that explain a swap', () => {
    const d = diffTables(t([1, 2, 3, 4, 5]), t([1, 4, 3, 2, 5]), { keys: ['id'] });
    expect(d.stats.rowsMoved).toBe(2);
  });

  it('is independent of modification, and absent without keys', () => {
    const d = diffTables(t([1, 2, 3]), t([3, 1, 2], 3), { keys: ['id'] });
    expect(movedIds(d)).toEqual([3]);
    expect(d.stats).toMatchObject({ rowsMoved: 1, rowsModified: 1 });
    const seq = diffTables(t([1, 2, 3]), t([3, 1, 2]), { keys: 'none' });
    expect(seq.stats.rowsMoved).toBe(0);
    expect(diffTables(t([1, 2, 3]), t([1, 2, 3]), { keys: ['id'] }).stats.rowsMoved).toBe(0);
  });
});
