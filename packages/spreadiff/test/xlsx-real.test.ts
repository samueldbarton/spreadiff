// Parser and diff behaviour found by testing against real Excel and
// LibreOffice files (see docs/tracing.md, "How it was tested"). Each test
// builds the XML the way that application writes it.
import { zipSync, strToU8 } from 'fflate';
import { describe, expect, it } from 'vitest';
import { computeMissingValues, diffWorkbooks, formulaRefs, packWorkbookDiff, scopedNames, packWorkbookTrace, parseFile, parseXlsx, structuredRange, traceWorkbook, unpackWorkbookDiff, unpackWorkbookTrace, type TableDefinition } from '../src/core';

interface SheetSpec {
  name: string;
  /** `<sheetData>` content. */
  data: string;
  /** Table part XML, if the sheet has an Excel table. */
  table?: string;
  /** Merged ranges (`A1:C1`). */
  merges?: string[];
}

function xlsx(sheets: SheetSpec[], names: Record<string, string> = {}): Uint8Array {
  const files: Record<string, Uint8Array> = {};
  // "Rate@1" defines Rate scoped to sheet 1 (localSheetId).
  const defined = Object.entries(names)
    .map(([k, v]) => {
      const m = /^(.+)@(\d+)$/.exec(k);
      return m ? `<definedName name="${m[1]}" localSheetId="${m[2]}">${v}</definedName>` : `<definedName name="${k}">${v}</definedName>`;
    })
    .join('');
  files['xl/workbook.xml'] = strToU8(
    `<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${sheets
      .map((s, i) => `<sheet name="${s.name}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`)
      .join('')}</sheets>${defined ? `<definedNames>${defined}</definedNames>` : ''}</workbook>`,
  );
  files['xl/_rels/workbook.xml.rels'] = strToU8(
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheets
      .map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`)
      .join('')}</Relationships>`,
  );
  sheets.forEach((s, i) => {
    files[`xl/worksheets/sheet${i + 1}.xml`] = strToU8(
      `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${s.data}</sheetData>${s.merges ? `<mergeCells count="${s.merges.length}">${s.merges.map((m) => `<mergeCell ref="${m}"/>`).join('')}</mergeCells>` : ''}</worksheet>`,
    );
    if (s.table) {
      files[`xl/worksheets/_rels/sheet${i + 1}.xml.rels`] = strToU8(
        `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table${i + 1}.xml"/></Relationships>`,
      );
      files[`xl/tables/table${i + 1}.xml`] = strToU8(s.table);
    }
  });
  return zipSync(files);
}

const num = (ref: string, v: number, f?: string, extra = '') => `<c r="${ref}"${extra}>${f !== undefined ? `<f>${f}</f>` : ''}<v>${v}</v></c>`;
const str = (ref: string, v: string) => `<c r="${ref}" t="inlineStr"><is><t>${v}</t></is></c>`;
const row = (r: number, ...cells: string[]) => `<row r="${r}">${cells.join('')}</row>`;

describe('header row detection', () => {
  it('keeps a text header that no formula reads', () => {
    const wb = parseXlsx(xlsx([{ name: 'S', data: row(1, str('A1', 'item'), str('B1', 'qty')) + row(2, str('A2', 'pen'), num('B2', 3)) + row(3, str('A3', 'sum'), num('B3', 3, 'SUM(B2:B2)')) }]));
    expect(wb.sheets[0].columns).toEqual(['item', 'qty']);
    expect(wb.sheets[0].origin).toEqual({ row: 1, col: 0 });
  });

  it('treats row 1 as data when a formula reads it (inputs at the top of a model)', () => {
    const data = row(1, str('A1', 'rate'), num('B1', 0.05)) + row(2, str('A2', 'x'), num('B2', 100, 'B1*2000'));
    const wb = parseXlsx(xlsx([{ name: 'S', data }]));
    expect(wb.sheets[0].columns).toEqual(['A', 'B']);
    expect(wb.sheets[0].rows[0]).toEqual(['rate', 0.05]);
  });

  it('treats years across the top as names, unless a formula reads them', () => {
    const plain = parseXlsx(xlsx([{ name: 'S', data: row(1, str('A1', 'line'), num('B1', 2025), num('C1', 2026)) + row(2, str('A2', 'rev'), num('B2', 1), num('C2', 2)) }]));
    expect(plain.sheets[0].columns).toEqual(['line', '2025', '2026']);
    const read = parseXlsx(xlsx([{ name: 'S', data: row(1, str('A1', 'line'), num('B1', 2025), num('C1', 2026)) + row(2, str('A2', 'age'), num('B2', 0, '2030-B1'), num('C2', 0, '2030-C1')) }]));
    expect(read.sheets[0].origin).toEqual({ row: 0, col: 0 }); // the years are inputs: row 1 is data
  });

  it('sees references from other sheets, including 3-D ones', () => {
    const wb = parseXlsx(
      xlsx([
        { name: 'Jan', data: row(1, str('A1', 'total')) + row(2, num('A2', 5)) },
        { name: 'Feb', data: row(1, str('A1', 'total')) + row(2, num('A2', 7)) },
        { name: 'Sum', data: row(1, num('A1', 0, 'SUM(Jan:Feb!A1)')) },
      ]),
    );
    expect(wb.sheets[0].columns).toEqual(['A']);
    expect(wb.sheets[1].columns).toEqual(['A']);
  });

  it('can be forced either way', () => {
    const data = row(1, str('A1', 'rate'), num('B1', 0.05)) + row(2, str('A2', 'x'), num('B2', 100, 'B1*2000'));
    expect(parseXlsx(xlsx([{ name: 'S', data }]), { header: true }).sheets[0].columns).toEqual(['rate', '0.05']);
    expect(parseXlsx(xlsx([{ name: 'S', data: row(1, str('A1', 'a')) }]), { header: false }).sheets[0].columns).toEqual(['A']);
  });
});

describe('formulas as Excel stores them', () => {
  it('shifts whole-column and whole-row references in shared formulas', () => {
    const data = row(
      1,
      '<c r="A1"><f t="shared" ref="A1:C1" si="0">SUM(A:A)+SUM(2:2)+A$3</f><v>0</v></c>',
      '<c r="B1"><f t="shared" si="0"/><v>0</v></c>',
      '<c r="C1"><f t="shared" si="0"/><v>0</v></c>',
    );
    const f = parseXlsx(xlsx([{ name: 'S', data }]), { header: false }).sheets[0].formulas![0];
    expect(f).toEqual(['=SUM(A:A)+SUM(2:2)+A$3', '=SUM(B:B)+SUM(2:2)+B$3', '=SUM(C:C)+SUM(2:2)+C$3']);
    const down = row(1, '<c r="A1"><f t="shared" ref="A1:A2" si="0">SUM(B1:B1)+SUM($C:C)+SUM(1:1)</f><v>0</v></c>') + row(2, '<c r="A2"><f t="shared" si="0"/><v>0</v></c>');
    expect(parseXlsx(xlsx([{ name: 'S', data: down }]), { header: false }).sheets[0].formulas![1][0]).toBe('=SUM(B2:B2)+SUM($C:C)+SUM(2:2)');
  });

  it('gives every cell of an array formula (or dynamic-array spill) the formula', () => {
    const data = row(1, num('A1', 1), num('B1', 2)) + row(2, '<c r="A2"><f t="array" ref="A2:B2">A1:B1*10</f><v>10</v></c>', num('B2', 20));
    const t = parseXlsx(xlsx([{ name: 'S', data }]), { header: false }).sheets[0];
    expect(t.formulas![1]).toEqual(['=A1:B1*10', '=A1:B1*10']);
    // So a changed spill cell is a computed result, not a typed edit.
    const edited = row(1, num('A1', 1), num('B1', 3)) + row(2, '<c r="A2"><f t="array" ref="A2:B2">A1:B1*10</f><v>10</v></c>', num('B2', 30));
    const d = diffWorkbooks(parseXlsx(xlsx([{ name: 'S', data }]), { header: false }), parseXlsx(xlsx([{ name: 'S', data: edited }]), { header: false }));
    const tr = traceWorkbook(d);
    expect(tr.get({ sheet: 0, row: 1, col: 1 })).toMatchObject({ kind: 'computed', causes: [{ sheet: 0, row: 0, col: 1 }] });
  });

  it('reads LibreOffice-style error constants and TRUE() as the same as Excel', () => {
    const excel = row(1, '<c r="A1" t="e"><v>#VALUE!</v></c>', '<c r="B1" t="b"><f>NOT(IF(1&gt;2,TRUE,FALSE))</f><v>1</v></c>');
    const lo = row(1, '<c r="A1" t="e"><f>#VALUE!</f><v>#VALUE!</v></c>', '<c r="B1" t="b"><f>NOT(IF(1&gt;2,TRUE(),FALSE()))</f><v>1</v></c>');
    const d = diffWorkbooks(parseXlsx(xlsx([{ name: 'S', data: excel }]), { header: false }), parseXlsx(xlsx([{ name: 'S', data: lo }]), { header: false }));
    expect(d.sheets[0].diff.stats.cellsChanged).toBe(0);
  });

  it('treats numbers equal to 15 significant digits as equal, as Excel does', () => {
    const a = row(1, str('A1', 'x'), num('B1', 0.9382716049382716), num('C1', 19.6282307692307));
    const b = row(1, str('A1', 'x'), num('B1', 0.938271604938272), num('C1', 19.6282307692308));
    const d = diffWorkbooks(parseXlsx(xlsx([{ name: 'S', data: a }]), { header: false }), parseXlsx(xlsx([{ name: 'S', data: b }]), { header: false }));
    expect(d.sheets[0].diff.stats.cellsChanged).toBe(0);
    const exact = diffWorkbooks(parseXlsx(xlsx([{ name: 'S', data: a }]), { header: false }), parseXlsx(xlsx([{ name: 'S', data: b }]), { header: false }), { precision: 17 });
    expect(exact.sheets[0].diff.stats.cellsChanged).toBe(2);
  });

  it('reads CR LF in text as LF (XML end-of-line handling)', () => {
    const a = parseXlsx(xlsx([{ name: 'S', data: row(1, '<c r="A1" t="inlineStr"><is><t>a\r\nb</t></is></c>', '<c r="B1" t="inlineStr"><is><t>c_x000D_d</t></is></c>') }]), { header: false });
    expect(a.sheets[0].rows[0]).toEqual(['a\nb', 'c\rd']);
  });
});

describe('structured references', () => {
  const tables: TableDefinition[] = [{ name: 'Sales', sheet: 'S', range: { r0: 0, c0: 1, r1: 5, c1: 3 }, columns: ['Qty', 'Price', 'Total '], headerRows: 1, totalsRows: 1 }];
  const at = (spec: string, row?: number) => structuredRange('Sales', spec, { tables, row });

  it('resolves columns, rows and specials', () => {
    expect(at('[Qty]')).toEqual({ sheet: 'S', r0: 1, r1: 4, c0: 1, c1: 1 });
    expect(at('[[#This Row],[Price]]', 3)).toEqual({ sheet: 'S', r0: 3, r1: 3, c0: 2, c1: 2 });
    expect(at('[@Price]', 2)).toEqual({ sheet: 'S', r0: 2, r1: 2, c0: 2, c1: 2 });
    expect(at('[[#Totals],[Total ]]')).toEqual({ sheet: 'S', r0: 5, r1: 5, c0: 3, c1: 3 });
    expect(at('[[#Headers],[Qty]:[Price]]')).toEqual({ sheet: 'S', r0: 0, r1: 0, c0: 1, c1: 2 });
    expect(at('[#All]')).toEqual({ sheet: 'S', r0: 0, r1: 5, c0: 1, c1: 3 });
    expect(at('[]')).toEqual({ sheet: 'S', r0: 1, r1: 4, c0: 1, c1: 3 });
    expect(at('[Nope]')).toBeNull();
    expect(structuredRange('Other', '[Qty]', { tables })).toBeNull();
  });

  it('feeds tracing from a real table part', () => {
    const table = `<?xml version="1.0"?><table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Sales" displayName="Sales" ref="A1:B3" totalsRowCount="1"><tableColumns count="2"><tableColumn id="1" name="Qty"/><tableColumn id="2" name="Line"/></tableColumns></table>`;
    const book = (q: number) =>
      parseXlsx(
        xlsx([
          {
            name: 'S',
            table,
            data:
              row(1, str('A1', 'Qty'), str('B1', 'Line')) +
              row(2, num('A2', q), num('B2', q * 2, 'Sales[[#This Row],[Qty]]*2')) +
              row(3, num('A3', q), num('B3', q * 2, 'SUBTOTAL(109,Sales[Line])')),
          },
        ]),
      );
    const wb = book(1);
    expect(wb.tables?.[0]).toMatchObject({ name: 'Sales', columns: ['Qty', 'Line'], totalsRows: 1 });
    expect(formulaRefs('=SUBTOTAL(109,Sales[Line])', undefined, { tables: wb.tables }).dynamic).toBe(false);
    const tr = traceWorkbook(diffWorkbooks(wb, book(4)));
    expect(tr.stats).toMatchObject({ inputs: 2, computed: 2, unexplained: 0 });
  });
});

describe('row pairing in model sheets', () => {
  it('keeps a row whose only typed value changed as modified, not removed + added', () => {
    // Angle → SIN, COS: one constant per row, the rest formulas.
    const sheet = (a: number) => row(1, num('A1', a), num('B1', Math.sin(a), 'SIN(A1)'), num('C1', Math.cos(a), 'COS(A1)')) + row(2, num('A2', 2), num('B2', Math.sin(2), 'SIN(A2)'));
    const d = diffWorkbooks(parseXlsx(xlsx([{ name: 'S', data: sheet(1) }]), { header: false }), parseXlsx(xlsx([{ name: 'S', data: sheet(5) }]), { header: false }));
    expect(d.sheets[0].diff.stats).toMatchObject({ rowsModified: 1, rowsAdded: 0, rowsRemoved: 0 });
  });

  it('pairs rows in place when an input moved every row of a model', () => {
    const sheet = (x: number) => [1, 2, 3, 4].map((r) => row(r, num(`A${r}`, x * r, r === 1 ? undefined : `A${r - 1}+${x}`))).join('');
    const d = diffWorkbooks(parseXlsx(xlsx([{ name: 'S', data: sheet(1) }]), { header: false }), parseXlsx(xlsx([{ name: 'S', data: sheet(2) }]), { header: false }));
    expect(d.sheets[0].diff.stats).toMatchObject({ rowsModified: 4, rowsAdded: 0, rowsRemoved: 0 });
  });
});

describe('hosted use: worker hand-off', () => {
  // Inputs in row 1 (read by formulas), a chain of formulas, and a SUM over a range.
  const book = (price: number, qty: number) =>
    xlsx([
      {
        name: 'Model',
        data:
          row(1, num('A1', price), num('B1', qty)) +
          row(2, num('A2', price * qty, 'A1*B1'), num('B2', price * qty * 1.1, 'A2*1.1')) +
          row(3, num('A3', price * qty * 2.1, 'SUM(A2:B2)')),
      },
    ]);

  it('parseFile defaults to header detection, as parseXlsx does', () => {
    const wb = parseFile('m.xlsx', book(10, 3));
    expect(wb.sheets[0].origin).toEqual({ row: 0, col: 0 }); // row 1 holds inputs: it's data
    expect(parseFile('m.xlsx', book(10, 3), { header: true }).sheets[0].origin).toEqual({ row: 1, col: 0 });
  });

  it('a packed trace answers exactly like the original after structured cloning', () => {
    const d = diffWorkbooks(parseFile('a.xlsx', book(10, 3)), parseFile('b.xlsx', book(12, 3)));
    const tr = traceWorkbook(d);
    expect(tr.stats.computed).toBe(3);
    // What a worker does: pack, post (structured clone), unpack.
    const packedTrace = structuredClone(packWorkbookTrace(tr, d).packed);
    const { packed } = packWorkbookDiff(d);
    const d2 = unpackWorkbookDiff(structuredClone(packed));
    const tr2 = unpackWorkbookTrace(packedTrace, d2);
    expect(tr2.stats).toEqual(tr.stats);
    expect(tr2.truncated).toBe(tr.truncated);
    expect([...tr2.sheet(0)]).toEqual([...tr.sheet(0)]);
    const width = d.sheets[0].diff.columns.length;
    for (const [k, t] of tr.sheet(0)) {
      const ref = { sheet: 0, row: Math.floor(k / width), col: k % width };
      expect(tr2.get(ref)).toEqual(t);
      if (t.kind === 'computed') expect(tr2.parents(ref)).toEqual(tr.parents(ref));
    }
    // The SUM's parents are the two changed cells it reads.
    expect(tr2.parents({ sheet: 0, row: 2, col: 0 })).toHaveLength(2);
  });
});

describe('hosted use: packed traces of real files', () => {
  const dir = new URL('../../../test/hosted/fixtures/', import.meta.url);
  it.each([
    ['three_statement', 'new.xlsx'],
    ['three_statement', 'stale.xlsx'],
    ['budget', 'new.xlsx'],
    ['Tables', 'new.xlsx'],
    ['61652', 'new.xlsx'],
  ])('%s / %s round-trips every trace and parent list', async (name, after) => {
    const { readFileSync } = await import('node:fs');
    const read = (f: string) => parseXlsx(new Uint8Array(readFileSync(new URL(`${name}/${f}`, dir))));
    const d = diffWorkbooks(read('old.xlsx'), read(after));
    const tr = traceWorkbook(d);
    const { packed: pt } = packWorkbookTrace(tr, d);
    const d2 = unpackWorkbookDiff(structuredClone(packWorkbookDiff(d).packed));
    const tr2 = unpackWorkbookTrace(structuredClone(pt), d2);
    expect(tr2.stats).toEqual(tr.stats);
    d.sheets.forEach((s, si) => {
      expect([...tr2.sheet(si)]).toEqual([...tr.sheet(si)]);
      const width = s.diff.columns.length;
      for (const [k, t] of tr.sheet(si)) {
        if (t.kind !== 'computed') continue;
        const ref = { sheet: si, row: Math.floor(k / width), col: k % width };
        expect(tr2.parents(ref)).toEqual(tr.parents(ref));
      }
    });
    if (after === 'stale.xlsx') expect(tr2.stats.stale).toBeGreaterThan(0);
  });
});

describe('row identity: relative formula hash', () => {
  it('is the same down a copied column, across digit boundaries, and differs otherwise', async () => {
    const { relativeFormulaHash: h } = await import('../src/core/diff');
    const down = new Set(Array.from({ length: 2000 }, (_, i) => h(`=B${i + 2}*C${i + 2}+$D$1`, i + 1, 4)));
    expect(down.size).toBe(1);
    expect(h('=A3+1', 3, 0)).not.toBe(h('=A1+2', 1, 0)); // the case that needs it: same shape of value, different formula
    expect(h('=A1', 1, 0)).not.toBe(h('=$A$1', 1, 0));
    expect(h('=$A$7', 9, 0)).not.toBe(h('=$A$007', 9, 0));
    expect(h('="B2"&A1', 0, 0)).not.toBe(h('="B3"&A2', 1, 0)); // text in quotes isn't a reference
    expect(h('=$$A1', 0, 0)).toBe(h('=$$A2', 1, 0)); // the second $ starts the reference, as the regex reads it
    expect(h('=VLOOKUP(A2,T,2)', 1, 1)).toBe(h('=VLOOKUP(A3,T,2)', 2, 1));
  });
});

describe("header: 'auto' on data sheets (no regression from header: true)", () => {
  it('keeps year column names', () => {
    const wb = parseXlsx(xlsx([{ name: 'S', data: row(1, str('A1', 'Region'), num('B1', 2023), num('C1', 2024)) + row(2, str('A2', 'EMEA'), num('B2', 1), num('C2', 2)) }]));
    expect(wb.sheets[0].columns).toEqual(['Region', '2023', '2024']);
    expect(wb.sheets[0].rows).toHaveLength(1);
  });

  it('names columns from a label row it keeps as data, so keys still work', () => {
    // A parameter (J1-style) among the labels, read by every formula: row 1 must stay data.
    const sheet = (t: number) =>
      row(1, str('A1', 'id'), str('B1', 'score'), str('C1', 'band'), num('D1', t)) +
      [2, 3, 4].map((r) => row(r, num(`A${r}`, r * 10), num(`B${r}`, r / 10), str(`C${r}`, r % 2 ? 'high' : 'low'), num(`D${r}`, 0))).join('');
    const a = parseXlsx(xlsx([{ name: 'S', data: sheet(0.25) }]));
    expect(a.sheets[0].columns).toEqual(['id', 'score', 'band', 'D']);
    expect(a.sheets[0].origin).toEqual({ row: 0, col: 0 });
    const b = parseXlsx(xlsx([{ name: 'S', data: sheet(0.35) }]));
    const d = diffWorkbooks(a, b).sheets[0].diff;
    expect(d.keyColumns).toEqual(['id']);
    expect(d.stats.cellsChanged).toBe(1); // the parameter edit is visible
  });
});

describe('sheet-scoped defined names', () => {
  // Rate is 0.1 workbook-wide, 0.2 on Inputs, 0.3 on Calc. Report has none of its own.
  const book = (inputsRate: number) =>
    xlsx(
      [
        { name: 'Inputs', data: row(1, num('A1', 100), num('B1', inputsRate)) + row(2, num('A2', 100 * inputsRate, 'A1*Rate')) },
        { name: 'Calc', data: row(1, num('A1', 100), num('B1', 0.3)) + row(2, num('A2', 30, 'A1*Rate')) },
        { name: 'Report', data: row(1, num('A1', 0.1), num('B1', 10, 'A1*100'), num('C1', 100 * inputsRate * 2, 'Inputs!Rate*1000'), num('D1', 0, 'Rate*100')) },
      ],
      { Rate: 'Report!$A$1', 'Rate@0': 'Inputs!$B$1', 'Rate@1': 'Calc!$B$1' },
    );

  it('reads them onto their sheets', () => {
    const wb = parseXlsx(book(0.2), { header: false });
    expect(wb.names).toEqual({ Rate: 'Report!$A$1' });
    expect(wb.sheets[0].names).toEqual({ Rate: 'Inputs!$B$1' });
    expect(wb.sheets[1].names).toEqual({ Rate: 'Calc!$B$1' });
    expect(wb.sheets[2].names).toBeUndefined();
  });

  it('resolves each formula to its own sheet, the workbook, or Sheet!Name', () => {
    const wb = parseXlsx(book(0.2), { header: false });
    const at = (s: number) => scopedNames(wb.names, wb.sheets, wb.sheets[s].name);
    const ref = (f: string, s: number) => formulaRefs(f, at(s)).ranges.map((r) => `${r.sheet}!${r.r0},${r.c0}`);
    expect(ref('=A1*Rate', 0)).toEqual(['null!0,0', 'Inputs!0,1']);
    expect(ref('=A1*Rate', 1)).toEqual(['null!0,0', 'Calc!0,1']);
    expect(ref('=Rate*100', 2)).toEqual(['Report!0,0']);
    expect(ref('=Inputs!Rate*1000', 2)).toEqual(['Inputs!0,1']);
    expect(formulaRefs('=Inputs!Rate*1000', at(2)).dynamic).toBe(false);
  });

  it('computes and traces through them', () => {
    const wb = parseXlsx(book(0.2), { header: false });
    // D1 was saved as 0 (wrong): computing it needs the workbook-level Rate.
    wb.sheets[2].rows[0][3] = null;
    expect(computeMissingValues(wb)).toEqual({ computed: 1, failed: 0 });
    expect(wb.sheets[2].rows[0][3]).toBeCloseTo(10);
    const tr = traceWorkbook(diffWorkbooks(parseXlsx(book(0.2), { header: false }), parseXlsx(book(0.25), { header: false })));
    // Editing Inputs!B1 moves Inputs!A2 (local Rate) and Report!C1 (Inputs!Rate), not Calc or Report!D1.
    expect(tr.stats).toMatchObject({ inputs: 1, computed: 2, unexplained: 0 });
  });
});

describe('merged cells', () => {
  it('reads merged ranges in sheet coordinates, ignoring single cells', () => {
    const wb = parseXlsx(xlsx([{ name: 'S', data: row(1, str('A1', 'Q1 report')) + row(2, str('A2', 'id'), str('B2', 'v'), str('C2', 'w')) + row(3, num('A3', 1), num('B3', 2), num('C3', 3)), merges: ['A1:C1', 'D5:D5', 'A3:A4'] }]), { header: false });
    expect(wb.sheets[0].merges).toEqual([{ r0: 0, c0: 0, r1: 0, c1: 2 }, { r0: 2, c0: 0, r1: 3, c1: 0 }]);
    // Values are unchanged: the merge's value is its top-left cell, the rest empty.
    expect(wb.sheets[0].rows[0]).toEqual(['Q1 report', null, null]);
  });
});
