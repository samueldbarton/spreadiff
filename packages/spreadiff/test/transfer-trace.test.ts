import { describe, expect, it } from 'vitest';
import {
  computeMissingValues,
  diffWorkbooks,
  packWorkbookDiff,
  packWorkbookTrace,
  traceWorkbook,
  unpackWorkbookDiff,
  unpackWorkbookTrace,
  type CellRef,
  type Table,
  type Workbook,
  type WorkbookTrace,
} from '../src/core';

/** A sheet from rows of values or `=formulas`. Row 1 is data. */
function sheet(name: string, rows: (string | number | null)[][], values?: (string | number | null)[][]): Table {
  const formulas = rows.map((r) => r.map((v) => (typeof v === 'string' && v.startsWith('=') ? v : null)));
  const data = rows.map((r, i) => r.map((v, j) => (typeof v === 'string' && v.startsWith('=') ? values?.[i]?.[j] ?? null : v)));
  const width = Math.max(...rows.map((r) => r.length));
  return { name, columns: Array.from({ length: width }, (_, i) => String.fromCharCode(65 + i)), rows: data.map((r) => [...r, ...Array(width - r.length).fill(null)]), formulas, origin: { row: 0, col: 0 } };
}

/** Every traced cell, as plain data, for comparing two traces. */
const dump = (t: WorkbookTrace, sheets: number) => {
  const out: unknown[] = [];
  for (let s = 0; s < sheets; s++) {
    for (const [k, v] of [...t.sheet(s)].sort((a, b) => a[0] - b[0])) out.push([s, k, v]);
  }
  return { cells: out, stats: t.stats, truncated: t.truncated };
};

describe('a trace packed in the worker and unpacked on the page', () => {
  it('is the same trace: kinds, causes, contributions, stale and outside-data inputs', () => {
    // A Bloomberg price, an edited input, a formula edit, an added row and a stale result.
    const book = (price: number, qty: number, extra: boolean, stale = false): Workbook => {
      const rows: (string | number | null)[][] = [
        ['=_xll.BDP("IBM US Equity","PX_LAST")', qty, '=A1*B1'],
        [2, 3, '=A2*B2'],
        ...(extra ? [[4, 5, '=A3*B3']] : []),
        [null, null, extra ? '=SUM(C1:C3)' : '=SUM(C1:C2)'],
      ];
      const wb: Workbook = { name: 'w', sheets: [sheet('S', rows, [[price]])] };
      computeMissingValues(wb);
      if (stale) wb.sheets[0].rows[1][2] = 999; // saved without recalculating
      return wb;
    };
    const d = diffWorkbooks(book(100, 10, false), book(110, 10, true, true), { keys: 'none' }); // the price row keeps its quantity, so it pairs
    const direct = traceWorkbook(d);
    expect(direct.stats.external).toBe(1);
    expect(direct.stats.stale).toBe(1);

    // What the worker does, then what the page does.
    const { packed: pd } = packWorkbookDiff(d);
    const { packed: pt } = packWorkbookTrace(traceWorkbook(d), d);
    const d2 = unpackWorkbookDiff(structuredClone(pd));
    const t2 = unpackWorkbookTrace(structuredClone(pt), d2);
    expect(dump(t2, d2.sheets.length)).toEqual(dump(direct, d.sheets.length));
    // parents() on the page matches for computed and formula-edit cells: the ones it's packed
    // for, and the only ones the grid, the viewer and spreadiff-lite ask about.
    const refs: CellRef[] = [];
    const w = d.sheets[0].diff.columns.length;
    for (const [k, t] of direct.sheet(0)) if (t.kind === 'computed' || t.kind === 'formula') refs.push({ sheet: 0, row: Math.floor(k / w), col: k % w });
    expect(refs.length).toBeGreaterThan(2);
    for (const r of refs) expect({ r, kind: direct.get(r)?.kind, p: t2.parents(r) }).toEqual({ r, kind: direct.get(r)?.kind, p: direct.parents(r) });
  });
});
