import { describe, expect, it } from 'vitest';
import ExcelJS from 'exceljs';
import { diffTables, diffWorkbooks, parseXlsx, traceWorkbook, tokenizeFormula, formulaRefs, type CellRef, type Table, type Workbook, type WorkbookDiff } from '../src/core';

/**
 * A small pricing model:
 *   Inputs:  rate (named TaxRate) in B2
 *   Lines:   item, price, qty, total = price*qty
 *   Summary: Subtotal = SUM(Lines!D2:D{n+1}), Tax = B2*TaxRate, Total = B2+B3, Stamp = NOW()
 */
function model(opts: { lines: [string, number, number][]; rate: number; discountLine?: string; stamp: number }): Workbook {
  const lines: Table = {
    name: 'Lines', columns: ['item', 'price', 'qty', 'total'], origin: { row: 1, col: 0 },
    rows: [], formulas: [],
  };
  opts.lines.forEach(([item, price, qty], i) => {
    const r = i + 2;
    const disc = item === opts.discountLine;
    lines.rows.push([item, price, qty, price * qty * (disc ? 0.9 : 1)]);
    lines.formulas!.push([null, null, null, disc ? `=B${r}*C${r}*0.9` : `=B${r}*C${r}`]);
  });
  const subtotal = lines.rows.reduce((s, r) => s + (r[3] as number), 0);
  const last = opts.lines.length + 1;
  const summary: Table = {
    name: 'Summary', columns: ['label', 'value'], origin: { row: 1, col: 0 },
    rows: [['Subtotal', subtotal], ['Tax', subtotal * opts.rate], ['Total', subtotal * (1 + opts.rate)], ['Stamp', opts.stamp]],
    formulas: [[null, `=SUM(Lines!D2:D${last})`], [null, '=B2*TaxRate'], [null, '=B2+B3'], [null, '=NOW()']],
  };
  const inputs: Table = { name: 'Inputs', columns: ['name', 'value'], origin: { row: 1, col: 0 }, rows: [['rate', opts.rate]], formulas: [[null, null]] };
  return { name: 'model.xlsx', sheets: [inputs, lines, summary], names: { TaxRate: 'Inputs!$B$2' } };
}

const base = { lines: [['apple', 2, 10], ['pear', 3, 5], ['fig', 5, 2]] as [string, number, number][], rate: 0.1, stamp: 1 };

function find(d: WorkbookDiff, sheet: string, rowLabel: string, col: string): CellRef {
  const si = d.sheets.findIndex((s) => s.name === sheet);
  const sd = d.sheets[si].diff;
  const ci = sd.columns.findIndex((c) => c.name === col);
  for (let row = 0; row < sd.rowCount; row++) {
    const ni = sd.newRows[row] >= 0 ? sd.newRows[row] : -1;
    const oi = sd.oldRows[row];
    const label = ni >= 0 ? sd.newTable.rows[ni][0] : sd.oldTable.rows[oi][0];
    if (label === rowLabel) return { sheet: si, row, col: ci };
  }
  throw new Error(`no row ${rowLabel} in ${sheet}`);
}
const names = (d: WorkbookDiff, refs: CellRef[]) =>
  refs.map((r) => {
    const sd = d.sheets[r.sheet].diff;
    const ni = sd.newRows[r.row];
    const label = ni >= 0 ? sd.newTable.rows[ni][0] : sd.oldTable.rows[sd.oldRows[r.row]][0];
    return `${d.sheets[r.sheet].name}.${label}.${sd.columns[r.col].name}`;
  }).sort();

describe('formula references', () => {
  it('finds refs, ranges, sheets, names and dynamic/volatile functions', () => {
    const r = formulaRefs("=SUM(Lines!D2:D9)+'Q&A 2026'!$B$3*TaxRate+LOG10(A1)+\"B7 in a string\"+C:C", { TaxRate: 'Inputs!$B$2' });
    expect(r.ranges).toEqual([
      { sheet: 'Lines', r0: 1, c0: 3, r1: 8, c1: 3 },
      { sheet: 'Q&A 2026', r0: 2, c0: 1, r1: 2, c1: 1 },
      { sheet: null, r0: 0, c0: 0, r1: 0, c1: 0 },
      { sheet: null, r0: -1, c0: 2, r1: -1, c1: 2 },
      { sheet: 'Inputs', r0: 1, c0: 1, r1: 1, c1: 1 },
    ]);
    expect(r.dynamic).toBe(false);
    expect(formulaRefs('=INDIRECT("A"&B1)').dynamic).toBe(true);
    expect(formulaRefs('=UnknownName*2').dynamic).toBe(true);
    expect(formulaRefs('=NOW()-A1').volatile).toBe(true);
    expect(formulaRefs('=[1]Other!A1').dynamic).toBe(true);
    expect(tokenizeFormula('=Table1[Col]').dynamic).toBe(true);
  });
});

describe('formula comparison through the alignment', () => {
  it('does not flag formulas that only shifted because a row was inserted above', () => {
    const d = diffWorkbooks(model(base), model({ ...base, lines: [['kiwi', 1, 1], ...base.lines] }));
    const lines = d.sheets.find((s) => s.name === 'Lines')!.diff;
    expect(lines.stats.cellsChanged).toBe(0);
    expect(lines.stats.rowsAdded).toBe(1);
    // SUM(D2:D4) → SUM(D2:D5) grew only by the inserted row: same formula,
    // so the total is a computed change caused by the added row.
    const summary = d.sheets.find((s) => s.name === 'Summary')!.diff;
    expect(summary.formulaChanged.size).toBe(0);
    const t = traceWorkbook(d);
    const sub = t.get(find(d, 'Summary', 'Subtotal', 'value'))!;
    expect(sub.kind).toBe('computed');
    expect(names(d, sub.causes)).toEqual(['Lines.kiwi.total']);
  });
  it('still flags a range that was really edited', () => {
    const v2 = model(base);
    v2.sheets[2].formulas![0][1] = '=SUM(Lines!D2:D3)'; // drops the last line on purpose
    expect(diffWorkbooks(model(base), v2).sheets.find((s) => s.name === 'Summary')!.diff.formulaChanged.size).toBe(1);
  });
  it('single-table diffs canonicalize same-sheet refs too', () => {
    const mk = (rows: [string, number, number][]): Table => ({
      name: 'S', columns: ['item', 'price', 'qty', 'total'], origin: { row: 1, col: 0 },
      rows: rows.map(([n, p, q]) => [n, p, q, p * q]),
      formulas: rows.map((_, i) => [null, null, null, `=B${i + 2}*C${i + 2}`]),
    });
    const d = diffTables(mk([['a', 1, 2], ['b', 3, 4]]), mk([['new', 5, 5], ['a', 1, 2], ['b', 3, 4]]));
    expect(d.stats.cellsChanged).toBe(0);
  });
});

describe('traceWorkbook', () => {
  it('classifies inputs, formula edits and computed changes, and finds root causes', () => {
    const d = diffWorkbooks(model(base), model({ ...base, lines: [['apple', 2.5, 10], ['pear', 3, 5], ['fig', 5, 2]], rate: 0.2, discountLine: 'fig', stamp: 2 }));
    const t = traceWorkbook(d);
    expect(t.get(find(d, 'Lines', 'apple', 'price'))!.kind).toBe('input');
    expect(t.get(find(d, 'Inputs', 'rate', 'value'))!.kind).toBe('input');
    expect(t.get(find(d, 'Lines', 'fig', 'total'))!.kind).toBe('formula');
    const appleTotal = t.get(find(d, 'Lines', 'apple', 'total'))!;
    expect(appleTotal.kind).toBe('computed');
    expect(names(d, appleTotal.causes)).toEqual(['Lines.apple.price']);
    const total = t.get(find(d, 'Summary', 'Total', 'value'))!;
    expect(total.kind).toBe('computed');
    expect(names(d, total.causes)).toEqual(['Inputs.rate.value', 'Lines.apple.price', 'Lines.fig.total']);
    // Impacts go the other way.
    const price = t.get(find(d, 'Lines', 'apple', 'price'))!;
    expect(names(d, price.impacts)).toEqual(['Lines.apple.total', 'Summary.Subtotal.value', 'Summary.Tax.value', 'Summary.Total.value']);
    // The rate only reaches Tax and Total (through the TaxRate name).
    expect(names(d, t.get(find(d, 'Inputs', 'rate', 'value'))!.impacts)).toEqual(['Summary.Tax.value', 'Summary.Total.value']);
    // NOW(): its own cause.
    expect(t.get(find(d, 'Summary', 'Stamp', 'value'))!.unexplained).toBe('volatile');
    // One level up, for cause trees.
    expect(names(d, t.parents(find(d, 'Summary', 'Total', 'value')))).toEqual(['Summary.Subtotal.value', 'Summary.Tax.value']);
    expect(t.stats).toMatchObject({ inputs: 2, formulas: 1 });
  });

  it('traces added and removed rows into the ranges that read them', () => {
    const removed = diffWorkbooks(model(base), model({ ...base, lines: [base.lines[0], base.lines[2]] }));
    const t1 = traceWorkbook(removed);
    expect(names(removed, t1.get(find(removed, 'Summary', 'Subtotal', 'value'))!.causes)).toContain('Lines.pear.total');
    expect(t1.get(find(removed, 'Lines', 'pear', 'total'))!.kind).toBe('removed');
  });

  it('stops at unchanged intermediate results', () => {
    // apple: price 2→4 and qty 10→5: total unchanged, so nothing downstream changed because of them.
    const d = diffWorkbooks(model(base), model({ ...base, lines: [['apple', 4, 5], ['pear', 3, 5], ['fig', 5, 2]] }));
    const t = traceWorkbook(d);
    expect(t.get(find(d, 'Lines', 'apple', 'price'))!.impactCount).toBe(0);
    expect(d.sheets.find((s) => s.name === 'Summary')!.diff.stats.cellsChanged).toBe(0);
  });

  it('works on a real .xlsx written by ExcelJS (defined names, cached values)', async () => {
    const write = async (price: number, rate: number) => {
      const wb = new ExcelJS.Workbook();
      const inputs = wb.addWorksheet('Inputs');
      inputs.addRow(['name', 'value']);
      inputs.addRow(['rate', rate]);
      wb.definedNames.add('Inputs!$B$2', 'TaxRate');
      const lines = wb.addWorksheet('Lines');
      lines.addRow(['item', 'price', 'qty', 'total', 'taxed']);
      lines.addRow(['apple', price, 10, { formula: 'B2*C2', result: price * 10 }, { formula: 'D2*(1+TaxRate)', result: price * 10 * (1 + rate) }]);
      return parseXlsx(new Uint8Array((await wb.xlsx.writeBuffer()) as ArrayBuffer));
    };
    const a = await write(2, 0.1);
    const b = await write(3, 0.1);
    expect(b.names).toEqual({ TaxRate: 'Inputs!$B$2' });
    const d = diffWorkbooks(a, b);
    const t = traceWorkbook(d);
    const taxed = t.get(find(d, 'Lines', 'apple', 'taxed'))!;
    expect(taxed.kind).toBe('computed');
    expect(names(d, taxed.causes)).toEqual(['Lines.apple.price']);
  });

  it('stays fast on a long running-total chain', () => {
    // C[i] = C[i-1] + B[i]: changing B2 ripples through 20k cells.
    const n = 20_000;
    const mk = (first: number): Workbook => {
      const rows: (string | number)[][] = [];
      const formulas: (string | null)[][] = [];
      let acc = 0;
      for (let i = 0; i < n; i++) {
        const b = i === 0 ? first : 1;
        acc += b;
        rows.push([i, b, acc]);
        formulas.push([null, null, i === 0 ? '=B2' : `=C${i + 1}+B${i + 2}`]);
      }
      return { name: 'w', sheets: [{ name: 'S', columns: ['id', 'b', 'c'], origin: { row: 1, col: 0 }, rows, formulas }] };
    };
    const t0 = performance.now();
    const d = diffWorkbooks(mk(1), mk(2));
    const t = traceWorkbook(d);
    expect(performance.now() - t0).toBeLessThan(3000);
    expect(t.stats.computed).toBe(n);
    expect(t.get({ sheet: 0, row: n - 1, col: 2 })!.causeCount).toBe(1);
  });
});
