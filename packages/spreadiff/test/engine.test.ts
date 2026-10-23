import { describe, expect, it } from 'vitest';
import { computeMissingValues, diffWorkbooks, traceWorkbook, type CellValue, type Table, type Workbook } from '../src/core';
import { Engine } from '../src/core/engine/evaluate';
import { tableHost } from '../src/core/engine/host';
import { isError, type Scalar } from '../src/core/engine/values';

/** A sheet from rows of values or `=formulas` (value null until computed). Row 1 is data. */
function sheet(name: string, rows: (CellValue | string)[][], values?: CellValue[][]): Table {
  const formulas = rows.map((r) => r.map((v) => (typeof v === 'string' && v.startsWith('=') ? v : null)));
  const data = rows.map((r, i) => r.map((v, j) => (typeof v === 'string' && v.startsWith('=') ? values?.[i]?.[j] ?? null : (v as CellValue))));
  const width = Math.max(...rows.map((r) => r.length));
  return { name, columns: Array.from({ length: width }, (_, i) => String.fromCharCode(65 + i)), rows: data.map((r) => [...r, ...Array(width - r.length).fill(null)]), formulas, origin: { row: 0, col: 0 } };
}

/** Evaluate one formula against a sheet. */
function calc(formula: string, rows: (CellValue | string)[][] = [], at = { row: 50, col: 10 }): Scalar {
  const wb: Workbook = { name: 'w', sheets: [sheet('S', rows)] };
  computeMissingValues(wb);
  const e = new Engine(tableHost(wb.sheets));
  return e.evaluate(formula, 'S', at.row, at.col).value;
}
const err = (v: Scalar) => (isError(v) ? v.code : v);

describe('engine: operators and coercion', () => {
  it('follows Excel precedence', () => {
    expect(calc('=-2^2')).toBe(4);
    expect(calc('=2^3^2')).toBe(64);
    expect(calc('=1+2*3-4/2')).toBe(5);
    expect(calc('=1&2+3')).toBe('15');
    expect(calc('=50%*2')).toBe(1);
    expect(calc('=1<2=TRUE')).toBe(true);
  });
  it('coerces like Excel', () => {
    expect(calc('="3"+1')).toBe(4);
    expect(calc('=TRUE+1')).toBe(2);
    expect(calc('=A1+1', [[null]])).toBe(1);
    expect(err(calc('="x"+1'))).toBe('#VALUE!');
    expect(calc('="a"="A"')).toBe(true);
    expect(calc('=1&""')).toBe('1');
    expect(calc('=1/3&""')).toBe('0.333333333333333');
    expect(err(calc('=1/0'))).toBe('#DIV/0!');
  });
});

describe('engine: functions', () => {
  const data = [
    ['apple', 'fruit', 3],
    ['kale', 'veg', 5],
    ['pear', 'fruit', 7],
    ['leek', 'veg', 11],
  ];
  it('aggregates, and conditional aggregates with wildcards and operators', () => {
    expect(calc('=SUM(C1:C4)', data)).toBe(26);
    expect(calc('=SUMIF(B1:B4,"fruit",C1:C4)', data)).toBe(10);
    expect(calc('=SUMIFS(C1:C4,B1:B4,"veg",C1:C4,">6")', data)).toBe(11);
    expect(calc('=COUNTIF(A1:A4,"*e*")', data)).toBe(4);
    expect(calc('=COUNTIF(A1:A4,"?e??")', data)).toBe(2);
    expect(calc('=AVERAGEIFS(C1:C4,B1:B4,"<>veg")', data)).toBe(5);
    expect(calc('=SUMPRODUCT((B1:B4="veg")*C1:C4)', data)).toBe(16);
  });
  it('looks up exactly and approximately', () => {
    expect(calc('=VLOOKUP("pear",A1:C4,3,FALSE)', data)).toBe(7);
    expect(err(calc('=VLOOKUP("fig",A1:C4,3,FALSE)', data))).toBe('#N/A');
    expect(calc('=VLOOKUP(6,C1:C4,1,TRUE)', data)).toBe(5);
    expect(calc('=INDEX(A1:A4,MATCH(11,C1:C4,0))', data)).toBe('leek');
    expect(calc('=XLOOKUP("kale",A1:A4,C1:C4)', data)).toBe(5);
    expect(calc('=XLOOKUP("fig",A1:A4,C1:C4,"none")', data)).toBe('none');
  });
  it('rounds the decimal Excel shows', () => {
    expect(calc('=ROUND(2.675,2)')).toBe(2.68);
    expect(calc('=ROUND(-2.5,0)')).toBe(-3);
    expect(calc('=ROUNDUP(1.21,1)')).toBe(1.3);
    expect(calc('=MROUND(10,3)')).toBe(9);
  });
  it('does loan and date maths', () => {
    expect(calc('=PMT(0.05/12,360,300000)')).toBeCloseTo(-1610.46, 2);
    expect(calc('=IPMT(0.05/12,1,360,300000)')).toBeCloseTo(-1250, 6);
    expect(calc('=DATE(2026,2,30)')).toBe(calc('=DATE(2026,3,2)'));
    expect(calc('=EOMONTH(DATE(2024,1,15),1)')).toBe(calc('=DATE(2024,2,29)'));
    expect(calc('=YEAR(DATE(1900,3,1))')).toBe(1900);
    expect(calc('=TEXT(1234.5,"#,##0.00")')).toBe('1,234.50');
    expect(calc('=TEXT(DATE(2026,7,4),"yyyy-mm-dd")')).toBe('2026-07-04');
  });
  it('evaluates OFFSET and INDIRECT, which static analysis can\'t follow', () => {
    expect(calc('=SUM(OFFSET(C1,1,0,2,1))', data)).toBe(12);
    expect(calc('=INDIRECT("C"&4)', data)).toBe(11);
  });
});

describe('engine: what a formula actually reads', () => {
  const e = () => {
    const wb: Workbook = { name: 'w', sheets: [sheet('S', [['a', 1, 10], ['b', 2, 20], ['c', 3, 30]])] };
    return new Engine(tableHost(wb.sheets));
  };
  it('a lookup reads its key column and the one cell it returns', () => {
    const { value, reads } = e().evaluate('=VLOOKUP("b",A1:C3,3,FALSE)', 'S', 9, 9);
    expect(value).toBe(20);
    expect(reads).toEqual([
      { sheet: 'S', r0: 0, c0: 0, r1: 2, c1: 0 },
      { sheet: 'S', r0: 1, c0: 2, r1: 1, c1: 2 },
    ]);
  });
  it('an IF reads only the branch it takes', () => {
    const { reads } = e().evaluate('=IF(B1>0,C1,C3)', 'S', 9, 9);
    expect(reads.map((r) => `${r.r0},${r.c0}`)).toEqual(['0,1', '0,2']);
  });
  it('copies of a formula share one parse, shifted to each cell (3-D references too)', () => {
    const wb: Workbook = { name: 'w', sheets: [sheet('A', [[1, 2]]), sheet('B', [[10, 20]])] };
    const eng = new Engine(tableHost(wb.sheets));
    expect(eng.evaluate('=SUM(A:B!A1)', 'A', 5, 0).value).toBe(11);
    expect(eng.evaluate('=SUM(A:B!B1)', 'A', 5, 1).value).toBe(22); // same shape, shifted one column
    expect(eng.compile('=A1*2', 1, 0).shape).toBe(eng.compile('=A5*2', 5, 0).shape);
  });
});

describe('computing values a file lacks', () => {
  it('computes formulas without cached values, in dependency order', () => {
    const wb: Workbook = { name: 'w', sheets: [sheet('S', [['=A2*2'], ['=A3+1'], [4]])] };
    expect(computeMissingValues(wb)).toEqual({ computed: 2, failed: 0 });
    expect(wb.sheets[0].rows.map((r) => r[0])).toEqual([10, 5, 4]);
  });
  it('leaves what it can\'t compute empty, and counts it', () => {
    // NOW() can't be computed, so neither can what reads it; the unrelated A3 can.
    const wb: Workbook = { name: 'w', sheets: [sheet('S', [['=NOW()'], ['=A1+1'], ['=2*3']])] };
    expect(computeMissingValues(wb)).toEqual({ computed: 1, failed: 2 });
    expect(wb.sheets[0].rows.map((r) => r[0])).toEqual([null, null, 6]);
  });
});

describe('trace: recalculated causes', () => {
  /** A one-sheet model: [old rows, new rows], with values computed as a calculating app would. */
  function model(oldRows: (CellValue | string)[][], newRows: (CellValue | string)[][], tweak?: (w: Workbook) => void) {
    const a: Workbook = { name: 'a', sheets: [sheet('S', oldRows)] };
    const b: Workbook = { name: 'b', sheets: [sheet('S', newRows)] };
    computeMissingValues(a);
    computeMissingValues(b);
    tweak?.(b);
    const d = diffWorkbooks(a, b, { keys: 'none' });
    return { d, t: traceWorkbook(d) };
  }
  const at = (row: number, col: number) => ({ sheet: 0, row, col });

  it('an IF that switched branch is caused by what it now reads, not the other branch', () => {
    // C1 = IF(A1>=100, B2, B3): the gold rate (B2) vs the silver rate (B3).
    const rows = (sales: number, silver: number) => [[sales, 0, '=A1*IF(A1>=100,B2,B3)'], [0, 0.1], [0, silver]];
    const { t } = model(rows(90, 0.05), rows(120, 0.07));
    const c = t.get(at(0, 2))!;
    expect(c.method).toBe('recalculated');
    expect(c.causes).toEqual([at(0, 0)]); // sales, not the (also edited) silver rate
    // References alone would blame both.
    const refsOnly = traceWorkbook(model(rows(90, 0.05), rows(120, 0.07)).d, { recalculate: false }).get(at(0, 2))!;
    expect(refsOnly.causes).toHaveLength(2);
  });

  it('a clamp that absorbed an edit is not caused by it', () => {
    // C1 = MAX(0, A1 - B1): B1 rose, but A1 fell below it either way.
    const { t } = model([[10, 5, '=MAX(0,A1-B1)']], [[3, 6, '=MAX(0,A1-B1)']]);
    expect(t.get(at(0, 2))!.causes).toEqual([at(0, 0)]);
  });

  it('reports each cause\'s contribution', () => {
    const { t } = model([[10, 2, '=A1*B1']], [[12, 3, '=A1*B1']]);
    const c = t.get(at(0, 2))!;
    expect(c.causes).toEqual([at(0, 0), at(0, 1)]);
    expect(c.contributions).toEqual([36 - 30, 36 - 24]); // without the units edit 3*10=30; without the price edit 12*2=24
  });

  it('finds a stale cached value: the input changed, the result was saved unrecalculated', () => {
    const { t } = model([[10, '=A1*2']], [[11, '=A1*2']], (w) => {
      w.sheets[0].rows[0][1] = 20; // still the old result: saved without recalculating
    });
    const c = t.get(at(0, 1))!;
    expect(c.kind).toBe('computed');
    expect(c.stale).toEqual({ cached: 20, recalculated: 22 });
    expect(c.causes).toEqual([at(0, 0)]);
    expect(t.stats.stale).toBe(1);
  });

  it('falls back to references when it can\'t reproduce the old file', () => {
    // The old cached value disagrees with the formula (another engine, or a function it lacks).
    const a: Workbook = { name: 'a', sheets: [sheet('S', [[10, '=A1*2']], [[null, 999]])] };
    const b: Workbook = { name: 'b', sheets: [sheet('S', [[11, '=A1*2']], [[null, 1000]])] };
    const t = traceWorkbook(diffWorkbooks(a, b, { keys: 'none' }));
    const c = t.get(at(0, 1))!;
    expect(c.method).toBe('references');
    expect(c.stale).toBeUndefined(); // no stale claim without calibration
    expect(c.causes).toEqual([at(0, 0)]);
    expect(t.stats.unmodelled).toBe(1);
  });
});

describe('applyRecalculation', () => {
  it('shows stale cells with their recalculated values, as changes', async () => {
    const { applyRecalculation } = await import('../src/core');
    const a: Workbook = { name: 'a', sheets: [sheet('S', [[10, '=A1*2'], ['x', 1]])] };
    const b: Workbook = { name: 'b', sheets: [sheet('S', [[11, '=A1*2'], ['x', 1]])] };
    computeMissingValues(a);
    computeMissingValues(b);
    b.sheets[0].rows[0][1] = 20; // saved without recalculating
    const d = diffWorkbooks(a, b, { keys: 'none' });
    const t = traceWorkbook(d);
    const r = applyRecalculation(d, t);
    expect(r.sheets[0].diff.newTable.rows[0][1]).toBe(22);
    expect(r.sheets[0].diff.changedCells.get(0)![1]).toBe(1);
    expect(r.sheets[0].diff.stats.cellsChanged).toBe(2);
    // The original diff is untouched.
    expect(d.sheets[0].diff.newTable.rows[0][1]).toBe(20);
    expect(d.sheets[0].diff.columns[1].changes).toBe(0);
  });
});

describe('numericDelta', () => {
  it('rounds away binary noise at the operands\' precision', async () => {
    const { numericDelta } = await import('../src/core');
    expect(numericDelta(3887856.67, 3878764.11)!.delta).toBe(-9092.56);
    expect(numericDelta(0.1, 0.3)!.delta).toBe(0.2);
    expect(numericDelta(1e-20, 2e-20)!.delta).toBe(1e-20);
  });
});
