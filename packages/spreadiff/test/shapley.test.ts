import { describe, expect, it } from 'vitest';
import { diffWorkbooks, traceWorkbook, type CellRef, type Table, type TraceOptions, type WorkbookDiff } from '../src/core';

/**
 * One sheet of named rows: `cells` maps a row name to a number (an input)
 * or `[formula, cached value]`. Rows sit at sheet row 2 onward, in order, so
 * formulas refer to B2, B3, ….
 */
function sheet(cells: [string, number | [string, number]][]): Table {
  return {
    name: 'Model',
    columns: ['name', 'value'],
    origin: { row: 1, col: 0 },
    rows: cells.map(([n, v]) => [n, typeof v === 'number' ? v : v[1]]),
    formulas: cells.map(([, v]) => [null, typeof v === 'number' ? null : v[0]]),
  };
}

function trace(before: Table, after: Table, options: TraceOptions = {}) {
  const diff = diffWorkbooks({ name: 'old.xlsx', sheets: [before] }, { name: 'new.xlsx', sheets: [after] });
  const t = traceWorkbook(diff, options);
  const at = (name: string): CellRef => {
    const d = diff.sheets[0].diff;
    for (let row = 0; row < d.rowCount; row++) if (d.newTable.rows[d.newRows[row]]?.[0] === name) return { sheet: 0, row, col: 1 };
    throw new Error(`no row ${name}`);
  };
  /** A cell's causes by name, each with its contribution, and how they were measured. */
  const shares = (name: string) => {
    const c = t.get(at(name))!;
    return {
      attribution: c.attribution,
      byCause: Object.fromEntries(c.causes.map((r, i) => [nameOf(diff, r), c.contributions?.[i]])),
    };
  };
  return { t, shares };
}

const nameOf = (diff: WorkbookDiff, r: CellRef) => {
  const d = diff.sheets[r.sheet].diff;
  return String(d.newTable.rows[d.newRows[r.row]][0]);
};

const sum = (o: Record<string, number | null | undefined>) => Object.values(o).reduce<number>((s, v) => s + (v ?? 0), 0);

describe('Shapley contributions', () => {
  // price × volume: the textbook case where one-at-a-time parts don't add up.
  const before = sheet([['price', 10], ['volume', 5], ['revenue', ['=B2*B3', 50]]]);
  const after = sheet([['price', 12], ['volume', 6], ['revenue', ['=B2*B3', 72]]]);

  it('but-for stays the default: parts measured one at a time, and they overshoot', () => {
    const { shares } = trace(before, after);
    const r = shares('revenue');
    expect(r.attribution).toBe('but-for');
    // 72 − 10×6 and 72 − 12×5: 24 in all, for a change of 22.
    expect(r.byCause).toEqual({ price: 12, volume: 12 });
  });

  it('splits price × volume evenly, and the parts add up to the change', () => {
    const { shares } = trace(before, after, { contributions: 'shapley' });
    const r = shares('revenue');
    expect(r.attribution).toBe('shapley');
    // price: ½[(12×5 − 50) + (72 − 10×6)] = 11; volume the same by symmetry.
    expect(r.byCause).toEqual({ price: 11, volume: 11 });
    expect(sum(r.byCause)).toBe(72 - 50);
  });

  it('credits an edit that mattered only together with another (MAX of two raised inputs)', () => {
    const b = sheet([['a', 5], ['b', 5], ['top', ['=MAX(B2,B3)', 5]]]);
    const a = sheet([['a', 10], ['b', 8], ['top', ['=MAX(B2,B3)', 10]]]);
    // Undoing b alone leaves MAX at 10, so but-for names only a.
    expect(trace(b, a).shares('top').byCause).toEqual({ a: 2 });
    // f(∅)=5, f(a)=10, f(b)=8, f(ab)=10: a = ½[5 + 2] = 3.5, b = ½[3 + 0] = 1.5.
    const r = trace(b, a, { contributions: 'shapley' }).shares('top');
    expect(r.byCause).toEqual({ a: 3.5, b: 1.5 });
    expect(sum(r.byCause)).toBe(5);
  });

  it('finds causes where no single edit was a but-for cause', () => {
    const b = sheet([['a', 5], ['b', 5], ['top', ['=MAX(B2,B3)', 5]]]);
    const a = sheet([['a', 10], ['b', 10], ['top', ['=MAX(B2,B3)', 10]]]);
    expect(trace(b, a).t.get({ sheet: 0, row: 2, col: 1 })!.method).toBe('references');
    const r = trace(b, a, { contributions: 'shapley' }).shares('top');
    expect(r).toEqual({ attribution: 'shapley', byCause: { a: 2.5, b: 2.5 } });
  });

  it('treats a formula edit as a player: undone, the old formula runs on the inputs of that subset', () => {
    // cost = price × volume, rewritten as price × volume × 2; total = cost + 1 reads it.
    const b = sheet([['price', 10], ['volume', 5], ['cost', ['=B2*B3', 50]], ['total', ['=B4+1', 51]]]);
    const a = sheet([['price', 12], ['volume', 5], ['cost', ['=B2*B3*2', 120]], ['total', ['=B4+1', 121]]]);
    // f(∅)=51, f(price)=61, f(formula)=101, f(both)=121.
    // price = ½[(61 − 51) + (121 − 101)] = 15; the formula edit = ½[(101 − 51) + (121 − 61)] = 55.
    const r = trace(b, a, { contributions: 'shapley' }).shares('total');
    expect(r.byCause).toEqual({ price: 15, cost: 55 });
    expect(sum(r.byCause)).toBe(121 - 51);
  });

  it('three players: matches the Shapley formula computed by hand', () => {
    // out = a × b + c: a 2→3, b 4→5, c 1→2. Old 9, new 17.
    const b = sheet([['a', 2], ['b', 4], ['c', 1], ['out', ['=B2*B3+B4', 9]]]);
    const a = sheet([['a', 3], ['b', 5], ['c', 2], ['out', ['=B2*B3+B4', 17]]]);
    const r = trace(b, a, { contributions: 'shapley' }).shares('out');
    // c is additive: exactly its own change, 1. a and b split the product's change 15 − 8 = 7:
    // a = ½[(3×4 − 8) + (15 − 2×5)] = 4.5, b = ½[(2×5 − 8) + (15 − 3×4)] = 2.5.
    expect(r.byCause).toEqual({ a: 4.5, b: 2.5, c: 1 });
    expect(sum(r.byCause)).toBe(8);
  });

  it('a lone cause gets its but-for contribution, labelled shapley', () => {
    const b = sheet([['a', 2], ['out', ['=B2*10', 20]]]);
    const a = sheet([['a', 3], ['out', ['=B2*10', 30]]]);
    expect(trace(b, a, { contributions: 'shapley' }).shares('out')).toEqual({ attribution: 'shapley', byCause: { a: 10 } });
  });

  it('more than 10 edits reaching a result: it keeps but-for contributions', () => {
    const n = 11;
    const ins = (v: number) => Array.from({ length: n }, (_, i) => [`x${i}`, v] as [string, number]);
    const b = sheet([...ins(1), ['total', [`=SUM(B2:B${n + 1})`, n]]]);
    const a = sheet([...ins(2), ['total', [`=SUM(B2:B${n + 1})`, 2 * n]]]);
    const r = trace(b, a, { contributions: 'shapley' }).shares('total');
    expect(r.attribution).toBe('but-for');
    expect(sum(r.byCause)).toBe(n);
  });

  it('a downstream chain: every result on it adds up', () => {
    const b = sheet([['rate', 0.1], ['base', 100], ['tax', ['=B2*B3', 10]], ['net', ['=B3-B4', 90]], ['ratio', ['=B5/B3', 0.9]]]);
    const a = sheet([['rate', 0.2], ['base', 150], ['tax', ['=B2*B3', 30]], ['net', ['=B3-B4', 120]], ['ratio', ['=B5/B3', 0.8]]]);
    const { shares } = trace(b, a, { contributions: 'shapley' });
    expect(sum(shares('tax').byCause)).toBeCloseTo(20, 12);
    expect(sum(shares('net').byCause)).toBeCloseTo(30, 12);
    expect(sum(shares('ratio').byCause)).toBeCloseTo(-0.1, 12);
    // The ratio doesn't depend on the base (net/base = 1 − rate): all of it is the rate's.
    expect(shares('ratio').byCause.rate).toBeCloseTo(-0.1, 12);
    expect(shares('ratio').byCause.base ?? 0).toBeCloseTo(0, 12);
  });
});
