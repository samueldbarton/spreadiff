import { describe, expect, it } from 'vitest';
import { diffTables, parseCsv, type Table } from '../src/core';

describe('review 2', () => {
  it('keeps text after a closing quote instead of silently dropping it', () => {
    // Not RFC 4180, but common in hand-edited files; Excel keeps the text.
    const t = parseCsv('a,b\n"ab"cd,e\n', { numbers: false });
    expect(t.rows).toEqual([['abcd', 'e']]);
  });

  it('honours Excel\'s "sep=" hint line', () => {
    const t = parseCsv('sep=;\na;b\n1,5;2\n', { numbers: false });
    expect(t.columns).toEqual(['a', 'b']);
    expect(t.rows).toEqual([['1,5', '2']]);
  });

  it('does not conflate rows whose cells contain the internal separator', () => {
    // ["x\u0001", "y"] and ["x", "\u0001y"] must not intern to the same row id.
    const a: Table = { name: 's', columns: ['p', 'q'], rows: [['x\u0001', 'y']] };
    const b: Table = { name: 's', columns: ['p', 'q'], rows: [['x', '\u0001y'], ['x\u0001', 'y']] };
    const d = diffTables(a, b, { keys: 'none' });
    // The identical row must be matched as equal, the other one added.
    const equalPairs = Array.from(d.status).filter((s) => s === 0).length;
    expect(equalPairs).toBe(1);
    expect(d.stats.rowsAdded).toBe(1);
  });

  it('does not conflate composite keys that contain the separator', () => {
    const a: Table = { name: 's', columns: ['k1', 'k2', 'v'], rows: [['a\u0001', 'b', '1'], ['a', '\u0001b', '2']] };
    const b: Table = { name: 's', columns: ['k1', 'k2', 'v'], rows: [['a', '\u0001b', '2'], ['a\u0001', 'b', '1']] };
    const d = diffTables(a, b, { keys: ['k1', 'k2'] });
    expect(d.stats.cellsChanged).toBe(0);
  });
});

describe('review 2: rename detection scale', () => {
  const wide = (W: number, names: (c: number) => string, order: (c: number) => number, rows = 3000): Table => ({
    name: 's',
    columns: Array.from({ length: W }, (_, c) => names(order(c))),
    rows: Array.from({ length: rows }, (_, i) => Array.from({ length: W }, (_, c) => `${i}-${order(c)}`)),
  });
  it('stays fast when every header of a wide sheet changed', () => {
    const t0 = performance.now();
    const d = diffTables(wide(400, (c) => `old${c}`, (c) => c), wide(400, (c) => `new${c}`, (c) => c), { keys: 'none' });
    expect(performance.now() - t0).toBeLessThan(2000);
    expect(d.stats.columnsRenamed).toBe(400);
  });
  it('still finds renamed columns that were also reordered on a wide sheet', () => {
    const W = 120; // > exhaustive threshold; relies on content fingerprints
    const d = diffTables(wide(W, (c) => `old${c}`, (c) => c, 300), wide(W, (c) => `new${c}`, (c) => W - 1 - c, 300), { keys: 'none' });
    expect(d.stats.columnsRenamed).toBe(W);
    for (const c of d.columns) expect(c.oldName!.slice(3)).toBe(c.name.slice(3));
  });
});

describe('review 2: duplicate keys scale', () => {
  it('matches 300k rows on a low-cardinality key in linear time', () => {
    const t: Table = { name: 's', columns: ['cat', 'v'], rows: Array.from({ length: 300_000 }, (_, i) => [i % 3 ? 'a' : 'b', String(i)]) };
    const t0 = performance.now();
    const d = diffTables(t, t, { keys: ['cat'] });
    expect(performance.now() - t0).toBeLessThan(3000);
    expect(d.stats.rowsEqual).toBe(300_000);
  });
});
