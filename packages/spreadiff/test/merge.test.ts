import { describe, expect, it } from 'vitest';
import { conflictAnnotations, mergeTables, parseCsv, type CellValue, type Table } from '../src/core';

const csv = (s: string) => parseCsv(s.replace(/^\s+/gm, ''));

describe('mergeTables', () => {
  it('takes each side’s changes, rows added and removed on either side', () => {
    const base = csv(`id,price,qty
      1,10,5
      2,20,1
      3,30,7`);
    const ours = csv(`id,price,qty
      1,12,5
      2,20,1
      3,30,7
      4,40,4`);
    const theirs = csv(`id,price,qty
      1,10,6
      3,30,7`);
    const m = mergeTables(base, ours, theirs);
    expect(m.conflicts).toEqual([]);
    expect(m.table.rows).toEqual([[1, 12, 6], [3, 30, 7], [4, 40, 4]]);
    expect(m.stats).toEqual({ ours: 2, theirs: 2, both: 0 });
    expect(m.diff.stats.rowsAdded).toBe(1);
    expect(m.diff.stats.rowsRemoved).toBe(1);
  });

  it('reports different changes to one cell, taking ours unless told otherwise', () => {
    const base = csv('id,price\n1,10');
    const ours = csv('id,price\n1,12');
    const theirs = csv('id,price\n1,15');
    const m = mergeTables(base, ours, theirs);
    expect(m.table.rows).toEqual([[1, 12]]);
    expect(m.conflicts).toEqual([{ kind: 'cell', row: 0, baseRow: 0, column: 'price', base: 10, ours: 12, theirs: 15, took: 'ours' }]);
    expect(conflictAnnotations(m)).toEqual([{ row: 0, column: 'price', data: m.conflicts[0], id: 0 }]);
    expect(mergeTables(base, ours, theirs, { prefer: 'theirs' }).table.rows).toEqual([[1, 15]]);
    // The same change on both sides isn't a conflict.
    expect(mergeTables(base, ours, ours).conflicts).toEqual([]);
  });

  it('flags a row one side deleted and the other edited', () => {
    const base = csv('id,price\n1,10\n2,20');
    const ours = csv('id,price\n1,10');
    const theirs = csv('id,price\n1,10\n2,25');
    const m = mergeTables(base, ours, theirs);
    expect(m.conflicts).toEqual([{ kind: 'delete-edit', row: -1, baseRow: 1, oursRow: null, theirsRow: [2, 25], took: 'ours' }]);
    expect(m.table.rows).toEqual([[1, 10]]);
    // The left-out row shows as removed in the merged diff, with its annotation on it.
    const [a] = conflictAnnotations(m);
    expect(m.diff.status[a.row]).toBe(2);
    const t = mergeTables(base, ours, theirs, { prefer: 'theirs' });
    expect(t.table.rows).toEqual([[1, 10], [2, 25]]);
    expect(t.conflicts[0].row).toBe(1);
  });

  it('merges a row both sides added with the same key, or flags it', () => {
    const base = csv('id,v\n1,a');
    expect(mergeTables(base, csv('id,v\n1,a\n2,b'), csv('id,v\n1,a\n2,b')).table.rows).toEqual([[1, 'a'], [2, 'b']]);
    const m = mergeTables(base, csv('id,v\n1,a\n2,b'), csv('id,v\n1,a\n2,c'), { keys: ['id'] });
    expect(m.conflicts.map((c) => c.kind)).toEqual(['both-added']);
    expect(m.table.rows).toEqual([[1, 'a'], [2, 'b']]);
    // Without keys there's no telling they're the same row: both stay.
    expect(mergeTables(base, csv('id,v\n1,a\n2,b'), csv('id,v\n1,a\n2,c'), { keys: 'none' }).table.rows).toEqual([[1, 'a'], [2, 'b'], [2, 'c']]);
  });

  it('carries column changes from either side', () => {
    const base = csv('id,price,note\n1,10,x');
    const ours = csv('id,price,note,owner\n1,10,x,me');
    const theirs = csv('id,price\n1,11');
    const m = mergeTables(base, ours, theirs);
    expect(m.table.columns).toEqual(['id', 'price', 'owner']);
    expect(m.table.rows).toEqual([[1, 11, 'me']]);
  });

  it('equals applying both sides’ edits when they touch different cells (random tables)', () => {
    let seed = 42;
    const rnd = () => ((seed = (seed * 1103515245 + 12345) % 2 ** 31) / 2 ** 31);
    for (let trial = 0; trial < 200; trial++) {
      const n = 3 + Math.floor(rnd() * 12);
      const rows: CellValue[][] = Array.from({ length: n }, (_, i) => [i + 1, Math.floor(rnd() * 100), Math.floor(rnd() * 100)]);
      const table = (rs: CellValue[][]): Table => ({ name: 'T', columns: ['id', 'a', 'b'], rows: rs.map((r) => r.slice()) });
      const expected = rows.map((r) => r.slice());
      const ours = rows.map((r) => r.slice());
      const theirs = rows.map((r) => r.slice());
      for (let i = 0; i < n; i++) {
        const x = rnd();
        // Ours edits column a, theirs column b: never the same cell.
        if (x < 0.3) expected[i][1] = ours[i][1] = 1000 + i;
        else if (x < 0.6) expected[i][2] = theirs[i][2] = 2000 + i;
      }
      const m = mergeTables(table(rows), table(ours), table(theirs));
      expect(m.conflicts).toEqual([]);
      expect(m.table.rows).toEqual(expected);
    }
  });
});
