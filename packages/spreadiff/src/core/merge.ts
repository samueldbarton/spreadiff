/**
 * Three-way merge of tables: two versions edited from a common base. Each
 * is diffed against the base; a change on one side is taken; the same
 * change on both is taken once; different changes to the same cell, or an
 * edit to a row the other side deleted, are conflicts. Like `git merge`,
 * but by rows and columns rather than lines.
 */
import { cellsEqual, diffTables } from './diff.js';
import { ROW_STATUS, type CellValue, type DiffOptions, type Table, type TableDiff } from './types.js';

/** A change both sides made differently, from {@link mergeTables}. */
export interface MergeConflict {
  /**
   * - `cell`: both changed the cell, differently;
   * - `delete-edit`: one side deleted the row, the other edited it;
   * - `both-added`: both added a row with the same key, differently.
   */
  kind: 'cell' | 'delete-edit' | 'both-added';
  /** Row index in the merged table, or -1 for a deleted row the merge left out. */
  row: number;
  /** Row index in the base, for `cell` and `delete-edit` conflicts. */
  baseRow?: number;
  /** Column name, for `cell` conflicts. */
  column?: string;
  /** For `cell`: the base value. */
  base?: CellValue;
  /** For `cell`: our value. */
  ours?: CellValue;
  /** For `cell`: their value. */
  theirs?: CellValue;
  /** For row conflicts: our row (in merged columns), or null where we deleted it. */
  oursRow?: CellValue[] | null;
  /** For row conflicts: their row, or null where they deleted it. */
  theirsRow?: CellValue[] | null;
  /** Which side the merged table took. */
  took: 'ours' | 'theirs';
}

/** The result of {@link mergeTables}. */
export interface MergeResult {
  /** The merged table. Conflicts take `options.prefer` (ours by default). */
  table: Table;
  /** Every conflict, in merged-row order. */
  conflicts: MergeConflict[];
  /** How many changes came from each side without conflict. */
  stats: { ours: number; theirs: number; both: number };
  /** The merged table against the base, to show in a grid (`DiffGrid`). */
  diff: TableDiff;
}

/** Options for {@link mergeTables}. */
export interface MergeOptions extends DiffOptions {
  /** Which side a conflict takes. Default `ours`. */
  prefer?: 'ours' | 'theirs';
}

/** For each base row, the aligned row of `d` holding it, or -1 (removed). */
function baseRows(d: TableDiff): Int32Array {
  const out = new Int32Array(d.oldTable.rows.length).fill(-1);
  for (let a = 0; a < d.rowCount; a++) if (d.oldRows[a] >= 0) out[d.oldRows[a]] = a;
  return out;
}

/** Rows a side added, grouped by the base row they follow (-1: before the first). */
function insertions(d: TableDiff): Map<number, number[]> {
  const out = new Map<number, number[]>();
  let lastBase = -1;
  for (let a = 0; a < d.rowCount; a++) {
    if (d.oldRows[a] >= 0) lastBase = d.oldRows[a];
    else if (d.status[a] === ROW_STATUS.added) {
      const list = out.get(lastBase) ?? [];
      list.push(d.newRows[a]);
      out.set(lastBase, list);
    }
  }
  return out;
}

/**
 * Merge two tables edited from the same base. Rows pair by key when the
 * tables have one (`keys`, or found automatically), else by position.
 *
 * ```ts
 * import { mergeTables, parseCsv } from 'spreadiff';
 *
 * const base = parseCsv('id,price,qty\n1,10,5\n2,20,1');
 * const ours = parseCsv('id,price,qty\n1,12,5\n2,20,1');
 * const theirs = parseCsv('id,price,qty\n1,10,6\n2,20,1\n3,30,2');
 * const { table, conflicts } = mergeTables(base, ours, theirs);
 * table.rows; // [[1, 12, 6], [2, 20, 1], [3, 30, 2]]
 * conflicts.length; // 0
 * ```
 */
export function mergeTables(base: Table, ours: Table, theirs: Table, options: MergeOptions = {}): MergeResult {
  const prefer = options.prefer ?? 'ours';
  const d1 = diffTables(base, ours, options);
  const d2 = diffTables(base, theirs, options);
  const eq = (a: CellValue, b: CellValue) => cellsEqual(a, b, options);

  // Columns: the base's, minus those a side removed, then each side's additions.
  const removed = new Set<string>();
  for (const d of [d1, d2]) for (const c of d.columns) if (c.status === 'removed') removed.add(c.oldName ?? c.name);
  type Col = { name: string; base: number; ours: number; theirs: number };
  const cols: Col[] = [];
  const oursOf = new Map(d1.columns.filter((c) => c.oldIndex >= 0).map((c) => [c.oldIndex, c.newIndex]));
  const theirsOf = new Map(d2.columns.filter((c) => c.oldIndex >= 0).map((c) => [c.oldIndex, c.newIndex]));
  base.columns.forEach((name, i) => {
    if (removed.has(name)) return;
    // A rename on one side carries over.
    const r1 = d1.columns.find((c) => c.oldIndex === i && c.status === 'renamed')?.name;
    const r2 = d2.columns.find((c) => c.oldIndex === i && c.status === 'renamed')?.name;
    cols.push({ name: r1 ?? r2 ?? name, base: i, ours: oursOf.get(i) ?? -1, theirs: theirsOf.get(i) ?? -1 });
  });
  for (const c of d1.columns) if (c.status === 'added') cols.push({ name: c.name, base: -1, ours: c.newIndex, theirs: theirs.columns.indexOf(c.name) });
  for (const c of d2.columns) if (c.status === 'added' && !cols.some((x) => x.name === c.name)) cols.push({ name: c.name, base: -1, ours: -1, theirs: c.newIndex });

  const b1 = baseRows(d1);
  const b2 = baseRows(d2);
  const ins1 = insertions(d1);
  const ins2 = insertions(d2);
  const rows: CellValue[][] = [];
  const formulas: (string | null)[][] = [];
  const hasFormulas = !!(base.formulas || ours.formulas || theirs.formulas);
  const conflicts: MergeConflict[] = [];
  const stats = { ours: 0, theirs: 0, both: 0 };
  const get = (t: Table, r: number, c: number): [CellValue, string | null] => (r >= 0 && c >= 0 ? [t.rows[r][c] ?? null, t.formulas?.[r]?.[c] ?? null] : [null, null]);
  const same = (x: [CellValue, string | null], y: [CellValue, string | null]) => eq(x[0], y[0]) && (x[1] ?? null) === (y[1] ?? null);
  const push = (cells: [CellValue, string | null][]) => {
    rows.push(cells.map((c) => c[0]));
    formulas.push(cells.map((c) => c[1]));
  };
  const rowOf = (t: Table, r: number, side: 'ours' | 'theirs') => cols.map((c) => get(t, r, side === 'ours' ? c.ours : c.theirs));
  const emitAdded = (k: number) => {
    const a = ins1.get(k) ?? [];
    const b = ins2.get(k) ?? [];
    // With keys, the same new row on both sides is one row (or a conflict when they differ).
    // Without keys, a row both added at the same place is the same row when every cell matches.
    const keyed = d1.keyColumns.length > 0 && d1.keyColumns.join() === d2.keyColumns.join();
    const keyOf = (t: Table, r: number, side: 'ours' | 'theirs') =>
      keyed
        ? d1.keyColumns.map((k2) => String(t.rows[r][t.columns.indexOf(k2)] ?? '')).join('\u0000')
        : JSON.stringify(rowOf(t, r, side).map((x) => x[0]));
    const theirsByKey = new Map(b.map((r) => [keyOf(theirs, r, 'theirs'), r]));
    const used = new Set<number>();
    for (const r of a) {
      const o = rowOf(ours, r, 'ours');
      const twin = theirsByKey.get(keyOf(ours, r, 'ours'));
      if (twin !== undefined) {
        used.add(twin);
        const t = rowOf(theirs, twin, 'theirs');
        if (o.every((x, i) => same(x, t[i]))) stats.both++;
        else conflicts.push({ kind: 'both-added', row: rows.length, oursRow: o.map((x) => x[0]), theirsRow: t.map((x) => x[0]), took: prefer });
        push(prefer === 'ours' ? o : t);
      } else {
        stats.ours++;
        push(o);
      }
    }
    for (const r of b) {
      if (used.has(r)) continue;
      stats.theirs++;
      push(rowOf(theirs, r, 'theirs'));
    }
  };

  emitAdded(-1);
  for (let k = 0; k < base.rows.length; k++) {
    const a1 = b1[k];
    const a2 = b2[k];
    const gone1 = a1 < 0 || d1.status[a1] === ROW_STATUS.removed;
    const gone2 = a2 < 0 || d2.status[a2] === ROW_STATUS.removed;
    const edited1 = !gone1 && d1.status[a1] === ROW_STATUS.modified;
    const edited2 = !gone2 && d2.status[a2] === ROW_STATUS.modified;
    if (gone1 && gone2) stats.both++;
    else if (gone1 || gone2) {
      // Deleted on one side: fine if the other left it alone; a conflict if it edited it.
      const keptEdited = gone1 ? edited2 : edited1;
      if (!keptEdited) {
        stats[gone1 ? 'ours' : 'theirs']++;
      } else {
        const edited = gone1 ? rowOf(theirs, d2.newRows[a2], 'theirs') : rowOf(ours, d1.newRows[a1], 'ours');
        const keep = (prefer === 'ours' && !gone1) || (prefer === 'theirs' && !gone2);
        conflicts.push({
          kind: 'delete-edit',
          row: keep ? rows.length : -1,
          baseRow: k,
          oursRow: gone1 ? null : edited.map((x) => x[0]),
          theirsRow: gone2 ? null : edited.map((x) => x[0]),
          took: prefer,
        });
        if (keep) push(edited);
      }
    } else {
      const r1 = d1.newRows[a1];
      const r2 = d2.newRows[a2];
      const out = cols.map((c) => {
        const b = get(base, c.base >= 0 ? k : -1, c.base);
        const o = get(ours, r1, c.ours);
        const t = get(theirs, r2, c.theirs);
        // A column only one side has: that side's value.
        if (c.base < 0) return c.ours >= 0 ? o : t;
        const oc = !same(o, b);
        const tc = !same(t, b);
        if (!oc && !tc) return o;
        if (oc && !tc) return (stats.ours++, o);
        if (!oc && tc) return (stats.theirs++, t);
        if (same(o, t)) return (stats.both++, o);
        conflicts.push({ kind: 'cell', row: rows.length, baseRow: k, column: c.name, base: b[0], ours: o[0], theirs: t[0], took: prefer });
        return prefer === 'ours' ? o : t;
      });
      push(out);
    }
    emitAdded(k);
  }

  const table: Table = { name: ours.name || base.name, columns: cols.map((c) => c.name), rows, origin: base.origin };
  if (hasFormulas) table.formulas = formulas;
  return { table, conflicts, stats, diff: diffTables(base, table, options) };
}

/**
 * The conflicts as grid annotations for `result.diff`: pass them to
 * `DiffGrid` as `annotations`, and show each one's ours and theirs in
 * `renderAnnotation`. Rows the merge left out point at the removed row.
 */
export function conflictAnnotations(result: MergeResult): { row: number; column?: string; data: MergeConflict; id: number }[] {
  const d = result.diff;
  const byNew = new Map<number, number>();
  const byOld = new Map<number, number>();
  for (let a = 0; a < d.rowCount; a++) {
    if (d.newRows[a] >= 0) byNew.set(d.newRows[a], a);
    if (d.oldRows[a] >= 0) byOld.set(d.oldRows[a], a);
  }
  const out: { row: number; column?: string; data: MergeConflict; id: number }[] = [];
  result.conflicts.forEach((c, i) => {
    const row = c.row >= 0 ? byNew.get(c.row) : c.baseRow !== undefined ? byOld.get(c.baseRow) : undefined;
    if (row !== undefined) out.push({ row, ...(c.column ? { column: c.column } : {}), data: c, id: i });
  });
  return out;
}
