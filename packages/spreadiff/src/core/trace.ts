/**
 * Change tracing: why did this cell change?
 *
 * Every changed cell is classified as
 *   - `input`     its value was edited (no formula),
 *   - `formula`   its formula was edited,
 *   - `computed`  same formula, different result: it changed *because*
 *                 something upstream did ("downstream" change).
 * Cells of added/removed rows and columns are `added` / `removed` sources.
 *
 * For computed cells we find the root causes (the input/formula/structural
 * changes they depend on, transitively) by walking the formula dependency
 * graph forward from every root change. The walk only passes through cells
 * whose value actually changed: if an intermediate result is the same, the
 * change didn't flow through it. That keeps the walk small and the answer
 * right.
 *
 * Formulas are not evaluated: values come from the files. That's what makes
 * this cheap, and also its limit (no "what if" attribution of how much of a
 * change each cause contributed).
 */
import { formulaRefs, type SheetRange } from './formula.js';
import { ROW_STATUS, type TableDiff, type WorkbookDiff } from './types.js';

/**
 * Why a cell changed:
 * - `input`: its value was edited (it has no formula).
 * - `formula`: its formula was edited (compared by meaning, not text).
 * - `computed`: same formula, different result. It changed *because*
 *   something upstream changed: a downstream change.
 * - `added` / `removed`: a cell of an added or removed row or column that
 *   a formula reads (only recorded when it actually moved something).
 */
export type ChangeKind = 'input' | 'formula' | 'computed' | 'added' | 'removed';

/** One cell of a {@link WorkbookDiff}, in aligned coordinates. */
export interface CellRef {
  /** Index into `WorkbookDiff.sheets`. */
  sheet: number;
  /** Aligned row index in that sheet's `TableDiff`. */
  row: number;
  /** Aligned column index (into `TableDiff.columns`). */
  col: number;
}

/** What the trace knows about one changed cell. */
export interface CellTrace {
  /** Why it changed. */
  kind: ChangeKind;
  /**
   * For `computed` cells: the root changes (inputs, formula edits, added or
   * removed cells) this result depends on, transitively, through cells whose
   * values changed. Capped at `TraceOptions.maxCauses`; see `causeCount`.
   */
  causes: CellRef[];
  /** Exact number of root causes, even when `causes` is capped. */
  causeCount: number;
  /**
   * For root changes: the `computed` cells this change flowed into. Capped
   * at `TraceOptions.maxImpacts`; see `impactCount`.
   */
  impacts: CellRef[];
  /** Exact number of computed cells this change flowed into, even when `impacts` is capped. */
  impactCount: number;
  /**
   * Set on `computed` cells with no traced cause:
   * - `volatile`: the formula uses NOW, TODAY, RAND, … and is treated as its own cause.
   * - `dynamic`: precedents can't be known statically (INDIRECT, OFFSET, unknown names, external links).
   * - `unknown`: the formula's inputs didn't change as far as the trace can see
   *   (e.g. it reads a cell outside the parsed range).
   */
  unexplained?: 'volatile' | 'dynamic' | 'unknown';
}

/** Totals over a {@link WorkbookTrace}. */
export interface TraceStats {
  /** Changed cells without a formula. */
  inputs: number;
  /** Cells whose formula was edited. */
  formulas: number;
  /** Formula results that changed because of upstream changes. */
  computed: number;
  /** Computed cells with no traced cause (including volatile and dynamic ones). */
  unexplained: number;
}

/** The result of {@link traceWorkbook}. Lookups are O(1). */
export interface WorkbookTrace {
  /** The trace of a changed cell, or undefined if the cell didn't change. */
  get(ref: CellRef): CellTrace | undefined;
  /**
   * Changed cells that this cell's formula reads directly: one level up,
   * for building cause trees. Includes cells of added/removed rows inside
   * the formula's ranges. At most 200.
   */
  parents(ref: CellRef): CellRef[];
  /** All traced cells of one sheet, keyed by `row * columns.length + col`. */
  sheet(index: number): ReadonlyMap<number, CellTrace>;
  /** Totals by kind. */
  stats: TraceStats;
  /** The work budget ran out; some causes and impacts may be missing. */
  truncated: boolean;
}

/** Options for {@link traceWorkbook}. */
export interface TraceOptions {
  /** Max causes kept per cell (`causeCount` stays exact). Default 25. */
  maxCauses?: number;
  /** Max impacts kept per root change (`impactCount` stays exact). Default 500. */
  maxImpacts?: number;
  /** Max graph steps before giving up and setting `truncated`. Default 5,000,000. */
  budget?: number;
}

type Side = 'old' | 'new';

// Cell ids: sheet × 2^34 + row × 2^14 + col (rows < 2^20, cols < 2^14).
const ROW_MUL = 2 ** 14;
const SHEET_MUL = 2 ** 34;
const cellId = (s: number, r: number, c: number) => s * SHEET_MUL + r * ROW_MUL + c;
const idSheet = (id: number) => Math.floor(id / SHEET_MUL);
const idRow = (id: number) => Math.floor((id % SHEET_MUL) / ROW_MUL);
const idCol = (id: number) => id % ROW_MUL;

interface Rect {
  sheet: number;
  r0: number;
  r1: number;
  c0: number;
  c1: number;
}

/** Formula graph of one side (old or new workbook), in that side's table coordinates. */
class Graph {
  /** dependents[sheet][col] = flat [r0, r1, formulaCellId, …] */
  private index: Map<number, number[]>[] = [];
  precedents = new Map<number, Rect[]>();
  flags = new Map<number, { dynamic: boolean; volatile: boolean }>();

  constructor(diff: WorkbookDiff, side: Side) {
    const tables = diff.sheets.map((s) => (side === 'old' ? s.diff.oldTable : s.diff.newTable));
    const byName = new Map(tables.map((t, i) => [t.name.toUpperCase(), i]));
    const names = side === 'old' ? diff.names?.old : diff.names?.new;
    tables.forEach((t, s) => {
      this.index[s] = new Map();
      const f = t.formulas;
      if (!f) return;
      const origin = t.origin ?? { row: 1, col: 0 };
      for (let r = 0; r < f.length; r++) {
        const row = f[r];
        if (!row) continue;
        for (let c = 0; c < row.length; c++) {
          const formula = row[c];
          if (!formula) continue;
          const refs = formulaRefs(formula, names);
          const id = cellId(s, r, c);
          const rects: Rect[] = [];
          for (const rg of refs.ranges) {
            const rect = this.clip(rg, s, byName, tables, origin);
            if (rect) rects.push(rect);
          }
          this.precedents.set(id, rects);
          if (refs.dynamic || refs.volatile) this.flags.set(id, { dynamic: refs.dynamic, volatile: refs.volatile });
        }
      }
    });
    for (const [id, rects] of this.precedents) {
      for (const rc of rects) {
        const cols = this.index[rc.sheet];
        for (let c = rc.c0; c <= rc.c1; c++) {
          let list = cols.get(c);
          if (!list) cols.set(c, (list = []));
          list.push(rc.r0, rc.r1, id);
        }
      }
    }
  }

  /** Sheet range → rectangle in the target table's coordinates (or null if outside it). */
  private clip(
    rg: SheetRange,
    self: number,
    byName: Map<string, number>,
    tables: { rows: unknown[]; columns: string[]; origin?: { row: number; col: number } }[],
    selfOrigin: { row: number; col: number },
  ): Rect | null {
    const sheet = rg.sheet === null ? self : byName.get(rg.sheet.toUpperCase());
    if (sheet === undefined) return null;
    const t = tables[sheet];
    const origin = sheet === self ? selfOrigin : t.origin ?? { row: 1, col: 0 };
    const nRows = t.rows.length;
    const nCols = t.columns.length;
    const r0 = rg.r0 < 0 ? 0 : Math.max(0, rg.r0 - origin.row);
    const r1 = rg.r1 < 0 ? nRows - 1 : Math.min(nRows - 1, rg.r1 - origin.row);
    const c0 = rg.c0 < 0 ? 0 : Math.max(0, rg.c0 - origin.col);
    const c1 = rg.c1 < 0 ? nCols - 1 : Math.min(nCols - 1, rg.c1 - origin.col);
    if (r0 > r1 || c0 > c1) return null;
    return { sheet, r0, r1, c0, c1 };
  }

  /** Formula cells that read (sheet, row, col). */
  dependents(sheet: number, row: number, col: number, out: number[]): void {
    const list = this.index[sheet]?.get(col);
    if (!list) return;
    for (let k = 0; k < list.length; k += 3) if (row >= list[k] && row <= list[k + 1]) out.push(list[k + 2]);
  }
}

/** Per-sheet translation between side coordinates and aligned coordinates. */
interface SheetMaps {
  diff: TableDiff;
  rows: Record<Side, Int32Array>; // side row → aligned row
  cols: Record<Side, Int32Array>; // side col → aligned col
}

function sheetMaps(diff: TableDiff): SheetMaps {
  const inv = (arr: Int32Array, n: number) => {
    const out = new Int32Array(n).fill(-1);
    for (let i = 0; i < arr.length; i++) if (arr[i] >= 0) out[arr[i]] = i;
    return out;
  };
  const cols = (side: Side, n: number) => {
    const out = new Int32Array(n).fill(-1);
    diff.columns.forEach((c, ci) => {
      const j = side === 'old' ? c.oldIndex : c.newIndex;
      if (j >= 0) out[j] = ci;
    });
    return out;
  };
  return {
    diff,
    rows: { old: inv(diff.oldRows, diff.oldTable.rows.length), new: inv(diff.newRows, diff.newTable.rows.length) },
    cols: { old: cols('old', diff.oldTable.columns.length), new: cols('new', diff.newTable.columns.length) },
  };
}

/** Trace every change in a workbook diff back to its root causes. */
export function traceWorkbook(diff: WorkbookDiff, options: TraceOptions = {}): WorkbookTrace {
  const maxCauses = options.maxCauses ?? 25;
  const maxImpacts = options.maxImpacts ?? 500;
  let budget = options.budget ?? 5_000_000;
  let truncated = false;
  const maps = diff.sheets.map((s) => sheetMaps(s.diff));
  const traces: Map<number, CellTrace>[] = diff.sheets.map(() => new Map());
  const key = (ref: CellRef) => ref.row * maps[ref.sheet].diff.columns.length + ref.col;
  const entry = (ref: CellRef, kind: ChangeKind): CellTrace => {
    const m = traces[ref.sheet];
    const k = key(ref);
    let t = m.get(k);
    if (!t) m.set(k, (t = { kind, causes: [], causeCount: 0, impacts: [], impactCount: 0 }));
    return t;
  };
  const stats: TraceStats = { inputs: 0, formulas: 0, computed: 0, unexplained: 0 };

  // ---- Classify changed cells ---------------------------------------------
  diff.sheets.forEach((s, si) => {
    const d = s.diff;
    for (const [row, mask] of d.changedCells) {
      const fm = d.formulaChanged?.get(row);
      const ni = d.newRows[row];
      for (let col = 0; col < mask.length; col++) {
        if (!mask[col]) continue;
        const nf = d.newTable.formulas?.[ni]?.[d.columns[col].newIndex] ?? null;
        const kind: ChangeKind = fm?.[col] ? 'formula' : nf ? 'computed' : 'input';
        entry({ sheet: si, row, col }, kind);
        if (kind === 'input') stats.inputs++;
        else if (kind === 'formula') stats.formulas++;
        else stats.computed++;
      }
    }
  });

  const changedKind = (ref: CellRef): ChangeKind | undefined => traces[ref.sheet].get(key(ref))?.kind;
  const toAligned = (side: Side, id: number): CellRef | null => {
    const sheet = idSheet(id);
    const m = maps[sheet];
    const row = m.rows[side][idRow(id)];
    const col = m.cols[side][idCol(id)];
    return row >= 0 && col >= 0 ? { sheet, row, col } : null;
  };

  // ---- Propagate from each root change ------------------------------------
  const graphs: Partial<Record<Side, Graph>> = {};
  const graph = (side: Side) => (graphs[side] ??= new Graph(diff, side));

  /** Walk forward from a root, recording it as a cause of every computed change reached. */
  const walk = (side: Side, startId: number, root: CellRef, rootTrace: CellTrace | null) => {
    const g = graph(side);
    const seen = new Set<number>([startId]);
    const queue = [startId];
    const deps: number[] = [];
    while (queue.length) {
      if (--budget < 0) {
        truncated = true;
        return;
      }
      const id = queue.pop()!;
      deps.length = 0;
      g.dependents(idSheet(id), idRow(id), idCol(id), deps);
      for (const dep of deps) {
        if (seen.has(dep)) continue;
        seen.add(dep);
        const ref = toAligned(side, dep);
        if (!ref) continue;
        const kind = changedKind(ref);
        // Unchanged result: the change stopped here.
        if (kind !== 'computed' && kind !== 'formula') continue;
        if (kind === 'computed') {
          const t = traces[ref.sheet].get(key(ref))!;
          t.causeCount++;
          if (t.causes.length < maxCauses) t.causes.push(root);
          if (rootTrace) {
            rootTrace.impactCount++;
            if (rootTrace.impacts.length < maxImpacts) rootTrace.impacts.push(ref);
          }
        }
        queue.push(dep);
      }
    }
  };

  // Roots in the new workbook: edited inputs/formulas, and added cells.
  diff.sheets.forEach((s, si) => {
    const d = s.diff;
    for (const [k, t] of [...traces[si]]) {
      if (t.kind !== 'input' && t.kind !== 'formula') continue;
      const ref = { sheet: si, row: Math.floor(k / d.columns.length), col: k % d.columns.length };
      walk('new', cellId(si, d.newRows[ref.row], d.columns[ref.col].newIndex), ref, t);
    }
    // Added rows / columns feed ranges like SUM(B2:B99).
    const g = graph('new');
    for (let row = 0; row < d.rowCount; row++) {
      const ni = d.newRows[row];
      if (ni < 0) continue;
      const rowAdded = d.status[row] === ROW_STATUS.added;
      d.columns.forEach((c, col) => {
        if (c.newIndex < 0 || !(rowAdded || c.status === 'added')) return;
        const deps: number[] = [];
        g.dependents(si, ni, c.newIndex, deps);
        if (!deps.length) return;
        const ref = { sheet: si, row, col };
        const before = traces[si].get(key(ref));
        const t = before ?? { kind: 'added' as const, causes: [], causeCount: 0, impacts: [], impactCount: 0 };
        walk('new', cellId(si, ni, c.newIndex), ref, t);
        if (!before && t.impactCount) traces[si].set(key(ref), t);
      });
    }
  });

  // Roots in the old workbook: removed rows / columns that formulas used to read.
  diff.sheets.forEach((s, si) => {
    const d = s.diff;
    const g = graph('old');
    for (let row = 0; row < d.rowCount; row++) {
      const oi = d.oldRows[row];
      if (oi < 0) continue;
      const rowRemoved = d.status[row] === ROW_STATUS.removed;
      d.columns.forEach((c, col) => {
        if (c.oldIndex < 0 || !(rowRemoved || c.status === 'removed')) return;
        const deps: number[] = [];
        g.dependents(si, oi, c.oldIndex, deps);
        if (!deps.length) return;
        const ref = { sheet: si, row, col };
        const t: CellTrace = { kind: 'removed', causes: [], causeCount: 0, impacts: [], impactCount: 0 };
        walk('old', cellId(si, oi, c.oldIndex), ref, t);
        if (t.impactCount) traces[si].set(key(ref), t);
      });
    }
  });

  // Computed changes with no cause: volatile / dynamic formulas are their own
  // cause, and their downstream inherits them.
  const g = graph('new');
  diff.sheets.forEach((s, si) => {
    const d = s.diff;
    for (const [k, t] of traces[si]) {
      if (t.kind !== 'computed' || t.causeCount) continue;
      const row = Math.floor(k / d.columns.length);
      const col = k % d.columns.length;
      const id = cellId(si, d.newRows[row], d.columns[col].newIndex);
      const flags = g.flags.get(id);
      if (flags) {
        t.unexplained = flags.volatile ? 'volatile' : 'dynamic';
        walk('new', id, { sheet: si, row, col }, t);
      }
    }
  });
  for (const m of traces) {
    for (const t of m.values()) {
      if (t.kind === 'computed' && !t.causeCount) {
        t.unexplained ??= 'unknown';
        stats.unexplained++;
      }
    }
  }

  return {
    get: (ref) => traces[ref.sheet]?.get(key(ref)),
    sheet: (i) => traces[i] ?? new Map(),
    stats,
    truncated,
    parents(ref) {
      const d = maps[ref.sheet]?.diff;
      if (!d) return [];
      const out: CellRef[] = [];
      const seen = new Set<string>();
      const add = (r: CellRef | null) => {
        if (!r) return;
        const s = `${r.sheet}:${r.row}:${r.col}`;
        if (seen.has(s) || (r.sheet === ref.sheet && r.row === ref.row && r.col === ref.col)) return;
        const t = traces[r.sheet].get(key(r));
        const st = maps[r.sheet].diff.status[r.row];
        if (t || st === ROW_STATUS.added || st === ROW_STATUS.removed) {
          seen.add(s);
          out.push(r);
        }
      };
      for (const side of ['new', 'old'] as const) {
        const ri = side === 'new' ? d.newRows[ref.row] : d.oldRows[ref.row];
        const ci = side === 'new' ? d.columns[ref.col].newIndex : d.columns[ref.col].oldIndex;
        if (ri < 0 || ci < 0) continue;
        const rects = graph(side).precedents.get(cellId(ref.sheet, ri, ci)) ?? [];
        for (const rc of rects) {
          const area = (rc.r1 - rc.r0 + 1) * (rc.c1 - rc.c0 + 1);
          const target = maps[rc.sheet].diff;
          if (area <= 4096) {
            for (let r = rc.r0; r <= rc.r1; r++) for (let c = rc.c0; c <= rc.c1; c++) add(toAligned(side, cellId(rc.sheet, r, c)));
          } else {
            // Big range: scan the (usually few) changed rows instead.
            for (let row = 0; row < target.rowCount && out.length < 200; row++) {
              const sr = side === 'new' ? target.newRows[row] : target.oldRows[row];
              if (sr < rc.r0 || sr > rc.r1) continue;
              const st = target.status[row];
              if (st === ROW_STATUS.equal) continue;
              for (let col = 0; col < target.columns.length; col++) {
                const sc = side === 'new' ? target.columns[col].newIndex : target.columns[col].oldIndex;
                if (sc >= rc.c0 && sc <= rc.c1) add({ sheet: rc.sheet, row, col });
              }
            }
          }
        }
      }
      return out.slice(0, 200);
    },
  };
}

/** Convenience: trace a single table diff (sheet index 0). */
export function traceTable(diff: TableDiff, options?: TraceOptions): WorkbookTrace {
  return traceWorkbook({ sheets: [{ name: diff.newTable.name, status: 'modified', diff }] }, options);
}
