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
 * Root causes are found in two layers:
 *
 * 1. **Recalculation** (default). The formulas around each edit are
 *    re-evaluated with the files' values. Each formula must first
 *    reproduce the *old* file's cached result (calibration); only then is
 *    it trusted. Each root edit is then undone in turn and the affected
 *    cells recalculated: a cell that moves back was caused by that edit
 *    (a but-for cause), and the difference is the edit's contribution.
 *    Evaluation also records what each formula really read (the branch an
 *    IF took, the row a lookup returned), and finds cached values the file
 *    left stale (saved without recalculating).
 * 2. **References.** Where a formula can't be evaluated (an unsupported
 *    function, a value the engine can't reproduce), its references are
 *    used instead: an edit causes a change if a chain of changed cells
 *    connects them. This over-approximates (both IF branches count) but
 *    doesn't miss a cause.
 */
import { externalSource, type ExternalSource } from './external.js';
import { formulaRefs, scopedNames, type SheetRange } from './formula.js';
import { numericDelta } from './text.js';
import { Engine, Unsupported, type Compiled, type Read } from './engine/evaluate.js';
import { tableHost, sheetCellKey, type TableHost } from './engine/host.js';
import { sameResult, isError, type Scalar } from './engine/values.js';
import { ROW_STATUS, type CellValue, type TableDiff, type WorkbookDiff } from './types.js';

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
   * removed cells) that caused this result. Capped at
   * `TraceOptions.maxCauses`; see `causeCount`.
   */
  causes: CellRef[];
  /** Exact number of root causes, even when `causes` is capped. */
  causeCount: number;
  /**
   * For `computed` cells, parallel to `causes`: how much of the change each
   * cause accounts for, as *new value − value without that cause*. Present
   * when `method` is `'recalculated'` (null where the values aren't
   * numbers). The parts needn't add up to the total: causes interact (a
   * rate times a balance), and a change can have several sufficient causes.
   */
  contributions?: (number | null)[];
  /**
   * How `causes` was found:
   * - `'recalculated'`: by undoing each edit and recalculating (exact).
   * - `'references'`: by following formula references through changed
   *   cells (may include causes that didn't actually matter).
   */
  method?: 'recalculated' | 'references';
  /** For root changes: the `computed` cells this change caused. Capped at `TraceOptions.maxImpacts`. */
  impacts: CellRef[];
  /** Exact number of computed cells this change caused, even when `impacts` is capped. */
  impactCount: number;
  /**
   * Set on `computed` cells with no traced cause:
   * - `volatile`: the formula uses NOW, TODAY, RAND, … and is treated as its own cause.
   * - `dynamic`: precedents can't be known statically (INDIRECT, OFFSET,
   *   unknown names, external links) and the formula couldn't be evaluated.
   * - `unknown`: the formula's inputs didn't change as far as the trace can
   *   see (e.g. it reads a cell outside the parsed range, or the two files
   *   were calculated by different spreadsheet engines).
   */
  unexplained?: 'volatile' | 'dynamic' | 'unknown';
  /**
   * The new file's cached value is stale: recalculating with the file's own
   * inputs gives `recalculated` (it was saved without recalculating, e.g.
   * in manual calculation mode or by a script). The cell is traced as
   * changing to `recalculated`, even if its cached value didn't change.
   */
  stale?: { cached: CellValue; recalculated: CellValue };
  /**
   * Set on `input` cells that hold a formula reading data from outside the
   * workbook (a Bloomberg `BDP`, `RTD`, a link to another file; see
   * {@link externalSource}). Their value can't be recalculated here, so the
   * file's saved value is taken as given: when it changed (the data
   * refreshed), the cell is an input, and a root cause of what it feeds.
   */
  external?: ExternalSource;
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
  /** Computed cells whose new cached value was stale (see `CellTrace.stale`). */
  stale: number;
  /** Computed cells whose causes were found by recalculation. */
  recalculated: number;
  /** Formulas near a change that couldn't be recalculated, so causes through them come from references. */
  unmodelled: number;
  /** Inputs whose value came from outside the workbook (`CellTrace.external`); counted in `inputs` too. */
  external: number;
}

/** The result of {@link traceWorkbook}. Lookups are O(1). */
export interface WorkbookTrace {
  /** The trace of a changed cell, or undefined if the cell didn't change. */
  get(ref: CellRef): CellTrace | undefined;
  /**
   * Changed cells that this cell's formula reads directly: one level up,
   * for building cause trees. Uses what the formula actually read when it
   * was recalculated, and its references otherwise. Includes cells of
   * added/removed rows inside the formula's ranges. At most 200.
   */
  parents(ref: CellRef): CellRef[];
  /** All traced cells of one sheet, keyed by `row * columns.length + col`. */
  sheet(index: number): ReadonlyMap<number, CellTrace>;
  /** Totals by kind. */
  stats: TraceStats;
  /** A work budget ran out; some causes may be missing or come from references. */
  truncated: boolean;
}

/**
 * A workbook trace, plus which sheet of the workbook a grid shows. Each
 * `SheetView` from {@link compareWorkbooks} or `diffFiles` carries one.
 */
export interface TraceContext {
  /** From {@link traceWorkbook}. */
  trace: WorkbookTrace;
  /** The workbook diff the trace was made from. Used to label causes on other sheets. */
  workbook: WorkbookDiff;
  /**
   * Index of the grid's sheet in `workbook.sheets`. Optional, and ignored
   * when the grid's diff is one of `workbook.sheets`: the grid uses that
   * sheet. Needed only for a diff that isn't (a copy, say).
   */
  sheet?: number;
}

/** Options for {@link traceWorkbook}. */
export interface TraceOptions {
  /** Max causes kept per cell (`causeCount` stays exact). Default 25. */
  maxCauses?: number;
  /** Max impacts kept per root change (`impactCount` stays exact). Default 500. */
  maxImpacts?: number;
  /** Max graph steps before giving up and setting `truncated`. Default 5,000,000. */
  budget?: number;
  /**
   * Recalculate formulas to find exact causes, contributions and stale
   * values. Default true. With false, causes come from references only.
   */
  recalculate?: boolean;
  /** Max formula evaluations for recalculation, across the trace. Default 2,000,000. */
  evaluationBudget?: number;
  /**
   * More live-data functions, by provider, on top of {@link LIVE_FUNCTIONS}
   * (`{ 'In-house': ['PRICE_OF'] }`). A changed cell calling one is an
   * external input rather than a recalculated result.
   */
  liveFunctions?: Readonly<Record<string, readonly string[]>>;
}

type Side = 'old' | 'new';

// Cell ids in table coordinates: sheet × 2^34 + row × 2^14 + col (rows < 2^20, cols < 2^14).
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

/** A formula dependency graph over table coordinates, indexed for "who reads this cell?". */
class Graph {
  /** Single-cell precedents: cell id → formula ids that read exactly it. */
  private single = new Map<number, number[]>();
  /** Range precedents: per sheet, per column, flat [r0, r1, formulaId, …]. */
  private ranges: Map<number, number[]>[] = [];

  constructor(
    readonly precedents: Map<number, Rect[]>,
    readonly flags: Map<number, { dynamic: boolean; volatile: boolean }> = new Map(),
  ) {
    for (const [id, rects] of precedents) {
      for (const rc of rects) {
        if (rc.r0 === rc.r1 && rc.c0 === rc.c1) {
          const k = cellId(rc.sheet, rc.r0, rc.c0);
          const list = this.single.get(k);
          if (list) list.push(id);
          else this.single.set(k, [id]);
          continue;
        }
        const cols = (this.ranges[rc.sheet] ??= new Map());
        for (let c = rc.c0; c <= rc.c1; c++) {
          let list = cols.get(c);
          if (!list) cols.set(c, (list = []));
          list.push(rc.r0, rc.r1, id);
        }
      }
    }
  }

  /** Formula cells that read (sheet, row, col). */
  dependents(sheet: number, row: number, col: number, out: number[]): void {
    const one = this.single.get(cellId(sheet, row, col));
    if (one) for (const id of one) out.push(id);
    const list = this.ranges[sheet]?.get(col);
    if (!list) return;
    for (let k = 0; k < list.length; k += 3) if (row >= list[k] && row <= list[k + 1]) out.push(list[k + 2]);
  }
}

/** What the walks need from a graph. */
interface DependencyGraph {
  readonly precedents: { get(id: number): Rect[] | undefined };
  readonly flags: Map<number, { dynamic: boolean; volatile: boolean }>;
  dependents(sheet: number, row: number, col: number, out: number[]): void;
}

/** A graph whose formulas in `replaced` read what's listed there instead of their references. */
class OverlayGraph implements DependencyGraph {
  private extra: Graph;
  readonly precedents: { get(id: number): Rect[] | undefined };
  readonly flags: Map<number, { dynamic: boolean; volatile: boolean }>;
  constructor(
    private base: Graph,
    private replaced: Map<number, Rect[]>,
  ) {
    this.extra = new Graph(replaced);
    this.flags = base.flags;
    this.precedents = { get: (id) => replaced.get(id) ?? base.precedents.get(id) };
  }
  dependents(sheet: number, row: number, col: number, out: number[]): void {
    const start = out.length;
    this.base.dependents(sheet, row, col, out);
    // Drop formulas whose actual reads replace their references.
    let w = start;
    for (let i = start; i < out.length; i++) if (!this.replaced.has(out[i])) out[w++] = out[i];
    out.length = w;
    this.extra.dependents(sheet, row, col, out);
  }
}

/**
 * Dependency order of `nodes` (formula cells), given `sources` that feed
 * them: Kahn's algorithm over the edges inside the set. `preds` lists each
 * node's predecessors in the set; nodes on a cycle come back in `cyclic`.
 */
function topoOrder(sources: number[], nodes: Set<number>, graph: DependencyGraph): { order: number[]; preds: Map<number, number[]>; cyclic: number[] } {
  const preds = new Map<number, number[]>();
  const indeg = new Map<number, number>();
  for (const n of nodes) indeg.set(n, 0);
  const deps: number[] = [];
  const edgesFrom = (u: number) => {
    deps.length = 0;
    graph.dependents(idSheet(u), idRow(u), idCol(u), deps);
    for (const v of deps) {
      if (v === u || !nodes.has(v)) continue;
      let p = preds.get(v);
      if (!p) preds.set(v, (p = []));
      if (p[p.length - 1] === u) continue; // the same edge listed twice (two references)
      p.push(u);
      if (nodes.has(u)) indeg.set(v, indeg.get(v)! + 1);
    }
  };
  for (const u of sources) if (!nodes.has(u)) edgesFrom(u);
  for (const u of nodes) edgesFrom(u);
  const order: number[] = [];
  const ready: number[] = [];
  for (const [n, d] of indeg) if (d === 0) ready.push(n);
  while (ready.length) {
    const u = ready.pop()!;
    order.push(u);
    deps.length = 0;
    graph.dependents(idSheet(u), idRow(u), idCol(u), deps);
    let last = -1;
    for (const v of deps) {
      if (v === u || v === last || !nodes.has(v)) continue;
      last = v;
      const d = indeg.get(v)! - 1;
      indeg.set(v, d);
      if (d === 0) ready.push(v);
    }
  }
  const cyclic = order.length === nodes.size ? [] : [...nodes].filter((n) => indeg.get(n)! > 0);
  return { order, preds, cyclic };
}

/** Sheet range → rectangle in the target table's coordinates (or null if outside it). */
function clip(
  rg: SheetRange,
  self: number,
  byName: Map<string, number>,
  tables: { rows: unknown[]; columns: string[]; origin?: { row: number; col: number } }[],
): Rect | null {
  const sheet = rg.sheet === null ? self : byName.get(rg.sheet.toUpperCase());
  if (sheet === undefined) return null;
  const t = tables[sheet];
  const origin = t.origin ?? { row: 1, col: 0 };
  const nRows = t.rows.length;
  const nCols = t.columns.length;
  const r0 = rg.r0 < 0 ? 0 : Math.max(0, rg.r0 - origin.row);
  const r1 = rg.r1 < 0 ? nRows - 1 : Math.min(nRows - 1, rg.r1 - origin.row);
  const c0 = rg.c0 < 0 ? 0 : Math.max(0, rg.c0 - origin.col);
  const c1 = rg.c1 < 0 ? nCols - 1 : Math.min(nCols - 1, rg.c1 - origin.col);
  if (r0 > r1 || c0 > c1) return null;
  return { sheet, r0, r1, c0, c1 };
}

/** A side's formula references, one cell at a time: precedents (clipped to the sheets) and flags, or null for a non-formula cell. */
function refReader(diff: WorkbookDiff, side: Side) {
  const tables = diff.sheets.map((s) => (side === 'old' ? s.diff.oldTable : s.diff.newTable));
  const byName = new Map(tables.map((t, i) => [t.name.toUpperCase(), i]));
  const names = side === 'old' ? diff.names?.old : diff.names?.new;
  const tableDefs = side === 'old' ? diff.tables?.old : diff.tables?.new;
  const sheetNames = tables.map((t) => t.name);
  // Sheet-scoped names: each sheet's formulas see their own over the workbook's.
  const scoped = tables.some((t) => t.names) ? tables.map((t) => scopedNames(names, tables, t.name)) : null;
  return {
    tables,
    read(s: number, r: number, c: number): { rects: Rect[]; dynamic: boolean; volatile: boolean } | null {
      const formula = tables[s]?.formulas?.[r]?.[c];
      if (!formula) return null;
      const origin = tables[s].origin ?? { row: 1, col: 0 };
      const refs = formulaRefs(formula, scoped ? scoped[s] : names, { tables: tableDefs, row: r + origin.row, sheets: sheetNames });
      const rects: Rect[] = [];
      for (const rg of refs.ranges) {
        const rect = clip(rg, s, byName, tables);
        if (rect) rects.push(rect);
      }
      return { rects, dynamic: refs.dynamic, volatile: refs.volatile };
    },
  };
}

/** Static formula graph of one side, from formula references. */
function staticGraph(diff: WorkbookDiff, side: Side): Graph {
  const reader = refReader(diff, side);
  const precedents = new Map<number, Rect[]>();
  const flags = new Map<number, { dynamic: boolean; volatile: boolean }>();
  reader.tables.forEach((t, s) => {
    const f = t.formulas;
    if (!f) return;
    for (let r = 0; r < f.length; r++) {
      const row = f[r];
      if (!row) continue;
      for (let c = 0; c < row.length; c++) {
        const refs = reader.read(s, r, c);
        if (!refs) continue;
        const id = cellId(s, r, c);
        precedents.set(id, refs.rects);
        if (refs.dynamic || refs.volatile) flags.set(id, { dynamic: refs.dynamic, volatile: refs.volatile });
      }
    }
  });
  return new Graph(precedents, flags);
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

interface Root {
  ref: CellRef;
  trace: CellTrace;
  side: Side;
  /** Cell id on `side`, in table coordinates. */
  id: number;
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
  const refOfKey = (si: number, k: number): CellRef => {
    const w = maps[si].diff.columns.length;
    return { sheet: si, row: Math.floor(k / w), col: k % w };
  };
  const blank = (kind: ChangeKind): CellTrace => ({ kind, causes: [], causeCount: 0, impacts: [], impactCount: 0 });

  // ---- Classify changed cells ---------------------------------------------
  diff.sheets.forEach((s, si) => {
    const d = s.diff;
    for (const [row, mask] of d.changedCells) {
      const fm = d.formulaChanged?.get(row);
      const ni = d.newRows[row];
      for (let col = 0; col < mask.length; col++) {
        if (!mask[col]) continue;
        const nf = d.newTable.formulas?.[ni]?.[d.columns[col].newIndex] ?? null;
        const ext = nf && !fm?.[col] ? externalSource(nf, options.liveFunctions) : null;
        const t = blank(fm?.[col] ? 'formula' : nf && !ext ? 'computed' : 'input');
        if (ext) t.external = ext;
        traces[si].set(key({ sheet: si, row, col }), t);
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
  const sideId = (side: Side, ref: CellRef): number | null => {
    const d = maps[ref.sheet].diff;
    const r = side === 'new' ? d.newRows[ref.row] : d.oldRows[ref.row];
    const c = side === 'new' ? d.columns[ref.col].newIndex : d.columns[ref.col].oldIndex;
    return r >= 0 && c >= 0 ? cellId(ref.sheet, r, c) : null;
  };

  // ---- Roots --------------------------------------------------------------
  const roots: Root[] = [];
  const staticNew = staticGraph(diff, 'new');
  let staticOld: Graph | null = null;
  const oldGraph = () => (staticOld ??= staticGraph(diff, 'old'));
  let oldReader: ReturnType<typeof refReader> | null = null;
  diff.sheets.forEach((s, si) => {
    const d = s.diff;
    for (const [k, t] of traces[si]) {
      if (t.kind !== 'input' && t.kind !== 'formula') continue;
      const ref = refOfKey(si, k);
      roots.push({ ref, trace: t, side: 'new', id: cellId(si, d.newRows[ref.row], d.columns[ref.col].newIndex) });
    }
    // Added rows / columns feed ranges like SUM(B2:B99).
    for (let row = 0; row < d.rowCount; row++) {
      const ni = d.newRows[row];
      if (ni < 0) continue;
      const rowAdded = d.status[row] === ROW_STATUS.added;
      d.columns.forEach((c, col) => {
        if (c.newIndex < 0 || !(rowAdded || c.status === 'added')) return;
        const deps: number[] = [];
        staticNew.dependents(si, ni, c.newIndex, deps);
        if (!deps.length) return;
        const ref = { sheet: si, row, col };
        roots.push({ ref, trace: traces[si].get(key(ref)) ?? blank('added'), side: 'new', id: cellId(si, ni, c.newIndex) });
      });
    }
  });
  // Removed rows / columns that formulas used to read (old workbook).
  diff.sheets.forEach((s, si) => {
    const d = s.diff;
    for (let row = 0; row < d.rowCount; row++) {
      const oi = d.oldRows[row];
      if (oi < 0) continue;
      const rowRemoved = d.status[row] === ROW_STATUS.removed;
      d.columns.forEach((c, col) => {
        if (c.oldIndex < 0 || !(rowRemoved || c.status === 'removed')) return;
        const deps: number[] = [];
        oldGraph().dependents(si, oi, c.oldIndex, deps);
        if (!deps.length) return;
        roots.push({ ref: { sheet: si, row, col }, trace: blank('removed'), side: 'old', id: cellId(si, oi, c.oldIndex) });
      });
    }
  });

  // ---- Recalculation ------------------------------------------------------
  const analysis = options.recalculate === false ? null : recalculate(diff, maps, traces, roots, staticNew, options.evaluationBudget ?? 2_000_000);
  if (analysis?.truncated) truncated = true;
  const newGraph: DependencyGraph = analysis?.graph ?? staticNew;

  // ---- Causes by references: walk forward through changed cells ----------
  const causeSets: Map<number, Set<number>>[] = diff.sheets.map(() => new Map());
  const walk = (side: Side, startId: number, root: number) => {
    const g = side === 'new' ? newGraph : oldGraph();
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
          const m = causeSets[ref.sheet];
          const k = key(ref);
          let set = m.get(k);
          if (!set) m.set(k, (set = new Set()));
          set.add(root);
        }
        queue.push(dep);
      }
    }
  };
  roots.forEach((r, i) => walk(r.side, r.id, i));

  // Computed changes with no cause: volatile / dynamic formulas that couldn't
  // be recalculated are their own cause, and their downstream inherits them.
  const selfCaused = new Set<number>();
  diff.sheets.forEach((s, si) => {
    for (const [k, t] of traces[si]) {
      if (t.kind !== 'computed' || causeSets[si].get(k)?.size) continue;
      const ref = refOfKey(si, k);
      const id = sideId('new', ref);
      if (id === null || analysis?.modelled.has(id)) continue;
      const flags = staticNew.flags.get(id);
      if (!flags) continue;
      t.unexplained = flags.volatile ? 'volatile' : 'dynamic';
      const i = roots.push({ ref, trace: t, side: 'new', id }) - 1;
      selfCaused.add(i);
      walk('new', id, i);
    }
  });

  // ---- Exact causes: undo each root and recalculate (but-for) -------------
  const exact = analysis?.butFor(roots) ?? null;

  // ---- Assemble -------------------------------------------------------------
  const stats: TraceStats = { inputs: 0, formulas: 0, computed: 0, unexplained: 0, stale: 0, recalculated: 0, unmodelled: analysis?.unmodelled ?? 0, external: 0 };
  diff.sheets.forEach((_, si) => {
    for (const [k, t] of traces[si]) {
      if (t.kind === 'input') stats.inputs++;
      if (t.external) stats.external++;
      else if (t.kind === 'formula') stats.formulas++;
      if (t.kind !== 'computed') continue;
      stats.computed++;
      if (t.stale) stats.stale++;
      const ref = refOfKey(si, k);
      const byRefs = causeSets[si].get(k);
      const precise = exact?.get(si)?.get(k);
      let list: number[];
      // Added/removed rows and self-caused cells aren't undone one by one: they come from references.
      const structural = [...(byRefs ?? [])].filter((r) => roots[r].trace.kind !== 'input' && roots[r].trace.kind !== 'formula');
      if (precise && precise.causes.length && !precise.tainted) {
        list = [...precise.causes, ...structural];
        t.contributions = [...precise.contributions, ...structural.map(() => null)].slice(0, maxCauses);
        t.method = 'recalculated';
        stats.recalculated++;
      } else {
        // Exact causes first, then any reachable only through formulas the engine couldn't evaluate.
        list = precise?.causes.length ? [...precise.causes, ...[...(byRefs ?? [])].filter((r) => !precise.causes.includes(r))] : [...(byRefs ?? [])];
        t.method = 'references';
      }
      list = list.filter((r) => !(selfCaused.has(r) && roots[r].trace === t));
      t.causeCount = list.length;
      t.causes = list.slice(0, maxCauses).map((r) => roots[r].ref);
      for (const r of list) {
        const rt = roots[r].trace;
        rt.impactCount++;
        if (rt.impacts.length < maxImpacts) rt.impacts.push(ref);
      }
      if (!t.causeCount) {
        t.unexplained ??= 'unknown';
        stats.unexplained++;
      }
    }
  });
  // Added/removed cells are traced only when they moved something.
  for (const r of roots) {
    if ((r.trace.kind === 'added' || r.trace.kind === 'removed') && r.trace.impactCount) traces[r.ref.sheet].set(key(r.ref), r.trace);
  }

  // Added/removed rows per sheet, found once: parents() checks big ranges against them.
  const structural: number[][] = [];
  const structuralRows = (si: number) => {
    let rows = structural[si];
    if (!rows) {
      rows = structural[si] = [];
      const d = maps[si].diff;
      for (let row = 0; row < d.rowCount; row++) if (d.status[row] === ROW_STATUS.added || d.status[row] === ROW_STATUS.removed) rows.push(row);
    }
    return rows;
  };

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
        const id = sideId(side, ref);
        if (id === null) continue;
        // The old side's whole graph is only built when removed cells need it; one cell's references are enough here.
        const rects = side === 'new' ? newGraph.precedents.get(id) : staticOld ? staticOld.precedents.get(id) : (oldReader ??= refReader(diff, 'old')).read(idSheet(id), idRow(id), idCol(id))?.rects;
        for (const rc of rects ?? []) {
          const area = (rc.r1 - rc.r0 + 1) * (rc.c1 - rc.c0 + 1);
          const target = maps[rc.sheet].diff;
          if (area <= 4096) {
            for (let r = rc.r0; r <= rc.r1; r++) for (let c = rc.c0; c <= rc.c1; c++) add(toAligned(side, cellId(rc.sheet, r, c)));
            continue;
          }
          // Big range: visit the traced cells inside it instead of every cell.
          for (const [k] of traces[rc.sheet]) {
            if (out.length >= 200) break;
            const cand = refOfKey(rc.sheet, k);
            const cid = sideId(side, cand);
            if (cid === null) continue;
            const r = idRow(cid);
            const c = idCol(cid);
            if (r >= rc.r0 && r <= rc.r1 && c >= rc.c0 && c <= rc.c1) add(cand);
          }
          for (const row of structuralRows(rc.sheet)) {
            if (out.length >= 200) break;
            const sr = side === 'new' ? target.newRows[row] : target.oldRows[row];
            if (sr < rc.r0 || sr > rc.r1) continue;
            for (let col = 0; col < target.columns.length; col++) {
              const sc = side === 'new' ? target.columns[col].newIndex : target.columns[col].oldIndex;
              if (sc >= rc.c0 && sc <= rc.c1) add({ sheet: rc.sheet, row, col });
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

// ===========================================================================
// Recalculation
// ===========================================================================

interface ExactCauses {
  /** Root indexes that are but-for causes. */
  causes: number[];
  /** Parallel to `causes`: new value − value without that root. */
  contributions: (number | null)[];
  /** Downstream of a formula the engine couldn't evaluate: exact causes may be incomplete. */
  tainted: boolean;
}

interface Analysis {
  /** New-side graph: what trusted formulas actually read, references for the rest. */
  graph: DependencyGraph;
  /** New-side formula cells (table ids) the engine evaluated and trusts. */
  modelled: Set<number>;
  /** Formulas near a change that couldn't be modelled. */
  unmodelled: number;
  truncated: boolean;
  /** Exact causes per computed cell: sheet → aligned key → causes. */
  butFor(roots: Root[]): Map<number, Map<number, ExactCauses>>;
}

/** Origin of a table, with the default. */
const originOf = (t: { origin?: { row: number; col: number } }) => t.origin ?? { row: 1, col: 0 };

function recalculate(diff: WorkbookDiff, maps: SheetMaps[], traces: Map<number, CellTrace>[], roots: Root[], staticNew: Graph, evalBudget: number): Analysis | null {
  const oldTables = diff.sheets.map((s) => s.diff.oldTable);
  const newTables = diff.sheets.map((s) => s.diff.newTable);
  if (!newTables.some((t) => t.formulas?.some((r) => r?.some(Boolean)))) return null;
  const oldHost = tableHost(oldTables, { names: diff.names?.old, tables: diff.tables?.old, date1904: diff.date1904?.old });
  const newHost = tableHost(newTables, { names: diff.names?.new, tables: diff.tables?.new, date1904: diff.date1904?.new });
  const oldEngine = new Engine(oldHost);
  const newEngine = new Engine(newHost);
  let evaluations = 0;
  let truncated = false;
  const traceKey = (ref: CellRef) => ref.row * maps[ref.sheet].diff.columns.length + ref.col;
  const tablesOf = (side: Side) => (side === 'new' ? newTables : oldTables);
  const hostOf = (side: Side) => (side === 'new' ? newHost : oldHost);
  /** Table id → engine key (sheet coordinates). */
  const sheetKey = (side: Side, id: number) => {
    const s = idSheet(id);
    const o = originOf(tablesOf(side)[s]);
    return sheetCellKey(s, idRow(id) + o.row, idCol(id) + o.col);
  };
  const formulaAt = (side: Side, id: number) => tablesOf(side)[idSheet(id)].formulas?.[idRow(id)]?.[idCol(id)] ?? null;
  const cachedAt = (side: Side, id: number): Scalar => {
    const s = idSheet(id);
    const o = originOf(tablesOf(side)[s]);
    return hostOf(side).cached(s, idRow(id) + o.row, idCol(id) + o.col);
  };
  // Compiled formulas by cell, per side (the but-for pass re-evaluates cells many times).
  const compiled: Record<Side, Map<number, Compiled | Unsupported>> = { old: new Map(), new: new Map() };
  const compileAt = (side: Side, id: number, formula: string): Compiled => {
    let c = compiled[side].get(id);
    if (c === undefined) {
      const s = idSheet(id);
      const o = originOf(tablesOf(side)[s]);
      try {
        c = (side === 'new' ? newEngine : oldEngine).compile(formula, idRow(id) + o.row, idCol(id) + o.col);
      } catch (e) {
        if (!(e instanceof Unsupported)) throw e;
        c = e;
      }
      compiled[side].set(id, c);
    }
    if (c instanceof Unsupported) throw c;
    return c;
  };
  const evaluate = (engine: Engine, host: TableHost, side: Side, id: number, formula: string, record = false) => {
    if (++evaluations > evalBudget) {
      truncated = true;
      throw new Unsupported('evaluation budget');
    }
    const c = compileAt(side, id, formula);
    const s = idSheet(id);
    const o = originOf(tablesOf(side)[s]);
    return engine.evaluate(c, host.sheets[s], idRow(id) + o.row, idCol(id) + o.col, record);
  };
  /** Calibration by formula shape: after enough copies of a formula reproduce the old file, trust the rest. */
  const CALIBRATION_SAMPLE = 8;
  const shapes = new Map<string, { passed: number; failed: boolean }>();
  const toAligned = (side: Side, id: number): CellRef | null => {
    const sheet = idSheet(id);
    const m = maps[sheet];
    const row = m.rows[side][idRow(id)];
    const col = m.cols[side][idCol(id)];
    return row >= 0 && col >= 0 ? { sheet, row, col } : null;
  };
  const oldIdOf = (ref: CellRef): number | null => {
    const d = maps[ref.sheet].diff;
    const r = d.oldRows[ref.row];
    const c = d.columns[ref.col].oldIndex;
    return r >= 0 && c >= 0 ? cellId(ref.sheet, r, c) : null;
  };
  /** Engine reads (sheet coordinates) → new-side table rectangles. */
  const rectsOf = (reads: Read[]): Rect[] => {
    const out: Rect[] = [];
    for (const rd of reads) {
      const s = newHost.sheetIndex(rd.sheet);
      if (s < 0) continue;
      const t = newTables[s];
      const o = originOf(t);
      const r0 = Math.max(0, rd.r0 - o.row);
      const r1 = Math.min(t.rows.length - 1, rd.r1 - o.row);
      const c0 = Math.max(0, rd.c0 - o.col);
      const c1 = Math.min(t.columns.length - 1, rd.c1 - o.col);
      if (r0 <= r1 && c0 <= c1) out.push({ sheet: s, r0, r1, c0, c1 });
    }
    return out;
  };

  // 1. The cone: new-side formulas reachable from a root through references.
  const MAX_CONE = 300_000;
  const cone = new Set<number>();
  const queue: number[] = [];
  const queued = new Set<number>();
  for (const r of roots) {
    if (r.side !== 'new' || queued.has(r.id)) continue;
    queued.add(r.id);
    queue.push(r.id);
    if (formulaAt('new', r.id) && !r.trace.external) cone.add(r.id);
  }
  const deps: number[] = [];
  for (let qi = 0; qi < queue.length && !truncated; qi++) {
    deps.length = 0;
    const id = queue[qi];
    staticNew.dependents(idSheet(id), idRow(id), idCol(id), deps);
    for (const d of deps) {
      if (cone.has(d)) continue;
      if (cone.size >= MAX_CONE) {
        truncated = true;
        break;
      }
      cone.add(d);
      if (!queued.has(d)) {
        queued.add(d);
        queue.push(d);
      }
    }
  }
  if (truncated) return { graph: staticNew, modelled: new Set(), unmodelled: 0, truncated, butFor: () => new Map() };

  // 2. Evaluate the cone in dependency order. Each formula must reproduce the
  //    old file's cached value first (calibration); a trusted formula whose
  //    result disagrees with the new file's cache marks that cache stale.
  const modelled = new Set<number>();
  const reads = new Map<number, Rect[]>();
  const stale = new Map<number, Scalar>();
  const state = new Map<number, 1 | 2>(); // 1 evaluating, 2 done
  let unmodelled = 0;
  let depth = 0;
  const baseGet = newHost.get;
  const baseGetRange = newHost.getRange;
  newHost.getRange = undefined; // cell by cell while evaluating, so the hook below sees every read
  // A cone cell read while evaluating another is evaluated first.
  newHost.get = (sheet, r, c) => {
    const s = newHost.sheetIndex(sheet);
    if (s >= 0) {
      const o = originOf(newTables[s]);
      const tr = r - o.row;
      const tc = c - o.col;
      if (tr >= 0 && tc >= 0) {
        const id = cellId(s, tr, tc);
        if (state.get(id) !== 2 && cone.has(id)) visit(id);
      }
    }
    return baseGet(sheet, r, c);
  };
  const visit = (id: number): void => {
    const st = state.get(id);
    if (st === 2) return;
    if (st === 1) throw new Unsupported('circular reference');
    if (depth > 400) throw new Unsupported('dependency chain too deep');
    state.set(id, 1);
    depth++;
    try {
      const formula = formulaAt('new', id);
      if (!formula) return;
      const ref = toAligned('new', id);
      const editedFormula = ref ? traces[ref.sheet].get(traceKey(ref))?.kind === 'formula' : false;
      const oid = ref ? oldIdOf(ref) : null;
      const oldFormula = oid !== null ? formulaAt('old', oid) : null;
      // Calibrate: the engine must reproduce the old file's value of the old formula.
      const calibrate = (): boolean => {
        try {
          const oldCached = cachedAt('old', oid!);
          return oldCached !== null && sameResult(evaluate(oldEngine, oldHost, 'old', oid!, oldFormula!).value, oldCached, diff.date1904?.old);
        } catch (e) {
          if (!(e instanceof Unsupported)) throw e;
          return false;
        }
      };
      let calibrated = false;
      if (oid !== null && oldFormula !== null) {
        // Copies of a formula that keeps reproducing the old file are trusted after a sample.
        // The same text at the same place has the new formula's shape (already compiled).
        let shapeText: string | null = null;
        try {
          const sameSpot = oldFormula === formula && sheetKey('old', oid) === sheetKey('new', id);
          shapeText = (sameSpot ? compileAt('new', id, formula) : compileAt('old', oid, oldFormula)).shape;
        } catch (e) {
          if (!(e instanceof Unsupported)) throw e;
        }
        const shape = `${idSheet(oid)}\u0000${shapeText}`;
        let stat = shapes.get(shape);
        if (!stat) shapes.set(shape, (stat = { passed: 0, failed: false }));
        if (stat.failed || stat.passed < CALIBRATION_SAMPLE) {
          calibrated = calibrate();
          if (!calibrated) {
            stat.failed = true;
            unmodelled++;
            return;
          }
          stat.passed++;
        }
      }
      let result: { value: Scalar; reads: Read[] };
      try {
        result = evaluate(newEngine, newHost, 'new', id, formula, true);
      } catch (e) {
        if (!(e instanceof Unsupported)) throw e;
        unmodelled++;
        return;
      }
      if (!sameResult(result.value, cachedAt('new', id), diff.date1904?.new)) {
        // Calling the file's value stale takes this very cell calibrated on
        // the same formula; a new or edited formula defers to the file.
        if (oldFormula === null || editedFormula || (!calibrated && !calibrate())) {
          unmodelled++;
          return;
        }
        stale.set(id, result.value);
        newHost.override(sheetKey('new', id), result.value);
      }
      modelled.add(id);
      reads.set(id, rectsOf(result.reads));
    } finally {
      depth--;
      state.set(id, 2);
    }
  };
  const sorted = topoOrder(queue, cone, staticNew);
  for (const id of sorted.cyclic) state.set(id, 2); // circular: references decide
  unmodelled += sorted.cyclic.length;
  for (const id of sorted.order) {
    if (truncated) break;
    try {
      visit(id);
    } catch (e) {
      if (!(e instanceof Unsupported)) throw e;
      state.set(id, 2); // references decide for this cell
    }
  }
  newHost.get = baseGet;
  newHost.getRange = baseGetRange;

  // 3. Stale cached values become changes (or annotate existing ones).
  for (const [id, value] of stale) {
    const ref = toAligned('new', id);
    if (!ref) continue;
    const oid = oldIdOf(ref);
    const oldValue = oid !== null ? cachedAt('old', oid) : null;
    const k = traceKey(ref);
    const info = { cached: toCellValue(cachedAt('new', id)), recalculated: toCellValue(value) };
    const t = traces[ref.sheet].get(k);
    if (t) t.stale = info;
    else if (!sameResult(oldValue, value, diff.date1904?.new)) traces[ref.sheet].set(k, { kind: 'computed', causes: [], causeCount: 0, impacts: [], impactCount: 0, stale: info });
  }

  // 4. The new-side graph from what trusted formulas actually read (an overlay on references).
  const graph = new OverlayGraph(staticNew, reads);
  const valueOf = (id: number): Scalar => (stale.has(id) ? stale.get(id)! : cachedAt('new', id));

  // Every sheet lines up cell for cell (no inserted/removed rows or columns, same origin).
  let layoutSame: boolean | null = null;
  const sameLayout = () =>
    (layoutSame ??= diff.sheets.every(({ diff: d }) => {
      const oo = originOf(d.oldTable);
      const no = originOf(d.newTable);
      if (oo.row !== no.row || oo.col !== no.col) return false;
      if (d.columns.some((c) => c.oldIndex !== c.newIndex)) return false;
      for (let r = 0; r < d.rowCount; r++) if (d.oldRows[r] !== d.newRows[r]) return false;
      return true;
    }));

  const butFor = (rootList: Root[]) => {
    const out = new Map<number, Map<number, ExactCauses>>();
    const record = (ref: CellRef) => {
      let m = out.get(ref.sheet);
      if (!m) out.set(ref.sheet, (m = new Map()));
      const k = traceKey(ref);
      let e = m.get(k);
      if (!e) m.set(k, (e = { causes: [], contributions: [], tainted: false }));
      return e;
    };
    const isComputed = (ref: CellRef | null): ref is CellRef => !!ref && traces[ref.sheet].get(traceKey(ref))?.kind === 'computed';
    rootList.forEach((root, ri) => {
      if (truncated || root.side !== 'new') return;
      // Without this root: an edited cell goes back to its old value; an added cell is empty.
      // Added and removed rows are causes as a whole (found by references), not cell by cell.
      if (root.trace.kind !== 'input' && root.trace.kind !== 'formula') return;
      const oid = oldIdOf(root.ref);
      let without: Scalar = oid !== null ? cachedAt('old', oid) : null;
      // An edited formula's counterfactual is the old formula on today's inputs, not its
      // old result (which also carries every other edit's effect). Its text is valid here
      // when no sheet's rows or columns moved; otherwise fall back to the old result.
      const oldFormula = oid !== null && root.trace.kind === 'formula' ? formulaAt('old', oid) : null;
      if (oldFormula && sameLayout()) {
        const s = idSheet(root.id);
        const o = originOf(newTables[s]);
        try {
          if (++evaluations > evalBudget) throw new Unsupported('evaluation budget');
          without = newEngine.evaluate(oldFormula, newHost.sheets[s], idRow(root.id) + o.row, idCol(root.id) + o.col, false).value;
        } catch (e) {
          if (!(e instanceof Unsupported)) throw e;
        }
      }
      const saved = new Map<number, Scalar | undefined>();
      const set = (k: number, v: Scalar) => {
        if (!saved.has(k)) saved.set(k, newHost.overrides.has(k) ? newHost.overrides.get(k) : undefined);
        newHost.override(k, v);
      };
      set(sheetKey('new', root.id), without);
      // Everything the root can reach through what formulas actually read, in dependency order.
      const reach = new Set<number>();
      const queue = [root.id];
      const deps: number[] = [];
      for (let qi = 0; qi < queue.length; qi++) {
        deps.length = 0;
        const src = queue[qi];
        graph.dependents(idSheet(src), idRow(src), idCol(src), deps);
        for (const d of deps) {
          if (reach.has(d) || d === root.id || !formulaAt('new', d)) continue;
          reach.add(d);
          queue.push(d);
        }
      }
      const { order, preds, cyclic } = topoOrder([root.id], reach, graph);
      const changed = new Set<number>([root.id]);
      const tainted = new Set<number>(cyclic);
      try {
        for (const d of order) {
          const p = preds.get(d) ?? [];
          if (p.some((x) => tainted.has(x)) || !modelled.has(d)) {
            // Can't recalculate it (or an input of it): whatever it feeds may or may not depend on this root.
            tainted.add(d);
            continue;
          }
          if (!p.some((x) => changed.has(x))) continue; // none of its inputs moved
          const v = evaluate(newEngine, newHost, 'new', d, formulaAt('new', d)!).value;
          const now = valueOf(d);
          if (sameResult(v, now, diff.date1904?.new)) continue; // unaffected: the effect stops here
          set(sheetKey('new', d), v);
          changed.add(d);
          const ref = toAligned('new', d);
          if (!isComputed(ref)) continue;
          const e = record(ref);
          e.causes.push(ri);
          e.contributions.push(typeof now === 'number' && typeof v === 'number' ? numericDelta(v, now)!.delta : null);
        }
        for (const d of tainted) {
          const ref = toAligned('new', d);
          if (isComputed(ref)) record(ref).tainted = true;
        }
      } catch (e) {
        if (!(e instanceof Unsupported)) throw e;
        truncated = true;
      } finally {
        for (const [k, v] of saved) newHost.override(k, v);
      }
    });
    return out;
  };

  return { graph, modelled, unmodelled, get truncated() { return truncated; }, butFor };
}

/** Engine value → table cell value (errors as their text). */
function toCellValue(v: Scalar): CellValue {
  return isError(v) ? v.code : v;
}

/**
 * The diff as it would be had the new file been recalculated: cells whose
 * saved values were stale (see `CellTrace.stale`) take their recalculated
 * values and count as changed. Alignment is unchanged, so `trace` still
 * applies to the result. Returns `diff` itself when nothing was stale.
 */
export function applyRecalculation(diff: WorkbookDiff, trace: WorkbookTrace): WorkbookDiff {
  if (!trace.stats.stale) return diff;
  const sheets = diff.sheets.map((s, si) => {
    const stale = [...trace.sheet(si)].filter(([, t]) => t.stale);
    if (!stale.length) return s;
    const d = s.diff;
    const w = d.columns.length;
    const rows = d.newTable.rows.slice();
    const status = d.status.slice();
    const changedCells = new Map(d.changedCells);
    const stats = { ...d.stats };
    const columns = d.columns.map((c) => ({ ...c }));
    for (const [k, t] of stale) {
      const row = Math.floor(k / w);
      const col = k % w;
      const ni = d.newRows[row];
      const ci = d.columns[col].newIndex;
      if (ni < 0 || ci < 0) continue;
      rows[ni] = rows[ni] === d.newTable.rows[ni] ? rows[ni].slice() : rows[ni];
      rows[ni][ci] = t.stale!.recalculated;
      if (status[row] === ROW_STATUS.equal) {
        status[row] = ROW_STATUS.modified;
        stats.rowsEqual--;
        stats.rowsModified++;
      }
      if (status[row] !== ROW_STATUS.modified) continue;
      const mask = changedCells.get(row) ? Uint8Array.from(changedCells.get(row)!) : new Uint8Array(w);
      if (!mask[col]) {
        mask[col] = 1;
        stats.cellsChanged++;
        columns[col].changes++;
      }
      changedCells.set(row, mask);
    }
    const newTable = { ...d.newTable, rows };
    return { ...s, status: s.status === 'equal' ? ('modified' as const) : s.status, diff: { ...d, newTable, columns, status, changedCells, stats } };
  });
  return { ...diff, sheets };
}
