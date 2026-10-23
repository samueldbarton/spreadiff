import { canonicalFormula } from './formula.js';
import { alignSequences } from './sequence.js';
import {
  ROW_STATUS,
  type AlignedColumn,
  type CellValue,
  type DiffOptions,
  type DiffStats,
  type SheetDiff,
  type Table,
  type TableDiff,
  type Workbook,
  type WorkbookDiff,
} from './types.js';

const DEFAULTS: Required<DiffOptions> = {
  keys: 'auto',
  numericTolerance: 0,
  trimWhitespace: true,
  ignoreCase: false,
  compareFormulas: true,
  rowSimilarity: 0.5,
  detectRenames: true,
};

/** How many candidate added rows a removed row is compared against when pairing. */
const PAIRING_WINDOW = 24;

type Normalizer = (v: CellValue) => string;

function makeNormalizer(o: Required<DiffOptions>): Normalizer {
  return (v) => {
    if (v === null || v === undefined) return '';
    if (typeof v === 'number') return String(v);
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    let s = v;
    if (o.trimWhitespace) s = s.trim();
    if (o.ignoreCase) s = s.toLowerCase();
    return s;
  };
}

function toNumber(v: CellValue): number {
  if (typeof v === 'number') return v;
  if (typeof v === 'string' && v.trim() !== '') return Number(v);
  return NaN;
}

/** Compare two cells under the diff options. Exported for renderers. */
export function cellsEqual(a: CellValue, b: CellValue, options: DiffOptions = {}): boolean {
  const o = { ...DEFAULTS, ...options };
  return makeComparer(o)(a, b);
}

function makeComparer(o: Required<DiffOptions>): (a: CellValue, b: CellValue) => boolean {
  const norm = makeNormalizer(o);
  const tol = o.numericTolerance;
  return (a, b) => {
    if (a === b) return true;
    if (tol > 0) {
      const x = toNumber(a);
      const y = toNumber(b);
      if (!Number.isNaN(x) && !Number.isNaN(y)) return Math.abs(x - y) <= tol;
    }
    return norm(a) === norm(b);
  };
}

/** Diff two tables: align columns, align rows, mark changed cells. */
export function diffTables(oldTable: Table, newTable: Table, options: DiffOptions = {}): TableDiff {
  const o: Required<DiffOptions> = { ...DEFAULTS, ...options };
  const aligned = alignTables(oldTable, newTable, o);
  return compareAligned(aligned, formulaKeyer([aligned], 0));
}

/** Rows and columns matched; cells not yet compared. */
interface Aligned {
  oldTable: Table;
  newTable: Table;
  columns: AlignedColumn[];
  pairs: number[];
  keyColumns: string[];
  o: Required<DiffOptions>;
}

function alignTables(oldTable: Table, newTable: Table, o: Required<DiffOptions>): Aligned {
  const norm = makeNormalizer(o);
  const equal = makeComparer(o);
  let columns = alignColumns(oldTable.columns, newTable.columns);
  const shared = columns.filter((c) => c.oldIndex >= 0 && c.newIndex >= 0);
  const keyColumns = resolveKeys(oldTable, newTable, shared, o, norm);
  const pairs: number[] = []; // [old, new, …]; -1 for the missing side
  if (keyColumns.length) alignByKey(oldTable, newTable, columns, keyColumns, norm, pairs);
  else alignBySequence(oldTable, newTable, shared, norm, equal, o.rowSimilarity, pairs);
  if (o.detectRenames) columns = detectRenames(columns, oldTable, newTable, pairs, equal);
  return { oldTable, newTable, columns, pairs, keyColumns, o };
}

type Side = 'old' | 'new';
/** Canonical form of a formula on one side, for "same meaning?" comparisons. */
type FormulaKeyer = (side: Side, formula: string) => string;

/**
 * Canonicalize references through the alignment: a reference becomes the
 * aligned row/column it points at, so `=B5*C5` (old) and `=B6*C6` (new,
 * after a row was inserted above) compare equal.
 */
function formulaKeyer(sheets: Aligned[], self: number): FormulaKeyer {
  type Axis = { map: Int32Array; next: Int32Array; prev: Int32Array };
  type SheetMap = { rows: Axis; cols: Axis; origin: { row: number; col: number } };
  // `next[i]` / `prev[i]`: nearest index ≥ / ≤ i that exists on both sides
  // (or -1). Lets a range that grew only by inserted rows compare equal.
  const axis = (map: Int32Array, survives: (i: number) => boolean): Axis => {
    const n = map.length;
    const next = new Int32Array(n + 1).fill(-1);
    const prev = new Int32Array(n).fill(-1);
    for (let i = n - 1; i >= 0; i--) next[i] = survives(i) ? i : next[i + 1];
    for (let i = 0, last = -1; i < n; i++) prev[i] = last = survives(i) ? i : last;
    return { map, next, prev };
  };
  const maps: Record<Side, SheetMap[]> = { old: [], new: [] };
  const byName: Record<Side, Map<string, number>> = { old: new Map(), new: new Map() };
  sheets.forEach((a, idx) => {
    for (const side of ['old', 'new'] as const) {
      const table = side === 'old' ? a.oldTable : a.newTable;
      const own = side === 'old' ? 0 : 1;
      const rows = new Int32Array(table.rows.length).fill(-1);
      const both = new Uint8Array(table.rows.length);
      for (let k = 0; k < a.pairs.length; k += 2) {
        const r = a.pairs[k + own];
        if (r < 0) continue;
        rows[r] = k / 2;
        both[r] = a.pairs[k + 1 - own] >= 0 ? 1 : 0;
      }
      const cols = new Int32Array(table.columns.length).fill(-1);
      const colBoth = new Uint8Array(table.columns.length);
      a.columns.forEach((c, ci) => {
        const j = side === 'old' ? c.oldIndex : c.newIndex;
        if (j < 0) return;
        cols[j] = ci;
        colBoth[j] = c.oldIndex >= 0 && c.newIndex >= 0 ? 1 : 0;
      });
      maps[side][idx] = {
        rows: axis(rows, (i) => both[i] === 1),
        cols: axis(cols, (j) => colBoth[j] === 1),
        origin: table.origin ?? { row: 1, col: 0 },
      };
      byName[side].set(table.name.toUpperCase(), idx);
    }
  });
  /** One axis of a reference, in aligned terms. Multi-cell spans snap inward to surviving rows. */
  const span = (v0: number, v1: number, origin: number, ax: Axis, tag: string) => {
    if (v0 < 0) return '*';
    const n = ax.map.length;
    const one = (i: number) => (i < 0 ? `${tag}h${i}` : i >= n ? `${tag}e${i - n}` : `${tag}${ax.map[i]}`);
    let i0 = v0 - origin;
    let i1 = v1 - origin;
    if (i1 > i0) {
      const a = i0 >= 0 && i0 < n ? ax.next[i0] : -1;
      const b = i1 >= 0 && i1 < n ? ax.prev[i1] : -1;
      if (a >= 0 && a <= i1) i0 = a;
      if (b >= 0 && b >= i0) i1 = b;
    }
    return `${one(i0)}:${one(i1)}`;
  };
  return (side, formula) =>
    canonicalFormula(formula, (rg) => {
      const idx = rg.sheet === null ? self : byName[side].get(rg.sheet.toUpperCase());
      if (idx === undefined) return `?${rg.sheet!.toUpperCase()}!${rg.r0},${rg.c0}:${rg.r1},${rg.c1}`;
      const m = maps[side][idx];
      return `${idx}!${span(rg.r0, rg.r1, m.origin.row, m.rows, 'r')}|${span(rg.c0, rg.c1, m.origin.col, m.cols, 'c')}`;
    });
}

function compareAligned(a: Aligned, keyer: FormulaKeyer): TableDiff {
  const { oldTable, newTable, columns, pairs, keyColumns, o } = a;
  const equal = makeComparer(o);
  const rowCount = pairs.length / 2;
  const oldRows = new Int32Array(rowCount);
  const newRows = new Int32Array(rowCount);
  const status = new Uint8Array(rowCount);
  const changedCells = new Map<number, Uint8Array>();
  const formulaChanged = new Map<number, Uint8Array>();
  const cmpCols = columns
    .map((c, i) => ({ c, i }))
    .filter(({ c }) => c.oldIndex >= 0 && c.newIndex >= 0);
  const of = oldTable.formulas;
  const nf = newTable.formulas;
  // Compare formulas only when both sources can carry them (XLSX vs. XLSX);
  // against a CSV every formula would otherwise look like a change.
  const checkFormulas = o.compareFormulas && !!of && !!nf;
  const stats: DiffStats = {
    rowsAdded: 0,
    rowsRemoved: 0,
    rowsModified: 0,
    rowsEqual: 0,
    cellsChanged: 0,
    columnsAdded: 0,
    columnsRemoved: 0,
    columnsRenamed: 0,
  };
  for (let r = 0; r < rowCount; r++) {
    const oi = pairs[2 * r];
    const ni = pairs[2 * r + 1];
    oldRows[r] = oi;
    newRows[r] = ni;
    if (oi < 0) {
      status[r] = ROW_STATUS.added;
      stats.rowsAdded++;
      continue;
    }
    if (ni < 0) {
      status[r] = ROW_STATUS.removed;
      stats.rowsRemoved++;
      continue;
    }
    const x = oldTable.rows[oi];
    const y = newTable.rows[ni];
    let mask: Uint8Array | null = null;
    for (const { c, i } of cmpCols) {
      const sameValue = equal(x[c.oldIndex] ?? null, y[c.newIndex] ?? null);
      let formulaDiffers = false;
      if (checkFormulas) {
        const fo = of![oi]?.[c.oldIndex] ?? null;
        const fn = nf![ni]?.[c.newIndex] ?? null;
        if (fo !== fn) formulaDiffers = fo === null || fn === null || keyer('old', fo) !== keyer('new', fn);
      }
      if (!sameValue || formulaDiffers) {
        mask ??= new Uint8Array(columns.length);
        mask[i] = 1;
        c.changes++;
        stats.cellsChanged++;
        if (formulaDiffers) {
          let fm = formulaChanged.get(r);
          if (!fm) formulaChanged.set(r, (fm = new Uint8Array(columns.length)));
          fm[i] = 1;
        }
      }
    }
    if (mask) {
      status[r] = ROW_STATUS.modified;
      changedCells.set(r, mask);
      stats.rowsModified++;
    } else {
      status[r] = ROW_STATUS.equal;
      stats.rowsEqual++;
    }
  }
  for (const c of columns) {
    if (c.status === 'added') stats.columnsAdded++;
    else if (c.status === 'removed') stats.columnsRemoved++;
    else if (c.status === 'renamed') stats.columnsRenamed++;
  }
  return { oldTable, newTable, columns, rowCount, oldRows, newRows, status, changedCells, formulaChanged, keyColumns, stats, options: o };
}

/** Diff every sheet of two workbooks, pairing sheets by name. */
export function diffWorkbooks(oldBook: Workbook, newBook: Workbook, options: DiffOptions = {}): WorkbookDiff {
  const o: Required<DiffOptions> = { ...DEFAULTS, ...options };
  const empty = (name: string): Table => ({ name, columns: [], rows: [] });
  const oldByName = new Map(oldBook.sheets.map((s) => [s.name, s]));
  const newNames = new Set(newBook.sheets.map((s) => s.name));
  // Single-sheet workbooks (e.g. CSV vs. CSV with different names) pair positionally.
  const positional = oldBook.sheets.length === 1 && newBook.sheets.length === 1;
  const plan: { name: string; status: SheetDiff['status'] | null; old: Table; new: Table }[] = [];
  for (const ns of newBook.sheets) {
    const os = positional ? oldBook.sheets[0] : oldByName.get(ns.name);
    plan.push({ name: ns.name, status: os ? null : 'added', old: os ?? empty(ns.name), new: ns });
  }
  if (!positional) {
    oldBook.sheets.forEach((os, i) => {
      if (newNames.has(os.name)) return;
      plan.splice(Math.min(i, plan.length), 0, { name: os.name, status: 'removed', old: os, new: empty(os.name) });
    });
  }
  // Align everything first, so formula references into other sheets can be
  // compared through those sheets' alignments.
  const aligned = plan.map((p) => alignTables(p.old, p.new, o));
  const sheets: SheetDiff[] = plan.map((p, i) => {
    const d = compareAligned(aligned[i], formulaKeyer(aligned, i));
    return { name: p.name, status: p.status ?? (hasChanges(d) ? 'modified' : 'equal'), diff: d };
  });
  return { sheets, names: oldBook.names || newBook.names ? { old: oldBook.names, new: newBook.names } : undefined };
}

export function hasChanges(d: TableDiff): boolean {
  const s = d.stats;
  return (
    s.rowsAdded + s.rowsRemoved + s.rowsModified + s.columnsAdded + s.columnsRemoved + s.columnsRenamed >
    0
  );
}

// ---------------------------------------------------------------------------

function alignColumns(oldCols: string[], newCols: string[]): AlignedColumn[] {
  // Duplicate names (possible in hand-built tables) match in order.
  const oldIndex = new Map<string, number[]>();
  oldCols.forEach((n, i) => {
    const q = oldIndex.get(n);
    if (q) q.push(i);
    else oldIndex.set(n, [i]);
  });
  const used = new Set<number>();
  const out: AlignedColumn[] = newCols.map((name, ni) => {
    const oi = oldIndex.get(name)?.shift();
    if (oi !== undefined) {
      used.add(oi);
      return { oldIndex: oi, newIndex: ni, name, status: 'equal', changes: 0 };
    }
    return { oldIndex: -1, newIndex: ni, name, status: 'added', changes: 0 };
  });
  // Insert removed columns right after their nearest surviving old predecessor.
  for (let oi = 0; oi < oldCols.length; oi++) {
    if (used.has(oi)) continue;
    let at = 0;
    for (let p = oi - 1; p >= 0; p--) {
      const pos = out.findIndex((c) => c.oldIndex === p);
      if (pos >= 0) {
        at = pos + 1;
        break;
      }
    }
    // Keep consecutive removed columns in their original order.
    while (at < out.length && out[at].status === 'removed' && out[at].oldIndex < oi) at++;
    out.splice(at, 0, { oldIndex: oi, newIndex: -1, name: oldCols[oi], status: 'removed', changes: 0 });
  }
  return out;
}

// "id", "ID", "Order ID", "order_id", "customerId" — but not "paid" or "valid".
const isKeyName = (name: string) =>
  /^(id|key|uuid|guid|sku|code)$/i.test(name) || /[ _.-](id|key|code)$/i.test(name) || /[a-z0-9](Id|ID)$/.test(name);

function resolveKeys(
  oldTable: Table,
  newTable: Table,
  shared: AlignedColumn[],
  o: Required<DiffOptions>,
  norm: Normalizer,
): string[] {
  if (o.keys === 'none') return [];
  if (Array.isArray(o.keys)) {
    const names = new Set(shared.map((c) => c.name));
    const missing = o.keys.filter((k) => !names.has(k));
    if (missing.length) throw new Error(`Key column(s) not in both tables: ${missing.join(', ')}`);
    return o.keys;
  }
  if (oldTable.rows.length < 2 || newTable.rows.length < 2) return [];
  // Auto: only id-like column names, leftmost first, with unique non-empty
  // values on both sides. A "name" column would turn renames into
  // remove + add, so plain data columns never become keys implicitly.
  const candidates = shared.filter((c) => isKeyName(c.name)).sort((x, y) => x.newIndex - y.newIndex);
  for (const c of candidates.slice(0, 4)) {
    if (isUniqueKey(oldTable, c.oldIndex, norm) && isUniqueKey(newTable, c.newIndex, norm)) {
      // A key that matches almost nothing (e.g. regenerated ids) is worse than none.
      const oldKeys = new Set(oldTable.rows.map((r) => norm(r[c.oldIndex] ?? null)));
      let hits = 0;
      for (const r of newTable.rows) if (oldKeys.has(norm(r[c.newIndex] ?? null))) hits++;
      if (hits >= 0.5 * Math.min(oldTable.rows.length, newTable.rows.length)) return [c.name];
    }
  }
  return [];
}

function isUniqueKey(t: Table, col: number, norm: Normalizer): boolean {
  const seen = new Set<string>();
  for (const r of t.rows) {
    const k = norm(r[col] ?? null);
    if (k === '' || seen.has(k)) return false;
    seen.add(k);
  }
  return true;
}

function alignByKey(
  oldTable: Table,
  newTable: Table,
  columns: AlignedColumn[],
  keys: string[],
  norm: Normalizer,
  pairs: number[],
): void {
  const keyCols = keys.map((k) => columns.find((c) => c.name === k)!);
  // Length-prefixed, so values containing the separator can't collide.
  const keyOf = (row: CellValue[], side: 'oldIndex' | 'newIndex') => {
    let k = '';
    for (const c of keyCols) {
      const v = norm(row[c[side]] ?? null);
      k += v.length + ':' + v;
    }
    return k;
  };
  const oldByKey = new Map<string, number[]>();
  oldTable.rows.forEach((r, i) => {
    const k = keyOf(r, 'oldIndex');
    const q = oldByKey.get(k);
    if (q) q.push(i);
    else oldByKey.set(k, [i]);
  });
  const matchedOld = new Int32Array(newTable.rows.length).fill(-1);
  const oldUsed = new Uint8Array(oldTable.rows.length);
  // Duplicate keys match in order. A cursor per key, not Array#shift (O(n)
  // per call: quadratic when a low-cardinality column is used as the key).
  const cursor = new Map<string, number>();
  newTable.rows.forEach((r, ni) => {
    const k = keyOf(r, 'newIndex');
    const q = oldByKey.get(k);
    const at = cursor.get(k) ?? 0;
    if (q && at < q.length) {
      const oi = q[at];
      cursor.set(k, at + 1);
      matchedOld[ni] = oi;
      oldUsed[oi] = 1;
    }
  });
  // Walk new order; slot each removed row in before the first new row whose
  // old counterpart comes after it, so deletions appear where they were.
  let nextOld = 0;
  const flushRemovedBefore = (limit: number) => {
    for (; nextOld < limit; nextOld++) if (!oldUsed[nextOld]) pairs.push(nextOld, -1);
  };
  // For added rows, look ahead to the next matched row so removals in the
  // same gap come first (the reading order of a code diff).
  const upcoming = new Int32Array(newTable.rows.length + 1);
  upcoming[newTable.rows.length] = oldTable.rows.length;
  for (let ni = newTable.rows.length - 1; ni >= 0; ni--) {
    upcoming[ni] = matchedOld[ni] >= 0 ? matchedOld[ni] : upcoming[ni + 1];
  }
  for (let ni = 0; ni < newTable.rows.length; ni++) {
    const oi = matchedOld[ni];
    if (oi >= nextOld || oi < 0) flushRemovedBefore(upcoming[ni]);
    pairs.push(oi, ni);
  }
  flushRemovedBefore(oldTable.rows.length);
}

function alignBySequence(
  oldTable: Table,
  newTable: Table,
  shared: AlignedColumn[],
  norm: Normalizer,
  equal: (a: CellValue, b: CellValue) => boolean,
  threshold: number,
  pairs: number[],
): void {
  // Intern whole rows to exact integer ids (no hash collisions).
  const ids = new Map<string, number>();
  const intern = (row: CellValue[], side: 'oldIndex' | 'newIndex') => {
    let s = '';
    for (const c of shared) {
      const v = norm(row[c[side]] ?? null);
      s += v.length + ':' + v; // length-prefixed: no separator collisions
    }
    let id = ids.get(s);
    if (id === undefined) ids.set(s, (id = ids.size));
    return id;
  };
  const a = Int32Array.from(oldTable.rows, (r) => intern(r, 'oldIndex'));
  const b = Int32Array.from(newTable.rows, (r) => intern(r, 'newIndex'));
  ids.clear();
  const matches = alignSequences(a, b);

  // Formula results change *because* inputs did, so a row's identity is its
  // constants: computed cells (formula on both sides) don't count against
  // similarity. Otherwise editing one price splits its row into remove + add.
  const of = oldTable.formulas;
  const nf = newTable.formulas;
  const similarity = (oi: number, ni: number) => {
    const x = oldTable.rows[oi];
    const y = newTable.rows[ni];
    let same = 0;
    let total = 0;
    for (const c of shared) {
      if (of?.[oi]?.[c.oldIndex] && nf?.[ni]?.[c.newIndex]) continue;
      total++;
      if (equal(x[c.oldIndex] ?? null, y[c.newIndex] ?? null)) same++;
    }
    if (total === 0) {
      // Only formulas: fall back to comparing results.
      for (const c of shared) {
        total++;
        if (equal(x[c.oldIndex] ?? null, y[c.newIndex] ?? null)) same++;
      }
    }
    return total ? same / total : 0;
  };

  // Pair up removed/added rows in each gap between matches as modifications.
  const gap = (o0: number, o1: number, n0: number, n1: number) => {
    let nj = n0;
    // One row out, one row in, at the same place: almost always an edit.
    // Pair it if anything at all still matches.
    const floor = o1 - o0 === 1 && n1 - n0 === 1 ? Math.min(threshold, Number.MIN_VALUE) : threshold;
    for (let oi = o0; oi < o1; oi++) {
      let best = -1;
      let bestScore = floor;
      const end = Math.min(n1, nj + PAIRING_WINDOW);
      for (let j = nj; j < end; j++) {
        const s = similarity(oi, j);
        if (s >= bestScore && (best < 0 || s > bestScore)) {
          best = j;
          bestScore = s;
          if (s === 1) break;
        }
      }
      if (best < 0) {
        pairs.push(oi, -1);
        continue;
      }
      for (; nj < best; nj++) pairs.push(-1, nj);
      pairs.push(oi, best);
      nj = best + 1;
    }
    for (; nj < n1; nj++) pairs.push(-1, nj);
  };

  let po = 0;
  let pn = 0;
  for (let k = 0; k < matches.length; k += 2) {
    const oi = matches[k];
    const ni = matches[k + 1];
    if (po < oi || pn < ni) gap(po, oi, pn, ni);
    pairs.push(oi, ni);
    po = oi + 1;
    pn = ni + 1;
  }
  gap(po, oldTable.rows.length, pn, newTable.rows.length);
}

function detectRenames(
  columns: AlignedColumn[],
  oldTable: Table,
  newTable: Table,
  pairs: number[],
  equal: (a: CellValue, b: CellValue) => boolean,
): AlignedColumn[] {
  const removed = columns.filter((c) => c.status === 'removed');
  const added = columns.filter((c) => c.status === 'added');
  if (!removed.length || !added.length) return columns;
  // Sample matched row pairs to keep this cheap on big tables.
  const matched: number[] = [];
  for (let k = 0; k < pairs.length; k += 2) if (pairs[k] >= 0 && pairs[k + 1] >= 0) matched.push(k);
  if (matched.length === 0) return columns;
  const step = Math.max(1, Math.floor(matched.length / 2000));
  const isBlank = (v: CellValue) => v === null || v === '';
  // Score only rows where either side has a value: two mostly-empty
  // columns are "equal" on their blanks and would otherwise pair up.
  const score = (rc: AlignedColumn, ac: AlignedColumn) => {
    let same = 0;
    let total = 0;
    for (let s = 0; s < matched.length; s += step) {
      const k = matched[s];
      const x = oldTable.rows[pairs[k]][rc.oldIndex] ?? null;
      const y = newTable.rows[pairs[k + 1]][ac.newIndex] ?? null;
      if (isBlank(x) && isBlank(y)) continue;
      total++;
      if (equal(x, y)) same++;
    }
    return total >= Math.min(3, matched.length) ? same / total : 0;
  };
  // Candidate order keeps this near-linear on wide sheets where every header
  // changed: the column in the same relative position, then columns whose
  // sampled contents hash identically, then (only for small sets) all of them.
  const fingerprint = (table: Table, col: number, side: 0 | 1) => {
    let h = '';
    for (let s = 0, n = 0; s < matched.length && n < 64; s += step, n++) {
      h += String(table.rows[pairs[matched[s] + side]][col] ?? '') + '\u0000';
    }
    return h;
  };
  const byPrint = new Map<string, AlignedColumn[]>();
  for (const ac of added) {
    const f = fingerprint(newTable, ac.newIndex, 1);
    const list = byPrint.get(f);
    if (list) list.push(ac);
    else byPrint.set(f, [ac]);
  }
  const exhaustive = removed.length * added.length <= 4096;
  const merged = new Set<AlignedColumn>();
  removed.forEach((rc, i) => {
    let best: AlignedColumn | null = null;
    let bestScore = 0.8;
    const consider = (ac: AlignedColumn | undefined) => {
      if (!ac || merged.has(ac)) return;
      const sc = score(rc, ac);
      if (sc >= bestScore && (!best || sc > bestScore)) {
        best = ac;
        bestScore = sc;
      }
    };
    consider(added[i]);
    if (bestScore < 1) for (const ac of byPrint.get(fingerprint(oldTable, rc.oldIndex, 0)) ?? []) consider(ac);
    if (bestScore < 1 && exhaustive) for (const ac of added) consider(ac);
    if (best) {
      const b = best as AlignedColumn;
      b.oldIndex = rc.oldIndex;
      b.oldName = rc.name;
      b.status = 'renamed';
      merged.add(b);
    }
  });
  const mergedOld = new Set([...merged].map((m) => m.oldIndex));
  return columns.filter((c) => !(c.status === 'removed' && mergedOld.has(c.oldIndex)));
}
