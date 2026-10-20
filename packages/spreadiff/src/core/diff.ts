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
    if (typeof v === 'number') return Number.isFinite(v) ? String(v) : String(v);
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
  const norm = makeNormalizer(o);
  const equal = makeComparer(o);

  let columns = alignColumns(oldTable.columns, newTable.columns);
  const shared = columns.filter((c) => c.oldIndex >= 0 && c.newIndex >= 0);

  // ---- Row alignment ---------------------------------------------------
  const keyColumns = resolveKeys(oldTable, newTable, shared, o, norm);
  const pairs: number[] = []; // [old, new, …]; -1 for the missing side
  if (keyColumns.length) {
    alignByKey(oldTable, newTable, columns, keyColumns, norm, pairs);
  } else {
    alignBySequence(oldTable, newTable, shared, norm, equal, o.rowSimilarity, pairs);
  }

  // ---- Column renames --------------------------------------------------
  if (o.detectRenames) columns = detectRenames(columns, oldTable, newTable, pairs, equal);

  // ---- Cell comparison -------------------------------------------------
  const rowCount = pairs.length / 2;
  const oldRows = new Int32Array(rowCount);
  const newRows = new Int32Array(rowCount);
  const status = new Uint8Array(rowCount);
  const changedCells = new Map<number, Uint8Array>();
  const cmpCols = columns
    .map((c, i) => ({ c, i }))
    .filter(({ c }) => c.oldIndex >= 0 && c.newIndex >= 0);
  const of = oldTable.formulas;
  const nf = newTable.formulas;
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
    const a = oldTable.rows[oi];
    const b = newTable.rows[ni];
    let mask: Uint8Array | null = null;
    for (const { c, i } of cmpCols) {
      let same = equal(a[c.oldIndex] ?? null, b[c.newIndex] ?? null);
      if (same && checkFormulas) {
        same = (of![oi]?.[c.oldIndex] ?? null) === (nf![ni]?.[c.newIndex] ?? null);
      }
      if (!same) {
        mask ??= new Uint8Array(columns.length);
        mask[i] = 1;
        c.changes++;
        stats.cellsChanged++;
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

  return {
    oldTable,
    newTable,
    columns,
    rowCount,
    oldRows,
    newRows,
    status,
    changedCells,
    keyColumns,
    stats,
    options: o,
  };
}

/** Diff every sheet of two workbooks, pairing sheets by name. */
export function diffWorkbooks(oldBook: Workbook, newBook: Workbook, options: DiffOptions = {}): WorkbookDiff {
  const sheets: SheetDiff[] = [];
  const oldByName = new Map(oldBook.sheets.map((s) => [s.name, s]));
  const newNames = new Set(newBook.sheets.map((s) => s.name));
  const empty = (name: string): Table => ({ name, columns: [], rows: [] });
  // Single-sheet workbooks (e.g. CSV vs. CSV with different names) pair positionally.
  const positional = oldBook.sheets.length === 1 && newBook.sheets.length === 1;
  for (const ns of newBook.sheets) {
    const os = positional ? oldBook.sheets[0] : oldByName.get(ns.name);
    if (!os) {
      const d = diffTables(empty(ns.name), ns, options);
      sheets.push({ name: ns.name, status: 'added', diff: d });
      continue;
    }
    const d = diffTables(os, ns, options);
    sheets.push({ name: ns.name, status: hasChanges(d) ? 'modified' : 'equal', diff: d });
  }
  if (!positional) {
    oldBook.sheets.forEach((os, i) => {
      if (newNames.has(os.name)) return;
      const d = diffTables(os, empty(os.name), options);
      sheets.splice(Math.min(i, sheets.length), 0, { name: os.name, status: 'removed', diff: d });
    });
  }
  return { sheets };
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
  const oldIndex = new Map<string, number>();
  oldCols.forEach((n, i) => oldIndex.set(n, i));
  const used = new Set<number>();
  const out: AlignedColumn[] = newCols.map((name, ni) => {
    const oi = oldIndex.get(name);
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

const KEY_NAME = /^(id|key|uuid|guid|sku|code|.*[ _-]id|.*Id)$/;

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
  const candidates = shared.filter((c) => KEY_NAME.test(c.name)).sort((x, y) => x.newIndex - y.newIndex);
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
  const keyOf = (row: CellValue[], side: 'oldIndex' | 'newIndex') =>
    keyCols.map((c) => norm(row[c[side]] ?? null)).join('\u0001');
  const oldByKey = new Map<string, number[]>();
  oldTable.rows.forEach((r, i) => {
    const k = keyOf(r, 'oldIndex');
    const q = oldByKey.get(k);
    if (q) q.push(i);
    else oldByKey.set(k, [i]);
  });
  const matchedOld = new Int32Array(newTable.rows.length).fill(-1);
  const oldUsed = new Uint8Array(oldTable.rows.length);
  newTable.rows.forEach((r, ni) => {
    const q = oldByKey.get(keyOf(r, 'newIndex'));
    if (q && q.length) {
      const oi = q.shift()!;
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
    for (const c of shared) s += norm(row[c[side]] ?? null) + '\u0001';
    let id = ids.get(s);
    if (id === undefined) ids.set(s, (id = ids.size));
    return id;
  };
  const a = Int32Array.from(oldTable.rows, (r) => intern(r, 'oldIndex'));
  const b = Int32Array.from(newTable.rows, (r) => intern(r, 'newIndex'));
  ids.clear();
  const matches = alignSequences(a, b);

  const similarity = (oi: number, ni: number) => {
    if (shared.length === 0) return 0;
    const x = oldTable.rows[oi];
    const y = newTable.rows[ni];
    let same = 0;
    for (const c of shared) if (equal(x[c.oldIndex] ?? null, y[c.newIndex] ?? null)) same++;
    return same / shared.length;
  };

  // Pair up removed/added rows in each gap between matches as modifications.
  const gap = (o0: number, o1: number, n0: number, n1: number) => {
    let nj = n0;
    for (let oi = o0; oi < o1; oi++) {
      let best = -1;
      let bestScore = threshold;
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
  const merged = new Set<AlignedColumn>();
  for (const rc of removed) {
    let best: AlignedColumn | null = null;
    let bestScore = 0.8;
    for (const ac of added) {
      if (merged.has(ac)) continue;
      let same = 0;
      let total = 0;
      let nonEmpty = 0;
      for (let s = 0; s < matched.length; s += step) {
        const k = matched[s];
        const x = oldTable.rows[pairs[k]][rc.oldIndex] ?? null;
        const y = newTable.rows[pairs[k + 1]][ac.newIndex] ?? null;
        total++;
        if (x !== null && x !== '') nonEmpty++;
        if (equal(x, y)) same++;
      }
      const score = same / total;
      if (nonEmpty > 0 && score >= bestScore) {
        best = ac;
        bestScore = score;
      }
    }
    if (best) {
      best.oldIndex = rc.oldIndex;
      best.oldName = rc.name;
      best.status = 'renamed';
      merged.add(best);
    }
  }
  return columns.filter((c) => !(c.status === 'removed' && [...merged].some((m) => m.oldIndex === c.oldIndex)));
}
