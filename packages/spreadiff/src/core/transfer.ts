import type { CellRef, CellTrace, TraceStats, WorkbookTrace } from './trace.js';
import type { CellValue, Table, TableDiff, WorkbookDiff } from './types.js';

/**
 * Compact encoding of a {@link Table} for crossing a worker boundary.
 * Structured-cloning `CellValue[][]` allocates one object per cell; this
 * sends two typed arrays (transferable, zero-copy) and one joined string.
 * Measured at 1M rows × 12 columns: 11.2 s → 2.2 s end to end, most of
 * the remainder being the unpack on the receiving side.
 */
export interface PackedTable {
  /** `Table.name`. */
  name: string;
  /** `Table.columns`. */
  columns: string[];
  /** Number of columns. */
  width: number;
  /** Number of rows. */
  height: number;
  /** Cell types, row-major: 0 null, 1 string, 2 number, 3 true, 4 false. */
  types: Uint8Array;
  /** Numeric cell values, row-major (0 for non-numbers). */
  numbers: Float64Array;
  /** String cell values joined by U+0000, in row-major order. U+0000 inside values becomes U+FFFD. */
  strings: string;
  /**
   * `Table.formulas`, packed like the values (nested arrays of strings are
   * slow to structured-clone): `cells` holds `row * width + col` of each
   * formula, in order, and `text` the formulas joined by U+0000.
   * `lengths` is each row's length, -1 for a missing row.
   */
  formulas?: { lengths: Int32Array; width: number; cells: Float64Array; text: string };
  /** `Table.origin`, unchanged. */
  origin?: { row: number; col: number };
  /** `Table.header`, unchanged. */
  header?: Table['header'];
  /** `Table.arrays`, unchanged. */
  arrays?: Table['arrays'];
  /** `Table.names`, unchanged. */
  names?: Table['names'];
  /** `Table.merges`, unchanged. */
  merges?: Table['merges'];
}

const SEP = '\u0000';

/**
 * Encode a table for `postMessage`: typed arrays you can transfer plus one
 * joined string, instead of one object per cell. See {@link unpackTable}.
 */
export function packTable(t: Table): PackedTable {
  const width = t.columns.length;
  const height = t.rows.length;
  const types = new Uint8Array(width * height);
  const numbers = new Float64Array(width * height);
  const parts: string[] = [];
  for (let r = 0; r < height; r++) {
    const row = t.rows[r];
    for (let c = 0; c < width; c++) {
      const v = row[c];
      const i = r * width + c;
      if (v === null || v === undefined) continue;
      if (typeof v === 'string') {
        types[i] = 1;
        parts.push(v.includes(SEP) ? v.replaceAll(SEP, '�') : v);
      } else if (typeof v === 'number') {
        types[i] = 2;
        numbers[i] = v;
      } else types[i] = v ? 3 : 4;
    }
  }
  return { name: t.name, columns: t.columns, width, height, types, numbers, strings: parts.join(SEP), formulas: t.formulas && packFormulas(t.formulas), origin: t.origin, header: t.header, arrays: t.arrays, names: t.names, merges: t.merges };
}

function packFormulas(f: (string | null)[][]): NonNullable<PackedTable['formulas']> {
  let width = 0;
  let count = 0;
  for (const row of f) if (row) {
    width = Math.max(width, row.length);
    for (const x of row) if (x) count++;
  }
  const cells = new Float64Array(count);
  const lengths = new Int32Array(f.length).fill(-1);
  const text: string[] = [];
  let n = 0;
  for (let r = 0; r < f.length; r++) {
    const row = f[r];
    if (!row) continue;
    lengths[r] = row.length;
    for (let c = 0; c < row.length; c++) {
      const x = row[c];
      if (!x) continue;
      cells[n++] = r * width + c;
      text.push(x); // formulas never contain U+0000
    }
  }
  return { lengths, width, cells, text: text.join(SEP) };
}

function unpackFormulas(p: NonNullable<PackedTable['formulas']>): (string | null)[][] {
  const f: (string | null)[][] = new Array(p.lengths.length);
  for (let r = 0; r < p.lengths.length; r++) if (p.lengths[r] >= 0) f[r] = new Array(p.lengths[r]).fill(null);
  const text = p.cells.length ? p.text.split(SEP) : [];
  for (let i = 0; i < p.cells.length; i++) {
    const r = Math.floor(p.cells[i] / p.width);
    f[r][p.cells[i] - r * p.width] = text[i];
  }
  return f;
}

/** Decode a table encoded by {@link packTable}. */
export function unpackTable(p: PackedTable): Table {
  const strs = p.strings.length || p.types.includes(1) ? p.strings.split(SEP) : [];
  let s = 0;
  const rows: CellValue[][] = new Array(p.height);
  const { types, numbers, width } = p;
  for (let r = 0; r < p.height; r++) {
    const row: CellValue[] = new Array(width);
    const base = r * width;
    for (let c = 0; c < width; c++) {
      const k = types[base + c];
      row[c] = k === 0 ? null : k === 1 ? strs[s++] : k === 2 ? numbers[base + c] : k === 3;
    }
    rows[r] = row;
  }
  const t: Table = { name: p.name, columns: p.columns, rows };
  if (p.formulas) t.formulas = unpackFormulas(p.formulas);
  if (p.origin) t.origin = p.origin;
  if (p.header) t.header = p.header;
  if (p.arrays) t.arrays = p.arrays;
  if (p.names) t.names = p.names;
  if (p.merges) t.merges = p.merges;
  return t;
}

/** A {@link TableDiff} whose tables are packed with {@link packTable}. */
export type PackedTableDiff = Omit<TableDiff, 'oldTable' | 'newTable'> & { oldTable: PackedTable; newTable: PackedTable };
/** A {@link WorkbookDiff} with packed tables, from {@link packWorkbookDiff}. */
export interface PackedWorkbookDiff {
  /** Sheets, as in `WorkbookDiff.sheets`, with packed tables. */
  sheets: { name: string; status: WorkbookDiff['sheets'][number]['status']; diff: PackedTableDiff }[];
  /** Defined names, as in `WorkbookDiff.names`. */
  names?: WorkbookDiff['names'];
  /** Excel tables, as in `WorkbookDiff.tables`. */
  tables?: WorkbookDiff['tables'];
  /** Date systems, as in `WorkbookDiff.date1904`. */
  date1904?: WorkbookDiff['date1904'];
  /** As in `WorkbookDiff.computedValues`. */
  computedValues?: WorkbookDiff['computedValues'];
}

/** Pack a workbook diff; returns the message and the buffers to transfer. */
export function packWorkbookDiff(d: WorkbookDiff): { packed: PackedWorkbookDiff; transfer: ArrayBuffer[] } {
  const transfer: ArrayBuffer[] = [];
  const packed: PackedWorkbookDiff = {
    sheets: d.sheets.map((s) => {
      const oldTable = packTable(s.diff.oldTable);
      const newTable = packTable(s.diff.newTable);
      for (const p of [oldTable, newTable]) {
        transfer.push(p.types.buffer as ArrayBuffer, p.numbers.buffer as ArrayBuffer);
        if (p.formulas) transfer.push(p.formulas.cells.buffer as ArrayBuffer, p.formulas.lengths.buffer as ArrayBuffer);
      }
      transfer.push(s.diff.oldRows.buffer as ArrayBuffer, s.diff.newRows.buffer as ArrayBuffer, s.diff.status.buffer as ArrayBuffer, s.diff.moved.buffer as ArrayBuffer);
      return { name: s.name, status: s.status, diff: { ...s.diff, oldTable, newTable } };
    }),
    names: d.names,
    tables: d.tables,
    date1904: d.date1904,
    computedValues: d.computedValues,
  };
  return { packed, transfer };
}

/** Decode a workbook diff encoded by {@link packWorkbookDiff}. */
export function unpackWorkbookDiff(p: PackedWorkbookDiff): WorkbookDiff {
  return {
    sheets: p.sheets.map((s) => ({
      name: s.name,
      status: s.status,
      diff: { ...s.diff, oldTable: unpackTable(s.diff.oldTable), newTable: unpackTable(s.diff.newTable) },
    })),
    names: p.names,
    tables: p.tables,
    date1904: p.date1904,
    computedValues: p.computedValues,
  };
}

/** One sheet of a {@link PackedWorkbookTrace}: every `CellTrace` field in typed arrays. */
export interface PackedSheetTrace {
  /** Trace keys (`row * columns.length + col`), in `WorkbookTrace.sheet()` order. */
  keys: Float64Array;
  /** Per trace: kind (index into `input, formula, computed, added, removed`), method (0 none, 1 recalculated, 2 references), unexplained (0 none, 1 volatile, 2 dynamic, 3 unknown), and whether it has contributions. */
  codes: Uint8Array;
  /** Per trace: `causeCount` and `impactCount`. */
  counts: Uint32Array;
  /** Per trace `i`: its causes are `refs` from `causeAt[i]` up to `impactAt[i]`. */
  causeAt: Uint32Array;
  /** Per trace `i`: its impacts are `refs` from `impactAt[i]` up to `causeAt[i + 1]`. */
  impactAt: Uint32Array;
  /** Per trace `i`: its contributions are `contributions[contribAt[i]]` up to `contribAt[i + 1]`. */
  contribAt: Uint32Array;
  /** Refs as `sheet, row, col` triples. */
  refs: Int32Array;
  /** Contributions; NaN for null. */
  contributions: Float64Array;
  /** Stale values, by trace index. */
  stale: [number, NonNullable<CellTrace['stale']>][];
}

/**
 * A {@link WorkbookTrace} as plain data, from {@link packWorkbookTrace}:
 * what a worker sends back so tracing never runs on the main thread. Traces
 * travel as typed arrays: structured-cloning one object per cell costs the
 * receiving page about 200 ms per 40,000 traced cells.
 */
export interface PackedWorkbookTrace {
  /** Per sheet. */
  sheets: PackedSheetTrace[];
  /** Per sheet: `[key, parents]` for computed and formula-edit cells, the ones whose parents are asked for. */
  parents: [number, CellRef[]][][];
  /** `WorkbookTrace.stats`. */
  stats: TraceStats;
  /** `WorkbookTrace.truncated`. */
  truncated: boolean;
}

const KINDS: CellTrace['kind'][] = ['input', 'formula', 'computed', 'added', 'removed'];
const METHODS = [undefined, 'recalculated', 'references'] as const;
const UNEXPLAINED = [undefined, 'volatile', 'dynamic', 'unknown'] as const;

function packSheetTrace(traces: ReadonlyMap<number, CellTrace>): PackedSheetTrace {
  const n = traces.size;
  const keys = new Float64Array(n);
  const codes = new Uint8Array(n * 4);
  const counts = new Uint32Array(n * 2);
  const causeAt = new Uint32Array(n + 1);
  const impactAt = new Uint32Array(n + 1);
  const contribAt = new Uint32Array(n + 1);
  const refs: number[] = [];
  const contributions: number[] = [];
  const stale: PackedSheetTrace['stale'] = [];
  let i = 0;
  let nRefs = 0;
  for (const [k, t] of traces) {
    keys[i] = k;
    codes[i * 4] = KINDS.indexOf(t.kind);
    codes[i * 4 + 1] = METHODS.indexOf(t.method);
    codes[i * 4 + 2] = UNEXPLAINED.indexOf(t.unexplained);
    codes[i * 4 + 3] = t.contributions ? 1 : 0;
    counts[i * 2] = t.causeCount;
    counts[i * 2 + 1] = t.impactCount;
    causeAt[i] = nRefs;
    for (const c of t.causes) refs.push(c.sheet, c.row, c.col), nRefs++;
    impactAt[i] = nRefs;
    for (const c of t.impacts) refs.push(c.sheet, c.row, c.col), nRefs++;
    contribAt[i] = contributions.length;
    if (t.contributions) for (const x of t.contributions) contributions.push(x ?? NaN);
    if (t.stale) stale.push([i, t.stale]);
    i++;
  }
  causeAt[n] = impactAt[n] = nRefs;
  contribAt[n] = contributions.length;
  return { keys, codes, counts, causeAt, impactAt, contribAt, refs: Int32Array.from(refs), contributions: Float64Array.from(contributions), stale };
}

function unpackSheetTrace(p: PackedSheetTrace): Map<number, CellTrace> {
  const out = new Map<number, CellTrace>();
  const { refs } = p;
  const refList = (from: number, to: number) => {
    const list = new Array(to - from);
    for (let j = from; j < to; j++) list[j - from] = { sheet: refs[j * 3], row: refs[j * 3 + 1], col: refs[j * 3 + 2] };
    return list;
  };
  const stale = new Map(p.stale);
  for (let i = 0; i < p.keys.length; i++) {
    const t: CellTrace = {
      kind: KINDS[p.codes[i * 4]],
      causes: refList(p.causeAt[i], p.impactAt[i]),
      causeCount: p.counts[i * 2],
      impacts: refList(p.impactAt[i], p.causeAt[i + 1]),
      impactCount: p.counts[i * 2 + 1],
    };
    if (p.codes[i * 4 + 3]) {
      const c: (number | null)[] = [];
      for (let j = p.contribAt[i]; j < p.contribAt[i + 1]; j++) c.push(Number.isNaN(p.contributions[j]) ? null : p.contributions[j]);
      t.contributions = c;
    }
    const method = METHODS[p.codes[i * 4 + 1]];
    if (method) t.method = method;
    const unexplained = UNEXPLAINED[p.codes[i * 4 + 2]];
    if (unexplained) t.unexplained = unexplained;
    const st = stale.get(i);
    if (st) t.stale = st;
    out.set(p.keys[i], t);
  }
  return out;
}

/**
 * Pack a trace of `diff` for `postMessage`; transfer the buffers it returns.
 * Parents are computed here, on the packing side, for every computed and
 * formula-edit cell; after {@link unpackWorkbookTrace}, `parents()` of any
 * other cell is empty.
 */
export function packWorkbookTrace(trace: WorkbookTrace, diff: WorkbookDiff): { packed: PackedWorkbookTrace; transfer: ArrayBuffer[] } {
  const sheets: PackedSheetTrace[] = [];
  const parents: PackedWorkbookTrace['parents'] = [];
  const transfer: ArrayBuffer[] = [];
  diff.sheets.forEach((s, si) => {
    const width = s.diff.columns.length;
    const ps: [number, CellRef[]][] = [];
    for (const [k, t] of trace.sheet(si)) {
      if (t.kind !== 'computed' && t.kind !== 'formula') continue;
      const list = trace.parents({ sheet: si, row: Math.floor(k / width), col: k % width });
      if (list.length) ps.push([k, list]);
    }
    const packed = packSheetTrace(trace.sheet(si));
    for (const a of [packed.keys, packed.codes, packed.counts, packed.causeAt, packed.impactAt, packed.contribAt, packed.refs, packed.contributions]) transfer.push(a.buffer as ArrayBuffer);
    sheets.push(packed);
    parents.push(ps);
  });
  return { packed: { sheets, parents, stats: trace.stats, truncated: trace.truncated }, transfer };
}

/** Rebuild a {@link WorkbookTrace} from {@link packWorkbookTrace}'s output. Needs the (unpacked) diff it was made for. */
export function unpackWorkbookTrace(p: PackedWorkbookTrace, diff: WorkbookDiff): WorkbookTrace {
  const traces = p.sheets.map(unpackSheetTrace);
  const parents = p.parents.map((entries) => new Map(entries));
  const width = (sheet: number) => diff.sheets[sheet]?.diff.columns.length ?? 0;
  const empty = new Map<number, CellTrace>();
  return {
    get: (ref) => traces[ref.sheet]?.get(ref.row * width(ref.sheet) + ref.col),
    parents: (ref) => parents[ref.sheet]?.get(ref.row * width(ref.sheet) + ref.col) ?? [],
    sheet: (i) => traces[i] ?? empty,
    stats: p.stats,
    truncated: p.truncated,
  };
}
