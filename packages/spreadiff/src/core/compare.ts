import { diffWorkbooks } from './diff.js';
import { parseFile, type ParseFileOptions } from './file.js';
import { reportOf, type ChangeReport } from './report.js';
import { traceWorkbook, type TraceContext, type TraceOptions, type WorkbookTrace } from './trace.js';
import type { DiffOptions, SheetDiff, Table, Workbook, WorkbookDiff } from './types.js';
import type { DiffWorker } from '../worker/index.js';

/**
 * Something {@link diffFiles} can read:
 * - a `File` or `Blob` (from an `<input>`, a drop, or `fetch(...).blob()`);
 * - a `URL`, fetched;
 * - `{ name, data }`: a file name (it picks the parser) and its bytes, or its text for CSV;
 * - raw bytes, with the format sniffed from the contents;
 * - a `Workbook` or `Table` you already parsed.
 */
export type SpreadsheetInput =
  | Blob
  | URL
  | ArrayBuffer
  | Uint8Array
  | { name: string; data: ArrayBuffer | Uint8Array | string }
  | Workbook
  | Table;

/** One sheet of a {@link Comparison}, ready for `DiffGrid`: its diff with its trace attached. */
export interface SheetView extends SheetDiff {
  /** Position in `Comparison.sheets` (and `Comparison.workbook.sheets`). */
  index: number;
  /** The trace, bound to this sheet. Undefined when nothing was traced. */
  trace: TraceContext | undefined;
}

/** The result of {@link compareWorkbooks} and {@link diffFiles}. */
export interface Comparison {
  /** Old and new file names. */
  names: { old: string; new: string };
  /** The workbook diff, from `diffWorkbooks`. */
  workbook: WorkbookDiff;
  /** The trace, from `traceWorkbook`; null when tracing was off or neither file has formulas. */
  trace: WorkbookTrace | null;
  /** One entry per sheet, in `workbook.sheets` order. Pass one to `DiffGrid` as `sheet`. */
  sheets: SheetView[];
  /** The first sheet with a change (0 when none changed): a good one to show first. */
  firstChanged: number;
  /**
   * What changed, cell by cell: edits, downstream results with their causes
   * and contributions, final outputs, and changes of shape. Built on first
   * call, then cached. Needs a trace; without one, every change is an edit.
   */
  report(): ChangeReport;
  /** Milliseconds spent parsing, diffing and tracing. */
  timings: { parse: number; diff: number; trace: number };
}

/** Options for {@link compareWorkbooks} and {@link diffFiles}. All are optional. */
export interface CompareOptions {
  /** How to match rows and compare cells. */
  diff?: DiffOptions;
  /** How to read files (`header`, `delimiter`). Ignored for parsed inputs. */
  parse?: ParseFileOptions;
  /**
   * Trace downstream changes: `true`, `false`, `TraceOptions`, or `'auto'`
   * (the default) to trace when either side has formulas.
   */
  trace?: boolean | 'auto' | TraceOptions;
}

/** Options for {@link diffFiles}. */
export interface DiffFilesOptions extends CompareOptions {
  /**
   * Do the work in a Web Worker, so a large file never freezes the page:
   * a client from `createDiffWorker`, or a function that creates the
   * `Worker` (it's created on first use and kept). Without it, everything
   * runs on the calling thread. Parsed inputs (`Workbook`, `Table`) always
   * run on the calling thread.
   *
   * ```ts
   * worker: () => new Worker(new URL('spreadiff/worker/worker.js', import.meta.url), { type: 'module' })
   * ```
   */
  worker?: DiffWorker | (() => Worker);
  /** Abort fetching URL inputs. */
  signal?: AbortSignal;
}

const isTable = (x: unknown): x is Table => !!x && typeof x === 'object' && 'columns' in x && 'rows' in x && Array.isArray((x as Table).rows);
const isWorkbook = (x: unknown): x is Workbook => !!x && typeof x === 'object' && 'sheets' in x && Array.isArray((x as Workbook).sheets);
const hasFormulas = (wb: Workbook) => wb.sheets.some((t) => t.formulas?.some((r) => r?.some((f) => f)));
const now = () => (typeof performance !== 'undefined' ? performance.now() : Date.now());

/** A sheet view for each sheet, with the trace bound to it. */
function viewsOf(workbook: WorkbookDiff, trace: WorkbookTrace | null): SheetView[] {
  return workbook.sheets.map((s, index) => ({ ...s, index, trace: trace ? { trace, workbook, sheet: index } : undefined }));
}

function comparisonOf(workbook: WorkbookDiff, trace: WorkbookTrace | null, names: Comparison['names'], timings: Comparison['timings']): Comparison {
  const sheets = viewsOf(workbook, trace);
  let report: ChangeReport | null = null;
  const first = sheets.findIndex((s) => s.status !== 'equal');
  return {
    names,
    workbook,
    trace,
    sheets,
    firstChanged: first < 0 ? 0 : first,
    report: () => (report ??= reportOf(workbook, trace ?? traceWorkbook(workbook, { recalculate: false }), names.old, names.new)),
    timings,
  };
}

const traceOptions = (t: CompareOptions['trace'], a: Workbook, b: Workbook): TraceOptions | null => {
  if (t === false) return null;
  if (t === undefined || t === 'auto') return hasFormulas(a) || hasFormulas(b) ? {} : null;
  return t === true ? {} : t;
};

/**
 * Diff and trace two parsed workbooks, synchronously. Each sheet of the
 * result carries its own trace, so it can go straight to a grid:
 *
 * ```ts
 * import { compareWorkbooks, DiffGrid, parseFile } from 'spreadiff';
 * declare const oldBytes: Uint8Array, newBytes: Uint8Array, container: HTMLElement;
 *
 * const cmp = compareWorkbooks(parseFile('v1.xlsx', oldBytes), parseFile('v2.xlsx', newBytes));
 * new DiffGrid().render({ sheet: cmp.sheets[cmp.firstChanged], container });
 * ```
 */
export function compareWorkbooks(a: Workbook | Table, b: Workbook | Table, options: CompareOptions = {}): Comparison {
  const wa = toWorkbook(a);
  const wb = toWorkbook(b);
  const t0 = now();
  const workbook = diffWorkbooks(wa, wb, options.diff);
  const t1 = now();
  const to = traceOptions(options.trace, wa, wb);
  const trace = to ? traceWorkbook(workbook, to) : null;
  return comparisonOf(workbook, trace, { old: wa.name, new: wb.name }, { parse: 0, diff: t1 - t0, trace: now() - t1 });
}

const toWorkbook = (x: Workbook | Table): Workbook => (isWorkbook(x) ? x : { name: x.name, sheets: [x] });

/**
 * Read an input's name and bytes: fetches a URL, reads a Blob, encodes
 * text. Parsed inputs aren't accepted. The bytes are a copy.
 */
export async function readSpreadsheetInput(input: Exclude<SpreadsheetInput, Workbook | Table>, fallback: string, signal?: AbortSignal): Promise<{ name: string; data: ArrayBuffer }> {
  const copy = (u: Uint8Array) => u.slice().buffer as ArrayBuffer; // its own buffer: the worker may take it
  if (input instanceof URL) {
    const res = await fetch(input, { signal });
    if (!res.ok) throw new Error(`${input}: HTTP ${res.status}`);
    const name = decodeURIComponent(input.pathname.split('/').pop() || fallback);
    return { name, data: await res.arrayBuffer() };
  }
  if (typeof Blob !== 'undefined' && input instanceof Blob) {
    const name = (input as File).name || fallback;
    return { name, data: await input.arrayBuffer() };
  }
  if (input instanceof ArrayBuffer) return { name: fallback, data: input.slice(0) };
  if (input instanceof Uint8Array) return { name: fallback, data: copy(input) };
  const { name, data } = input as { name: string; data: ArrayBuffer | Uint8Array | string };
  if (typeof data === 'string') return { name, data: copy(new TextEncoder().encode(data)) };
  return { name, data: data instanceof Uint8Array ? copy(data) : data.slice(0) };
}

const workers = new WeakMap<() => Worker, DiffWorker>();
let createClient: ((w: Worker) => DiffWorker) | null = null;

/**
 * Diff two spreadsheets from whatever you have: `File`s, `Blob`s, URLs,
 * bytes or text. Picks the parser from the contents (XLSX, XLSB, XLS, ODS,
 * CSV, TSV), diffs every sheet and, when there are formulas, traces
 * downstream changes. Each sheet of the result carries its trace, so it
 * goes straight to a grid.
 *
 * ```ts
 * import { diffFiles, DiffGrid } from 'spreadiff';
 * declare const container: HTMLElement;
 *
 * const cmp = await diffFiles(new URL('/v1.xlsx', location.href), new URL('/v2.xlsx', location.href));
 * new DiffGrid().render({ sheet: cmp.sheets[cmp.firstChanged], container });
 * ```
 *
 * Pass `worker` to do the work off the main thread. The `SpreadsheetDiff`
 * component calls this for you.
 */
export async function diffFiles(oldInput: SpreadsheetInput, newInput: SpreadsheetInput, options: DiffFilesOptions = {}): Promise<Comparison> {
  const parsed = (x: SpreadsheetInput): x is Workbook | Table => isWorkbook(x) || isTable(x);
  if (parsed(oldInput) && parsed(newInput)) return compareWorkbooks(oldInput, newInput, options);
  const t0 = now();
  const toParsed = async (x: SpreadsheetInput, fallback: string): Promise<Workbook> => {
    if (parsed(x)) return toWorkbook(x);
    const { name, data } = await readSpreadsheetInput(x, fallback, options.signal);
    return parseFile(name, data, { header: 'auto', ...options.parse });
  };
  if (options.worker && !parsed(oldInput) && !parsed(newInput)) {
    const [a, b] = await Promise.all([readSpreadsheetInput(oldInput, 'old', options.signal), readSpreadsheetInput(newInput, 'new', options.signal)]);
    let client: DiffWorker;
    if (typeof options.worker === 'function') {
      const factory = options.worker;
      client = workers.get(factory) ?? (await workerClient(factory));
    } else client = options.worker;
    // 'auto' needs to know whether there are formulas; the worker decides from what it parses.
    const trace = options.trace === undefined || options.trace === 'auto' ? 'auto' : options.trace;
    const r = await client.diffBuffers(a, b, options.diff, { header: options.parse?.header ?? 'auto', trace });
    return comparisonOf(r.result, r.trace ?? null, { old: a.name, new: b.name }, r.timings);
  }
  const [a, b] = await Promise.all([toParsed(oldInput, 'old'), toParsed(newInput, 'new')]);
  const parse = now() - t0;
  const cmp = compareWorkbooks(a, b, options);
  cmp.timings.parse = parse;
  return cmp;
}

async function workerClient(factory: () => Worker): Promise<DiffWorker> {
  createClient ??= (await import('../worker/index.js')).createDiffWorker;
  const client = createClient(factory());
  workers.set(factory, client);
  return client;
}
