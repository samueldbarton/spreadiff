import {
  columnLabel,
  diffWorkbooks,
  formatLabel,
  numericDelta,
  parseFile,
  ROW_STATUS,
  tableLabeler,
  traceWorkbook,
  type CellLabel,
  type CellRef,
  type CellTrace,
  type CellValue,
  type DiffOptions,
  type ExternalSource,
  type ParseFileOptions,
  type TableDiff,
  type TraceOptions,
  type Workbook,
  type WorkbookDiff,
  type WorkbookTrace,
} from 'spreadiff/core';

/** A file to compare: its name (which picks the parser) and its bytes, or an already-parsed workbook. */
export type FileInput = Workbook | { name: string; data: ArrayBuffer | Uint8Array };

/** Options for {@link compare}. All are optional. */
export interface CompareOptions {
  /** How to match rows and compare cells (`keys`, `numericTolerance`, …). Default: spreadiff's defaults. */
  diff?: DiffOptions;
  /** How to read files (`header`, `delimiter`). Ignored for workbooks passed in parsed. */
  parse?: ParseFileOptions;
  /**
   * Tracing options. `recalculate: false` follows references only (faster,
   * may over-attribute); `liveFunctions` adds data-feed functions to treat
   * as inputs. See spreadiff's `TraceOptions`.
   */
  trace?: TraceOptions;
  /** Only these sheets (by name). */
  sheets?: string[];
}

/** One changed cell in a {@link ChangeReport}. */
export interface CellChange {
  /** `Sheet!A1` in the new file (the old file for removed cells). */
  id: string;
  /** Sheet name. */
  sheet: string;
  /** A1 address. */
  cell: string;
  /** What the sheet calls this cell: `'India › Revenue · FY2025'`. May be empty. */
  label: string;
  /** The parts of `label`. */
  labels: CellLabel;
  /**
   * - `input`: a value was typed over.
   * - `formula`: the formula was edited.
   * - `external`: a formula reading outside data (Bloomberg, RTD, another workbook) returned a new value. Not recalculated.
   * - `computed`: same formula, new result, because something upstream changed.
   */
  kind: 'input' | 'formula' | 'external' | 'computed';
  /** Value in the old file. */
  before: CellValue;
  /** Value in the new file (recalculated, when the file's own was stale). */
  after: CellValue;
  /** Formula text on each side, when either has one. */
  formula?: { before: string | null; after: string | null };
  /** `after − before`, when both are numbers. */
  delta?: number;
  /** `delta / |before|`, when both are numbers and `before` isn't 0. */
  ratio?: number;
  /** For `external`: where the value comes from. */
  external?: ExternalSource;
  /** For edits: how many downstream results this edit moved. */
  impact?: number;
  /** For `computed`: the edits that moved it, largest contribution first (at most 25; see `causeCount`). */
  causes?: Cause[];
  /** For `computed`: exact number of causes. */
  causeCount?: number;
  /** For `computed`: `recalculated` (exact) or `references` (by formula references: may include a cause that didn't matter). */
  method?: 'recalculated' | 'references';
  /** For `computed`: nothing else that changed depends on it (a final output, not an intermediate step). */
  output?: boolean;
  /** For `computed` with no traced cause: why. */
  unexplained?: 'volatile' | 'dynamic' | 'unknown';
  /** The new file's saved value was stale: it was saved without recalculating. `after` is the recalculated value. */
  stale?: { saved: CellValue };
}

/** One cause of a downstream change. */
export interface Cause {
  /** `Sheet!A1` of the edit. */
  id: string;
  /** Its label. */
  label: string;
  /** How much of the change it accounts for: the result now minus the result without that edit. Null when not recalculated or not numeric. */
  contribution: number | null;
}

/** A change to a sheet's shape. */
export interface StructureChange {
  /** Sheet name. */
  sheet: string;
  /** What changed. */
  kind: 'sheet-added' | 'sheet-removed' | 'rows-added' | 'rows-removed' | 'rows-moved' | 'column-added' | 'column-removed' | 'column-renamed';
  /** Spreadsheet rows (1-based, new file; old file for removed), for row changes. */
  rows?: number[];
  /** Row labels, parallel to `rows` (when known). */
  labels?: string[];
  /** Column name, for column changes (`from → to` for renames). */
  column?: string;
  /** Downstream results these rows or columns moved (formulas whose ranges cover them). */
  impact?: number;
}

/** Everything that changed between two workbooks, from {@link compare}. */
export interface ChangeReport {
  /** Old file name. */
  old: string;
  /** New file name. */
  new: string;
  /** Counts. */
  totals: {
    /** Cells typed over. */
    inputs: number;
    /** Formulas edited. */
    formulas: number;
    /** Outside-data values that changed. */
    external: number;
    /** Results that moved downstream. */
    downstream: number;
    /** Downstream results with no other changed cell depending on them. */
    outputs: number;
    /** Downstream results whose causes were found by recalculating. */
    recalculated: number;
    /** Downstream results with no cause found. */
    unexplained: number;
    /** Results the new file saved without recalculating. */
    stale: number;
  };
  /** Input, formula and external-data changes, in sheet order. */
  edits: CellChange[];
  /** Results that moved because of the edits: final outputs first, then sheet order. */
  downstream: CellChange[];
  /** Added, removed and moved rows, columns and sheets. */
  structure: StructureChange[];
  /** Caveats a reader should know about (stale files, unmodelled formulas, truncation). */
  notes: string[];
  /** The underlying diff, for callers that want more than the report. Not enumerable, so `JSON.stringify` leaves it out. */
  diff: WorkbookDiff;
  /** The underlying trace. Not enumerable. */
  trace: WorkbookTrace;
}

const isWorkbook = (f: FileInput): f is Workbook => 'sheets' in f;
const toWorkbook = (f: FileInput, options: CompareOptions): Workbook => {
  const wb = isWorkbook(f) ? f : parseFile(f.name, f.data, { header: 'auto', ...options.parse });
  return options.sheets ? { ...wb, sheets: wb.sheets.filter((s) => options.sheets!.includes(s.name)) } : wb;
};

/**
 * Compare two workbooks: what was edited, what moved downstream because of
 * it (and by how much each edit contributed), and what changed shape.
 * Formulas are recalculated with spreadiff's engine; formulas that read
 * outside data (Bloomberg, Capital IQ, RTD, other workbooks) are taken as
 * saved and reported as `external` edits when they changed.
 *
 * ```ts
 * import { readFileSync } from 'node:fs';
 * import { compare, toMarkdown } from 'spreadiff-lite';
 *
 * const report = compare(
 *   { name: 'model-v1.xlsx', data: readFileSync('model-v1.xlsx') },
 *   { name: 'model-v2.xlsx', data: readFileSync('model-v2.xlsx') },
 * );
 * console.log(toMarkdown(report));
 * ```
 */
export function compare(oldFile: FileInput, newFile: FileInput, options: CompareOptions = {}): ChangeReport {
  const a = toWorkbook(oldFile, options);
  const b = toWorkbook(newFile, options);
  const diff = diffWorkbooks(a, b, options.diff);
  return reportOf(diff, traceWorkbook(diff, options.trace), a.name, b.name);
}

/** The report for a diff and trace you already have (from `diffWorkbooks` and `traceWorkbook`). */
export function reportOf(diff: WorkbookDiff, trace: WorkbookTrace, oldName = 'old', newName = 'new'): ChangeReport {
  const labelers = diff.sheets.map((s) => ({
    old: tableLabeler(s.diff.oldTable, { keyColumns: s.diff.keyColumns }),
    new: tableLabeler(s.diff.newTable, { keyColumns: s.diff.keyColumns }),
  }));
  const side = (d: TableDiff, row: number): 'new' | 'old' => (d.newRows[row] >= 0 ? 'new' : 'old');
  /** Address and labels of an aligned cell. */
  const locate = (ref: CellRef) => {
    const s = diff.sheets[ref.sheet];
    const d = s.diff;
    const sd = side(d, ref.row);
    const col = d.columns[ref.col];
    const t = sd === 'new' ? d.newTable : d.oldTable;
    const tr = sd === 'new' ? d.newRows[ref.row] : d.oldRows[ref.row];
    const tc = sd === 'new' ? col.newIndex : col.oldIndex;
    const o = t.origin ?? { row: 1, col: 0 };
    const cell = tc >= 0 ? `${columnLabel(tc + o.col)}${tr + o.row + 1}` : `?${tr + o.row + 1}`;
    const labels = tc >= 0 ? labelers[ref.sheet][sd](tr, tc) : { sections: [], row: [], column: [col.name] };
    return { sheet: s.name, cell, id: `${s.name}!${cell}`, labels, label: formatLabel(labels) };
  };
  const valueOf = (d: TableDiff, row: number, col: number, sd: 'old' | 'new'): CellValue => {
    const r = sd === 'new' ? d.newRows[row] : d.oldRows[row];
    const c = sd === 'new' ? d.columns[col].newIndex : d.columns[col].oldIndex;
    return r >= 0 && c >= 0 ? (sd === 'new' ? d.newTable : d.oldTable).rows[r][c] : null;
  };
  const formulaOf = (d: TableDiff, row: number, col: number, sd: 'old' | 'new'): string | null => {
    const r = sd === 'new' ? d.newRows[row] : d.oldRows[row];
    const c = sd === 'new' ? d.columns[col].newIndex : d.columns[col].oldIndex;
    return r >= 0 && c >= 0 ? ((sd === 'new' ? d.newTable : d.oldTable).formulas?.[r]?.[c] ?? null) : null;
  };

  const edits: CellChange[] = [];
  const downstream: CellChange[] = [];
  const intermediate = new Set<string>();
  const computedRefs: CellRef[] = [];
  diff.sheets.forEach((s, si) => {
    const d = s.diff;
    const w = d.columns.length;
    const keys = [...trace.sheet(si).keys()].sort((x, y) => x - y);
    for (const k of keys) {
      const t = trace.sheet(si).get(k)!;
      if (t.kind === 'added' || t.kind === 'removed') continue; // reported as structure
      const ref = { sheet: si, row: Math.floor(k / w), col: k % w };
      const change = cellChange(ref, t);
      if (t.kind === 'computed') {
        downstream.push(change);
        computedRefs.push(ref);
      } else edits.push(change);
    }
  });
  function cellChange(ref: CellRef, t: CellTrace): CellChange {
    const d = diff.sheets[ref.sheet].diff;
    const before = valueOf(d, ref.row, ref.col, 'old');
    const after = t.stale ? t.stale.recalculated : valueOf(d, ref.row, ref.col, 'new');
    const fb = formulaOf(d, ref.row, ref.col, 'old');
    const fa = formulaOf(d, ref.row, ref.col, 'new');
    const kind: CellChange['kind'] = t.external ? 'external' : t.kind === 'computed' ? 'computed' : t.kind === 'formula' ? 'formula' : 'input';
    const c: CellChange = { ...locate(ref), kind, before, after };
    if (fb || fa) c.formula = { before: fb, after: fa };
    const nd = numericDelta(before, after);
    if (nd && typeof before === 'number' && typeof after === 'number') {
      c.delta = nd.delta;
      if (nd.ratio !== null) c.ratio = Number(nd.ratio.toPrecision(15)); // 0.2, not 0.19999999999999998
    }
    if (t.external) c.external = t.external;
    if (t.stale) c.stale = { saved: t.stale.cached };
    if (kind !== 'computed') {
      c.impact = t.impactCount;
      return c;
    }
    c.causeCount = t.causeCount;
    if (t.method) c.method = t.method;
    if (t.unexplained) c.unexplained = t.unexplained;
    const causes = t.causes.map((r, i) => ({ id: locate(r).id, label: locate(r).label, contribution: t.contributions?.[i] ?? null }));
    causes.sort((x, y) => Math.abs(y.contribution ?? 0) - Math.abs(x.contribution ?? 0));
    c.causes = causes;
    return c;
  }
  // Final outputs: computed cells no other changed formula reads.
  if (computedRefs.length <= 50_000) {
    for (const ref of computedRefs) for (const p of trace.parents(ref)) intermediate.add(`${p.sheet}:${p.row}:${p.col}`);
    computedRefs.forEach((ref, i) => (downstream[i].output = !intermediate.has(`${ref.sheet}:${ref.row}:${ref.col}`)));
    const order = downstream.map((c, i) => ({ c, i }));
    order.sort((x, y) => Number(!!y.c.output) - Number(!!x.c.output) || x.i - y.i);
    downstream.splice(0, downstream.length, ...order.map((o) => o.c));
  }

  const structure: StructureChange[] = [];
  diff.sheets.forEach((s, si) => {
    const d = s.diff;
    if (s.status === 'added' || s.status === 'removed') {
      structure.push({ sheet: s.name, kind: s.status === 'added' ? 'sheet-added' : 'sheet-removed' });
      return;
    }
    // Impact of added/removed cells, per aligned row and column.
    const rowImpact = new Map<number, number>();
    const colImpact = new Map<number, number>();
    const w = d.columns.length;
    for (const [k, t] of trace.sheet(si)) {
      if (t.kind !== 'added' && t.kind !== 'removed') continue;
      const row = Math.floor(k / w);
      const col = k % w;
      const st = d.columns[col].status;
      if (st === 'added' || st === 'removed') colImpact.set(col, (colImpact.get(col) ?? 0) + t.impactCount);
      else rowImpact.set(row, (rowImpact.get(row) ?? 0) + t.impactCount);
    }
    for (const status of [ROW_STATUS.added, ROW_STATUS.removed] as const) {
      const sd = status === ROW_STATUS.added ? 'new' : 'old';
      let run: number[] = [];
      const flush = () => {
        if (!run.length) return;
        const t = sd === 'new' ? d.newTable : d.oldTable;
        const o = t.origin ?? { row: 1, col: 0 };
        const idx = run.map((r) => (sd === 'new' ? d.newRows[r] : d.oldRows[r]));
        const lab = labelers[si][sd];
        const impact = run.reduce((n, r) => n + (rowImpact.get(r) ?? 0), 0);
        structure.push({
          sheet: s.name,
          kind: status === ROW_STATUS.added ? 'rows-added' : 'rows-removed',
          rows: idx.map((r) => r + o.row + 1),
          labels: idx.map((r) => {
            const l = lab(r, Math.max(0, t.columns.length - 1));
            return [...l.sections, ...l.row].join(' › ');
          }),
          ...(impact ? { impact } : {}),
        });
        run = [];
      };
      for (let r = 0; r < d.rowCount; r++) {
        if (d.status[r] === status) {
          if (run.length && run[run.length - 1] !== r - 1) flush();
          run.push(r);
        }
      }
      flush();
    }
    const moved: number[] = [];
    for (let r = 0; r < d.rowCount; r++) if (d.moved[r]) moved.push(r);
    if (moved.length) {
      const o = d.newTable.origin ?? { row: 1, col: 0 };
      structure.push({
        sheet: s.name,
        kind: 'rows-moved',
        rows: moved.map((r) => d.newRows[r] + o.row + 1),
        labels: moved.map((r) => {
          const l = labelers[si].new(d.newRows[r], 0);
          return [...l.sections, ...l.row].join(' › ');
        }),
      });
    }
    d.columns.forEach((c, ci) => {
      if (c.status === 'added' || c.status === 'removed') {
        const impact = colImpact.get(ci);
        structure.push({ sheet: s.name, kind: c.status === 'added' ? 'column-added' : 'column-removed', column: c.name, ...(impact ? { impact } : {}) });
      } else if (c.status === 'renamed') structure.push({ sheet: s.name, kind: 'column-renamed', column: `${c.oldName} → ${c.name}` });
    });
  });

  const st = trace.stats;
  const notes: string[] = [];
  if (st.stale) notes.push(`${st.stale} result${st.stale === 1 ? ' was' : 's were'} saved in the new file without recalculating; the recalculated values are shown.`);
  if (st.external) notes.push(`${st.external} value${st.external === 1 ? '' : 's'} came from outside data (a data add-in, a web or RTD function, or another workbook). ${st.external === 1 ? 'It is' : 'They are'} taken as saved, not recalculated.`);
  const byRefs = downstream.filter((c) => c.method === 'references').length;
  if (byRefs) notes.push(`${byRefs} downstream result${byRefs === 1 ? "'s causes come" : "s' causes come"} from formula references, not recalculation (a formula on the path couldn't be evaluated). They may list a cause that didn't matter, and have no contributions.`);
  if (st.unexplained) notes.push(`${st.unexplained} result${st.unexplained === 1 ? '' : 's'} changed with no edit found upstream (volatile functions such as NOW, or files calculated by different apps).`);
  if (trace.truncated) notes.push('The workbook was too large to trace completely; some causes may be missing.');
  const computed = [diff.computedValues?.old, diff.computedValues?.new].reduce((n, c) => n + (c?.computed ?? 0), 0);
  if (computed) notes.push(`${computed} formula result${computed === 1 ? ' was' : 's were'} missing from the files (saved by a tool that doesn't calculate) and computed by spreadiff.`);

  const report = {
    old: oldName,
    new: newName,
    totals: {
      inputs: edits.filter((e) => e.kind === 'input').length,
      formulas: edits.filter((e) => e.kind === 'formula').length,
      external: edits.filter((e) => e.kind === 'external').length,
      downstream: downstream.length,
      outputs: downstream.filter((c) => c.output).length,
      recalculated: st.recalculated,
      unexplained: st.unexplained,
      stale: st.stale,
    },
    edits,
    downstream,
    structure,
    notes,
  } as ChangeReport;
  // Not enumerable: JSON.stringify(report) stays the report.
  Object.defineProperties(report, { diff: { value: diff }, trace: { value: trace } });
  return report;
}
