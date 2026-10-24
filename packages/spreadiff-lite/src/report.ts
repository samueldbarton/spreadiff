import {
  diffWorkbooks,
  parseFile,
  reportOf,
  traceWorkbook,
  type ChangeReport,
  type DiffOptions,
  type ParseFileOptions,
  type TraceOptions,
  type Workbook,
} from 'spreadiff/core';

export { reportOf, type ChangeReport, type CellChange, type Cause, type StructureChange } from 'spreadiff/core';

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

