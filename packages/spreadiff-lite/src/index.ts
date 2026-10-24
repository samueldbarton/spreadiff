/**
 * spreadiff-lite: what changed between two spreadsheets, and what it moved
 * downstream, as Markdown or JSON. No UI.
 *
 * - {@link compare} two files (XLSX, CSV, TSV) or parsed workbooks into a
 *   {@link ChangeReport}: each edit with its before and after values and
 *   what the sheet calls the cell, each downstream result with the edits
 *   that moved it and by how much, and shape changes.
 * - {@link toMarkdown} renders a report for a person, a pull request or an agent.
 * - {@link summarize} does both.
 */
import { compare, type CompareOptions, type FileInput } from './report.js';
import { toMarkdown, type MarkdownOptions } from './markdown.js';

export { compare, reportOf, type ChangeReport, type CellChange, type Cause, type CompareOptions, type FileInput, type StructureChange } from './report.js';
export { toMarkdown, type MarkdownOptions } from './markdown.js';
export { externalSource, LIVE_FUNCTIONS, tableLabeler, formatLabel, type CellLabel, type ExternalSource } from 'spreadiff/core';

/**
 * Compare two files and return the Markdown summary in one call.
 *
 * ```ts
 * import { readFileSync } from 'node:fs';
 * import { summarize } from 'spreadiff-lite';
 *
 * const md = summarize(
 *   { name: 'before.xlsx', data: readFileSync('before.xlsx') },
 *   { name: 'after.xlsx', data: readFileSync('after.xlsx') },
 * );
 * ```
 */
export function summarize(oldFile: FileInput, newFile: FileInput, options: CompareOptions & MarkdownOptions = {}): string {
  return toMarkdown(compare(oldFile, newFile, options), options);
}
