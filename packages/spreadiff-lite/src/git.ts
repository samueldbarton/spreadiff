/**
 * Git integration: a textconv rendering for `git diff`, an external diff
 * driver that prints the full report, and a merge driver for tables.
 * See docs/git.md for the configuration.
 */
import { displayCell, mergeTables, parseFile, tableToCsv, type Table, type Workbook } from 'spreadiff/core';
import { compare, type CompareOptions } from './report.js';
import { toMarkdown, type MarkdownOptions } from './markdown.js';

/**
 * A workbook as stable, line-per-row text, for `git diff` through a
 * textconv driver: each sheet under a `== Sheet ==` line, each row as its
 * sheet row number and its cells as the sheet shows them, formulas after
 * their values in braces. Empty rows are left out.
 */
export function textconv(wb: Workbook): string {
  const out: string[] = [];
  for (const t of wb.sheets) {
    out.push(`== ${t.name} ==`);
    const o = t.origin ?? { row: 1, col: 0 };
    if (o.row > 0 && t.columns.length) out.push(`${o.row}: ${t.columns.join(' | ')}`);
    t.rows.forEach((row, r) => {
      if (row.every((v) => v === null || v === '')) return;
      const cells = row.map((_, c) => {
        const shown = displayCell(t, r, c).trim();
        const f = t.formulas?.[r]?.[c];
        return f ? `${shown} {${f}}` : shown;
      });
      out.push(`${r + o.row + 1}: ${cells.join(' | ').replace(/\s+\|\s*$/g, '').replace(/[\r\n]+/g, ' ')}`);
    });
  }
  return out.join('\n') + '\n';
}

/** An empty workbook, for files git reports as added or deleted (`/dev/null`). */
const empty = (name: string): Workbook => ({ name, sheets: [] });

/**
 * What `git diff` prints through `diff.<driver>.command`: git passes
 * `path old-file old-hex old-mode new-file new-hex new-mode`. Returns the
 * text to print: a heading and the report.
 */
export function gitDiff(
  args: string[],
  read: (path: string) => Uint8Array,
  options: CompareOptions & MarkdownOptions = {},
): string {
  const [path, oldFile, , , newFile] = args;
  if (!path || oldFile === undefined || newFile === undefined) throw new Error('--git needs the 7 arguments git passes to a diff driver');
  const load = (file: string) => (file === '/dev/null' ? empty(path) : parseFile(path, read(file), { header: 'auto', ...options.parse }));
  const report = compare(load(oldFile), load(newFile), options);
  report.old = `a/${path}`;
  report.new = `b/${path}`;
  return toMarkdown(report, { ...options, headingLevel: options.headingLevel ?? 2 });
}

/**
 * Three-way merge of two CSV/TSV files (`merge.<driver>.driver`): the
 * merged table, and the conflicts in words. Workbooks with several sheets
 * merge sheet by sheet; only single-table files can be written back as
 * CSV, so .xlsx files report their conflicts and aren't merged.
 */
export function mergeFiles(
  name: string,
  base: Uint8Array,
  ours: Uint8Array,
  theirs: Uint8Array,
  options: { prefer?: 'ours' | 'theirs'; keys?: string[] } = {},
): { csv: string | null; conflicts: string[]; merged: Table[] } {
  const read = (b: Uint8Array) => parseFile(name, b, { header: 'auto' });
  const [b, o, t] = [read(base), read(ours), read(theirs)];
  const conflicts: string[] = [];
  const merged: Table[] = [];
  for (const sheet of o.sheets) {
    const bs = b.sheets.find((x) => x.name === sheet.name) ?? { ...sheet, rows: [] };
    const ts = t.sheets.find((x) => x.name === sheet.name) ?? bs;
    const m = mergeTables(bs, sheet, ts, { prefer: options.prefer, ...(options.keys ? { keys: options.keys } : {}) });
    merged.push(m.table);
    for (const c of m.conflicts) {
      const where = c.row >= 0 ? `row ${c.row + 2}` : `base row ${(c.baseRow ?? 0) + 2}`;
      const prefix = o.sheets.length > 1 ? `${sheet.name} ` : '';
      if (c.kind === 'cell') conflicts.push(`${prefix}${where}, ${c.column}: ours ${JSON.stringify(c.ours)}, theirs ${JSON.stringify(c.theirs)} (base ${JSON.stringify(c.base)}); took ${c.took}`);
      else if (c.kind === 'delete-edit') conflicts.push(`${prefix}${where}: ${c.oursRow ? 'we edited it, they deleted it' : 'we deleted it, they edited it'}; took ${c.took}`);
      else conflicts.push(`${prefix}${where}: both added it, differently; took ${c.took}`);
    }
  }
  const single = !/\.xls[xmb]?$|\.ods$/i.test(name) && merged.length === 1;
  return { csv: single ? tableToCsv(merged[0], /\.tsv$/i.test(name) ? '\t' : ',') : null, conflicts, merged };
}
