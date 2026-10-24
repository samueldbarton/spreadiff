/**
 * What a person reading the sheet would call a cell: the section headings
 * above its row, the row's own label, and the column headings above it.
 *
 * ```text
 *        A            B        C
 *  1                  FY2024   FY2025      ← column headings
 *  2  India
 *  3    Revenue       100      120         ← C3 is "India › Revenue · FY2025"
 *  4    Costs         60       70
 *  5  China
 *  6    Revenue       80       95          ← C6 is "China › Revenue · FY2025"
 * ```
 *
 * Inferred from the layout, without styles where there are none:
 * - **Label columns** are the columns left of the first mostly-numeric
 *   one. A row's label is its first text cells there.
 * - **Section headings** are rows with a label and nothing to their right.
 *   Indentation and bold (from `Table.labelStyles`, or leading spaces) set
 *   their depth. Without either, a heading directly above another
 *   heading is its parent, and a `Total …` row closes the innermost section.
 * - **Column headings** are the nearest rows above the cell that hold
 *   only text, dates or years to the right of the label columns (and the
 *   header row, when there is one). In a stack of them, an upper heading
 *   spans to its right until the next one, as a merged or centred title does.
 */
import { columnLabel } from './csv.js';
import type { CellValue, Table } from './types.js';

/** A cell's labels, from {@link tableLabeler}. Each part may be empty. */
export interface CellLabel {
  /** Section headings above the row, outermost first: `['India']`. */
  sections: string[];
  /** The row's own label cells, left to right: `['Revenue']`, or `['P-1003', 'Smoked Mug']`. */
  row: string[];
  /** Column headings above the cell, outermost first: `['FY2025']`, `['Forecast', 'Q1']`, or the column's name. */
  column: string[];
}

/** Options for {@link tableLabeler}. */
export interface LabelOptions {
  /** Key columns, by name (`TableDiff.keyColumns`): their values lead each row's label. */
  keyColumns?: string[];
  /** Max label cells per row. Default 2. */
  maxRowCells?: number;
}

/** Rows scanned to find label columns. */
const SAMPLE = 5000;
const TOTAL = /^(?:total|sub-?total|grand total)\b|\btotal$/i;

const isBlank = (v: CellValue | undefined) => v === null || v === undefined || v === '';
const text = (v: CellValue | undefined): string => (v === null || v === undefined ? '' : typeof v === 'string' ? v.replace(/\s+/g, ' ').trim() : String(v));
const ERROR = /^#(?:VALUE!|N\/A|REF!|DIV\/0!|NUM!|NAME\?|NULL!|SPILL!|CALC!|GETTING_DATA)$/;
/** Column headings kept per cell: the nearest rows of a heading block (a title block above them adds noise). */
const MAX_HEADING_ROWS = 3;
const yearLike = (v: CellValue) => typeof v === 'number' && Number.isInteger(v) && v >= 1900 && v <= 2100;
const ISO_DATE = /^\d{4}-\d{2}-\d{2}(?:[ T]|$)/;

interface RowInfo {
  /** Label cells (column, text) in the label columns. */
  labels: [number, string][];
  /** Non-empty cells right of the label columns. */
  data: number;
  /** Only text, dates and years right of the label columns (a column-heading row). */
  band: boolean;
  /** Depth key of the first label: column × 64 + indent. */
  depth: number;
  bold: boolean;
}

/**
 * Labels for the cells of one table. Returns a function of table
 * coordinates (`row` into `table.rows`, `col` into `table.columns`); the
 * layout is analysed once, on the first call.
 *
 * ```ts
 * import { formatLabel, parseCsv, tableLabeler } from 'spreadiff/core';
 * const t = parseCsv(',FY2024,FY2025\nIndia,,\nRevenue,100,120\nChina,,\nRevenue,80,95', { header: false });
 * formatLabel(tableLabeler(t)(4, 2)); // 'China › Revenue · FY2025'
 * ```
 */
export function tableLabeler(table: Table, options: LabelOptions = {}): (row: number, col: number) => CellLabel {
  let analysed: ReturnType<typeof analyse> | null = null;
  return (row, col) => {
    analysed ??= analyse(table, options);
    return analysed.label(row, col);
  };
}

function analyse(table: Table, options: LabelOptions) {
  const { rows, columns } = table;
  const width = columns.length;
  const o = table.origin ?? { row: 1, col: 0 };
  // Column names count as headings when they came from a header row (not letters).
  const lettered = columns.every((c, i) => c === columnLabel(i + o.col));
  const header: CellValue[] | null = table.header?.length ? table.header : !lettered && width ? columns : null;
  const maxCells = options.maxRowCells ?? 2;
  const keyIdx = (options.keyColumns ?? []).map((k) => columns.indexOf(k)).filter((i) => i >= 0);

  // Label columns: left of the first column that is mostly numbers, in rows
  // that hold any (heading rows of text alone don't count).
  const n = Math.min(rows.length, SAMPLE);
  const numberRows: number[] = [];
  for (let r = 0; r < n; r++) if (rows[r].some((v) => typeof v === 'number')) numberRows.push(r);
  let zone = 0;
  for (; zone < width; zone++) {
    let filled = 0;
    let numeric = 0;
    for (const r of numberRows) {
      const v = rows[r][zone];
      if (isBlank(v)) continue;
      filled++;
      if (typeof v === 'number' || (table.formulas?.[r]?.[zone] && typeof v !== 'string')) numeric++;
    }
    if (filled && numeric * 2 >= filled) break;
  }
  if (zone === width) zone = Math.min(1, width); // no numeric column: a list; its first column labels the rows

  // Indent and bold of text cells, from the file's styles, in table coordinates.
  const styles = new Map<number, { indent?: number; bold?: boolean }>();
  for (const s of table.labelStyles ?? []) styles.set((s.r - o.row) * width + (s.c - o.col), s);
  const merges = (table.merges ?? []).filter((m) => m.r0 === m.r1).map((m) => ({ r: m.r0 - o.row, c0: m.c0 - o.col, c1: m.c1 - o.col }));

  const info: RowInfo[] = new Array(rows.length);
  const rowInfo = (r: number): RowInfo => {
    const cells = rows[r];
    const labels: [number, string][] = [];
    for (let c = 0; c < zone; c++) {
      const v = cells[c];
      if (typeof v === 'string' && v.trim() && !ERROR.test(v)) labels.push([c, text(v)]);
      else if (typeof v === 'boolean') labels.push([c, text(v)]);
    }
    // A heading row holds text, or dates or years running on (2024, 2025, or
    // 2024, 2024, … for quarters). One date or year alone is data, unless
    // the row's label introduces it ("Period end").
    let data = 0;
    let band = true;
    let words = 0;
    let years = 0;
    let lastYear = NaN;
    let dates = 0;
    let lastDate = '';
    for (let c = zone; c < width; c++) {
      const v = cells[c];
      if (isBlank(v)) continue;
      data++;
      if (typeof v === 'string' && ISO_DATE.test(v)) {
        if (v <= lastDate) band = false;
        lastDate = v;
        dates++;
      } else if (typeof v === 'string' && ERROR.test(v)) band = false;
      else if (typeof v === 'string' || typeof v === 'boolean') words++;
      else if (yearLike(v) && (Number.isNaN(lastYear) || v - lastYear === 0 || v - lastYear === 1)) {
        lastYear = v;
        years++;
      } else band = false;
    }
    if (!words && years + dates === 1 && (labels.length ? years === 1 : dates === 1)) band = false;
    let depth = 0;
    let bold = false;
    if (labels.length) {
      const c = labels[0][0];
      const raw = cells[c];
      const lead = typeof raw === 'string' ? raw.length - raw.trimStart().length : 0;
      const st = styles.get(r * width + c);
      depth = c * 64 + Math.min(63, (st?.indent ?? 0) + Math.floor(lead / 2));
      bold = !!st?.bold;
    }
    return { labels, data, band: data > 0 && band, depth, bold };
  };
  for (let r = 0; r < rows.length; r++) info[r] = rowInfo(r);

  // Section headings, top to bottom: a stack of open sections. Tables
  // matched by key are records, not a model: no sections there.
  const sections: string[][] = new Array(rows.length);
  const noSections = keyIdx.length > 0 || zone === 0;
  const isHeading = (r: number) => !!info[r].labels.length && info[r].data === 0 && !info[r].band;
  // A heading whose next row is another heading opens a group ("India" over "Revenue").
  const group = new Uint8Array(rows.length);
  if (!noSections) {
    let next = -1; // the next row with content below r
    for (let r = rows.length - 1; r >= 0; r--) {
      if (isHeading(r)) group[r] = next >= 0 && isHeading(next) ? 1 : 0;
      if (info[r].labels.length || info[r].data) next = r;
    }
  }
  // With indentation (or labels in several columns) depth alone orders headings.
  let depths = 0;
  const seen = new Set<number>();
  for (let r = 0; r < rows.length; r++) if (info[r].labels.length && !seen.has(info[r].depth)) seen.add(info[r].depth), depths++;
  const byDepth = depths > 1;
  type Open = { label: string; depth: number; group: boolean };
  const stack: Open[] = [];
  const snapshot = () => stack.map((x) => x.label);
  let prev: 'heading' | 'data' | null = null;
  for (let r = 0; r < rows.length; r++) {
    const ri = info[r];
    const label = ri.labels[0]?.[1];
    sections[r] = snapshot();
    if (noSections || !label || ri.band) continue; // blank and column-heading rows leave sections as they are
    while (stack.length && stack[stack.length - 1].depth > ri.depth) stack.pop();
    if (isHeading(r)) {
      // Right under another heading, it nests. Otherwise it ends its sibling
      // at the same depth: a group ends the last group too.
      const isGroup = !!group[r] || ri.bold;
      if (byDepth) {
        while (stack.length && stack[stack.length - 1].depth >= ri.depth) stack.pop();
      } else if (prev !== 'heading') {
        const atDepth = (x: Open | undefined) => !!x && x.depth === ri.depth;
        while (atDepth(stack[stack.length - 1]) && !stack[stack.length - 1].group) stack.pop();
        if (isGroup && atDepth(stack[stack.length - 1])) stack.pop();
      }
      sections[r] = snapshot();
      stack.push({ label, depth: ri.depth, group: isGroup });
      prev = 'heading';
      continue;
    }
    sections[r] = snapshot();
    if (TOTAL.test(label)) {
      // "Total India" closes India (and anything inside it); a bare "Total revenue" closes the innermost.
      const rest = label.replace(TOTAL, '').trim().toLowerCase();
      const named = rest ? stack.map((x) => x.label.toLowerCase()).lastIndexOf(rest) : -1;
      stack.length = named >= 0 ? named : Math.max(0, stack.length - 1);
    }
    prev = 'data';
  }

  // Column headings: the nearest block of heading rows above each row.
  const bandRows: number[][] = new Array(rows.length); // shared arrays per block
  let current: number[] = header ? [-1] : [];
  let open: number[] | null = null;
  for (let r = 0; r < rows.length; r++) {
    bandRows[r] = current;
    if (info[r].band) {
      if (!open) open = [];
      open.push(r);
    } else if (info[r].data > 0 || info[r].labels.length) {
      if (open) current = open;
      open = null;
      bandRows[r] = current;
    }
  }
  const valueAt = (r: number, c: number): CellValue => (r < 0 ? header![c] ?? null : rows[r][c]);
  const heading = (r: number, c: number, spanRight: boolean): string => {
    const own = valueAt(r, c);
    if (!isBlank(own)) return text(own);
    const m = merges.find((g) => g.r === r && c >= g.c0 && c <= g.c1);
    if (m) return text(valueAt(r, m.c0));
    if (!spanRight || c < zone) return '';
    for (let x = c - 1; x >= zone; x--) if (!isBlank(valueAt(r, x))) return text(valueAt(r, x));
    return '';
  };

  return {
    label(row: number, col: number): CellLabel {
      const ri = info[row];
      if (!ri) return { sections: [], row: [], column: [] };
      const own: string[] = [];
      for (const k of keyIdx) if (!isBlank(rows[row][k])) own.push(text(rows[row][k]));
      for (const [c, t] of ri.labels) {
        if (own.length >= maxCells) break;
        if (c === col && ri.data > 0) continue; // the cell itself isn't its own label
        if (!own.includes(t)) own.push(t);
      }
      const block = bandRows[row].slice(-MAX_HEADING_ROWS);
      const column: string[] = [];
      block.forEach((br, i) => {
        const h = heading(br, col, i < block.length - 1);
        if (h && column[column.length - 1] !== h) column.push(h);
      });
      if (!column.length && header && !block.includes(-1)) {
        const h = text(header[col]);
        if (h) column.push(h);
      }
      return { sections: sections[row], row: own, column };
    },
  };
}

/**
 * One line for a {@link CellLabel}: sections and row joined by ` › `, then
 * ` · ` and the column headings (`'India › Revenue · FY2025'`). Empty when
 * there are no labels.
 */
export function formatLabel(label: CellLabel): string {
  const rowPart = [...label.sections, ...label.row].join(' › ');
  const colPart = label.column.join(' › ');
  return rowPart && colPart ? `${rowPart} · ${colPart}` : rowPart || colPart;
}
