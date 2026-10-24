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
const TOTAL = /^\s*(?:total|sub-?total|grand total)\b|\btotal\s*$/i;
/** Whether a string has a non-space character, without allocating. */
const hasText = (v: string) => {
  for (let i = 0; i < v.length; i++) if (v.charCodeAt(i) > 32) return true;
  return false;
};

const isBlank = (v: CellValue | undefined) => v === null || v === undefined || v === '';
const text = (v: CellValue | undefined): string => (v === null || v === undefined ? '' : typeof v === 'string' ? v.replace(/\s+/g, ' ').trim() : String(v));
const ERROR = /^#(?:VALUE!|N\/A|REF!|DIV\/0!|NUM!|NAME\?|NULL!|SPILL!|CALC!|GETTING_DATA)$/;
/** Column headings kept per cell: the nearest rows of a heading block (a title block above them adds noise). */
const MAX_HEADING_ROWS = 3;
const yearLike = (v: CellValue) => typeof v === 'number' && Number.isInteger(v) && v >= 1900 && v <= 2100;
const ISO_DATE = /^\d{4}-\d{2}-\d{2}(?:[ T]|$)/;

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
  let analysed: SheetLabels | null = null;
  return (row, col) => {
    analysed ??= sheetLabels(table, options);
    return analysed.cell(row, col);
  };
}

/** A row's labels, from {@link SheetLabels}. */
export interface RowLabel {
  /** Section headings above the row, outermost first. */
  sections: string[];
  /** The row's own label cells. */
  row: string[];
  /** The row is itself a section heading (a label with nothing to its right). */
  heading: boolean;
}

/** The labels of a whole table, from {@link sheetLabels}. Table coordinates throughout. */
export interface SheetLabels {
  /** A cell's labels: what {@link tableLabeler} returns. */
  cell(row: number, col: number): CellLabel;
  /** A row's labels, and whether it is a section heading. */
  row(row: number): RowLabel;
  /** The column headings in force at a row (the nearest heading rows above it, or the header). */
  column(row: number, col: number): string[];
  /** The row holds column headings (a row of years, dates or titles right of the labels), and no label of its own. */
  headingRow(row: number): boolean;
  /**
   * Columns whose every non-blank cell is part of its row's label: shown
   * by the row labels, so a view that shows those can hide the columns.
   * Computed once, on the first call.
   */
  labelColumns(): number[];
}

/**
 * Every label of a table at once: rows, columns and cells. The layout is
 * analysed once, when this is called; lookups are then cheap. `DiffGrid`
 * uses this for its row and column headers.
 */
export function sheetLabels(table: Table, options: LabelOptions = {}): SheetLabels {
  return analyse(table, options);
}

function analyse(table: Table, options: LabelOptions): SheetLabels {
  const { rows, columns } = table;
  const width = columns.length;
  const n = rows.length;
  const o = table.origin ?? { row: 1, col: 0 };
  // Column names count as headings when they came from a header row (not letters).
  const lettered = columns.every((c, i) => c === columnLabel(i + o.col));
  const header: CellValue[] | null = table.header?.length ? table.header : !lettered && width ? columns : null;
  const hasHeaderRow = !!header;
  const maxCells = options.maxRowCells ?? 2;
  const keyIdx = (options.keyColumns ?? []).map((k) => columns.indexOf(k)).filter((i) => i >= 0);

  // Label columns: left of the first column that is mostly numbers, in rows
  // that hold any (heading rows of text alone don't count).
  const sample = Math.min(n, SAMPLE);
  const numberRows: number[] = [];
  for (let r = 0; r < sample; r++) if (rows[r].some((v) => typeof v === 'number')) numberRows.push(r);
  let zone = 0;
  for (; zone < width; zone++) {
    if (keyIdx.includes(zone)) continue; // a numeric id doesn't end the labels (id, name, qty)
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

  // One pass over the rows, into typed arrays (a million rows allocate nothing per row).
  const isLabel = (v: CellValue | undefined) => (typeof v === 'string' ? hasText(v) && !(v.charCodeAt(0) === 35 && ERROR.test(v)) : typeof v === 'boolean');
  const firstLabel = new Int16Array(n).fill(-1); // column of the row's first label cell
  const dataCount = new Uint32Array(n); // non-empty cells right of the label columns
  const band = new Uint8Array(n); // a column-heading row
  const depth = new Uint16Array(n); // column × 64 + indent of the first label
  const bold = new Uint8Array(n);
  for (let r = 0; r < n; r++) {
    const cells = rows[r];
    let first = -1;
    for (let c = 0; c < zone; c++) {
      if (isLabel(cells[c])) {
        first = c;
        break;
      }
    }
    // A heading row holds text, or dates or years running on (2024, 2025, or
    // 2024, 2024, … for quarters). One date or year alone is data, unless
    // the row's label introduces it ("Period end").
    let data = 0;
    let isBand = true;
    let words = 0;
    let years = 0;
    let lastYear = NaN;
    let dates = 0;
    let lastDate = '';
    for (let c = zone; c < width; c++) {
      const v = cells[c];
      if (isBlank(v)) continue;
      data++;
      if (!isBand) continue; // only the count matters now
      if (typeof v === 'string') {
        if (v.length >= 10 && ISO_DATE.test(v)) {
          if (v <= lastDate) isBand = false;
          lastDate = v;
          dates++;
        } else if (v.charCodeAt(0) === 35 && ERROR.test(v)) isBand = false;
        else words++;
      } else if (typeof v === 'boolean') words++;
      else if (yearLike(v) && (Number.isNaN(lastYear) || v - lastYear === 0 || v - lastYear === 1)) {
        lastYear = v;
        years++;
      } else isBand = false;
    }
    // A run of years or dates, or text alone; one year or date beside text is data (an id of 2001, a due date).
    const runs = years + dates;
    if (runs === 1 && (words > 0 || first >= 0 || dates === 1)) isBand = false;
    // A table with a header row has its column headings already.
    if (hasHeaderRow) isBand = false;
    firstLabel[r] = first;
    dataCount[r] = data;
    band[r] = data > 0 && isBand ? 1 : 0;
    if (first >= 0) {
      const raw = cells[first];
      const lead = typeof raw === 'string' ? raw.length - raw.trimStart().length : 0;
      const st = styles.size ? styles.get(r * width + first) : undefined;
      depth[r] = first * 64 + Math.min(63, (st?.indent ?? 0) + Math.floor(lead / 2));
      bold[r] = st?.bold ? 1 : 0;
    }
  }
  const labelOf = (r: number) => text(rows[r][firstLabel[r]]);
  const hasLabel = (r: number) => firstLabel[r] >= 0;
  const isHeading = (r: number) => hasLabel(r) && dataCount[r] === 0 && !band[r];

  // Section headings, top to bottom: a stack of open sections, kept as a
  // tree of nodes; each row records the innermost open one. Tables matched
  // by key are records, not a model: no sections there.
  const nodeLabel: string[] = [];
  const nodeParent: number[] = [];
  const sectionAt = new Int32Array(n).fill(-1);
  const noSections = keyIdx.length > 0 || zone === 0;
  if (!noSections) {
    // A heading whose next row is another heading opens a group ("India" over "Revenue").
    const group = new Uint8Array(n);
    let next = -1; // the next row with content below r
    for (let r = n - 1; r >= 0; r--) {
      if (isHeading(r)) group[r] = next >= 0 && isHeading(next) ? 1 : 0;
      if (hasLabel(r) || dataCount[r]) next = r;
    }
    // With indentation (or labels in several columns) depth alone orders headings.
    let firstDepth = -1;
    let byDepth = false;
    for (let r = 0; r < n && !byDepth; r++) {
      if (!hasLabel(r)) continue;
      if (firstDepth < 0) firstDepth = depth[r];
      else if (depth[r] !== firstDepth) byDepth = true;
    }
    type Open = { node: number; depth: number; group: boolean; label: string };
    const stack: Open[] = [];
    const top = () => stack[stack.length - 1];
    const innermost = () => (stack.length ? top().node : -1);
    let prev: 'heading' | 'data' | null = null;
    for (let r = 0; r < n; r++) {
      sectionAt[r] = innermost();
      if (!hasLabel(r) || band[r]) continue; // blank and column-heading rows leave sections as they are
      const d = depth[r];
      while (stack.length && top().depth > d) stack.pop();
      if (isHeading(r)) {
        const label = labelOf(r);
        // Right under another heading, it nests. Otherwise it ends its sibling
        // at the same depth: a group ends the last group too.
        const isGroup = !!group[r] || !!bold[r];
        if (byDepth) {
          while (stack.length && top().depth >= d) stack.pop();
        } else if (prev !== 'heading') {
          while (stack.length && top().depth === d && !top().group) stack.pop();
          if (isGroup && stack.length && top().depth === d) stack.pop();
        }
        sectionAt[r] = innermost();
        nodeLabel.push(label);
        nodeParent.push(innermost());
        stack.push({ node: nodeLabel.length - 1, depth: d, group: isGroup, label });
        prev = 'heading';
        continue;
      }
      // With indentation, a row at a heading's own depth isn't inside it ("Total profit" under "  Costs").
      if (byDepth) while (stack.length && top().depth >= d) stack.pop();
      sectionAt[r] = innermost();
      const raw = rows[r][firstLabel[r]];
      if (typeof raw === 'string' && TOTAL.test(raw)) {
        // "Total India" closes India (and anything inside it); a bare "Total revenue" closes the innermost.
        const rest = text(raw).replace(TOTAL, '').trim().toLowerCase();
        let named = -1;
        if (rest) for (let i = stack.length - 1; i >= 0 && named < 0; i--) if (stack[i].label.toLowerCase() === rest) named = i;
        stack.length = named >= 0 ? named : Math.max(0, stack.length - 1);
      }
      prev = 'data';
    }
  }
  const sectionsOf = (r: number): string[] => {
    const out: string[] = [];
    for (let k = sectionAt[r] ?? -1; k >= 0; k = nodeParent[k]) out.push(nodeLabel[k]);
    return out.reverse();
  };

  // Column headings: the nearest block of heading rows above each row, as a block id.
  const blocks: number[][] = [header ? [-1] : []];
  const blockAt = new Int32Array(n);
  let current = 0;
  let open: number[] | null = null;
  for (let r = 0; r < n; r++) {
    blockAt[r] = current;
    if (band[r]) {
      if (!open) open = [];
      open.push(r);
    } else if (dataCount[r] > 0 || hasLabel(r)) {
      if (open) {
        blocks.push(open);
        current = blocks.length - 1;
      }
      open = null;
      blockAt[r] = current;
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

  const rowLabel = (row: number, col = -1) => {
    const own: string[] = [];
    if (row < 0 || row >= n) return own;
    const cells = rows[row];
    // Text keys lead the label; a numeric id would read like a row number.
    for (const k of keyIdx) if (typeof cells[k] === 'string' && hasText(cells[k] as string)) own.push(text(cells[k]));
    for (let c = Math.max(0, firstLabel[row]); c < zone && own.length < maxCells && firstLabel[row] >= 0; c++) {
      if (!isLabel(cells[c])) continue;
      if (c === col && dataCount[row] > 0) continue; // the cell itself isn't its own label
      const t = text(cells[c]);
      if (!own.includes(t)) own.push(t);
    }
    return own;
  };
  const columnHeading = (row: number, col: number) => {
    const column: string[] = [];
    let b = row >= 0 && row < n ? blockAt[row] : 0;
    // Above the first heading rows (or on them), with no header: the headings just below apply.
    if (!blocks[b].length && b + 1 < blocks.length) b++;
    const block = blocks[b].slice(-MAX_HEADING_ROWS);
    block.forEach((br, i) => {
      const h = heading(br, col, i < block.length - 1);
      if (h && column[column.length - 1] !== h) column.push(h);
    });
    if (!column.length && header && !block.includes(-1)) {
      const h = text(header[col]);
      if (h) column.push(h);
    }
    return column;
  };
  let absorbed: number[] | null = null;
  return {
    cell: (row: number, col: number): CellLabel => (row >= 0 && row < n ? { sections: sectionsOf(row), row: rowLabel(row, col), column: columnHeading(row, col) } : { sections: [], row: [], column: [] }),
    row: (row: number): RowLabel => (row >= 0 && row < n ? { sections: sectionsOf(row), row: rowLabel(row), heading: isHeading(row) } : { sections: [], row: [], heading: false }),
    column: columnHeading,
    headingRow: (row: number) => row >= 0 && row < n && !!band[row] && !hasLabel(row),
    labelColumns(): number[] {
      if (absorbed) return absorbed;
      // A column every cell of which is blank or one of its row's label cells: the row labels show it.
      const textKeys = keyIdx.filter((k) => rows.every((r) => isBlank(r[k]) || typeof r[k] === 'string'));
      const candidates = new Set<number>(keyIdx.length <= maxCells ? textKeys : []);
      for (let c = 0; c < zone; c++) candidates.add(c);
      absorbed = [...candidates].filter((c) => {
        let used = false;
        const isKey = textKeys.includes(c);
        for (let r = 0; r < n; r++) {
          const v = rows[r][c];
          if (isBlank(v)) continue;
          used = true;
          if (isKey) continue;
          if (!isLabel(v)) return false;
          // Cheap path: the label holds every key and the first label cells up to maxCells.
          let before = 0;
          for (const k of keyIdx) if (typeof rows[r][k] === 'string') before++;
          for (let x = Math.max(0, firstLabel[r]); x < c; x++) if (!keyIdx.includes(x) && isLabel(rows[r][x])) before++;
          if (before < maxCells) continue;
          if (!rowLabel(r).includes(text(v))) return false;
        }
        return used;
      });
      return absorbed;
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
