/**
 * A single cell value. Parsers produce numbers only when the conversion is
 * lossless (`"007"` stays a string), and dates as ISO strings
 * (`"2026-01-31"`, `"2026-01-31 14:05:00"`, or `"14:05:00"` for times).
 */
export type CellValue = string | number | boolean | null;

/** One sheet of tabular data. */
export interface Table {
  /** Sheet name. Used to pair sheets between workbooks and to resolve `Sheet!A1` references. */
  name: string;
  /** Column names. Parsers make them unique; when the source has no header row they are letters: `A`, `B`, … */
  columns: string[];
  /** Data rows, without the header row. Each row has one value per column. */
  rows: CellValue[][];
  /**
   * Formulas, parallel to `rows`: `formulas[r][c]` is the formula text
   * (`"=B2*C2"`) or `null`. Rows may be missing (sparse). Set this, even
   * empty, for any source that can hold formulas, so a formula replaced by
   * a constant counts as a change. CSV tables leave it undefined.
   */
  formulas?: (string | null)[][];
  /**
   * Where `rows[0][0]` sits in the source sheet, 0-based. Used to show real
   * spreadsheet row numbers and column letters, and to resolve formula
   * references. Defaults to `{ row: 1, col: 0 }`: data starts at A2, under
   * a header row.
   */
  origin?: { row: number; col: number };
}

/** An ordered list of sheets. A CSV file is a one-sheet workbook. */
export interface Workbook {
  /** Usually the file name. */
  name: string;
  /** Sheets in workbook order. */
  sheets: Table[];
  /** Workbook-scoped defined names (`TaxRate` → `Inputs!$B$2`), used to resolve formula references. */
  names?: Record<string, string>;
}

/** How an aligned row changed. */
export type RowStatus = 'equal' | 'added' | 'removed' | 'modified';

/** How an aligned column changed. `renamed` columns are matched by content, not name. */
export type ColumnStatus = 'equal' | 'added' | 'removed' | 'renamed';

/** One column of the diff, matched across the old and new tables. */
export interface AlignedColumn {
  /** Index in the old table's columns, or -1 when the column was added. */
  oldIndex: number;
  /** Index in the new table's columns, or -1 when the column was removed. */
  newIndex: number;
  /** Column name: the new name, or the old one for removed columns. */
  name: string;
  /** Previous name, set when `status === 'renamed'`. */
  oldName?: string;
  /** How the column changed. */
  status: ColumnStatus;
  /** Number of changed cells in this column. */
  changes: number;
}

/** Options for {@link diffTables} and {@link diffWorkbooks}. All are optional. */
export interface DiffOptions {
  /**
   * How to match rows between the two tables.
   * - `string[]`: match by these key columns (must exist in both tables).
   *   Finds edits regardless of row order; duplicate keys match in order.
   * - `'auto'` (default): use the leftmost id-like column (`id`, `ID`,
   *   `Order ID`, `order_id`, `customerId`, `sku`, `key`, `code`, …) whose
   *   values are unique and non-empty on both sides and that matches at
   *   least half the rows; otherwise match by row order.
   * - `'none'`: always match by row order (sequence alignment).
   */
  keys?: string[] | 'auto' | 'none';
  /** Treat numbers (or numeric strings) within this absolute difference as equal. Default 0. */
  numericTolerance?: number;
  /** Ignore leading and trailing whitespace in text. Default true. */
  trimWhitespace?: boolean;
  /** Compare text case-insensitively. Default false. */
  ignoreCase?: boolean;
  /**
   * Count a cell as changed when its formula's meaning changed, even if the
   * cached value didn't. References are compared by the aligned cells they
   * point at, so formulas shifted by inserted rows aren't changes. Applies
   * only when both tables carry formulas (XLSX vs. XLSX). Default true.
   */
  compareFormulas?: boolean;
  /**
   * When matching by row order, a removed and an added row in the same gap
   * pair up as one "modified" row if at least this share of their
   * non-formula cells are equal. Default 0.5. A single removed and a
   * single added row at the same place pair if any cell matches.
   */
  rowSimilarity?: number;
  /**
   * Pair a removed and an added column as a rename when at least 80% of
   * their non-blank values match across matched rows. Default true.
   */
  detectRenames?: boolean;
}

/** Counts of what changed in a {@link TableDiff}. */
export interface DiffStats {
  /** Rows only in the new table. */
  rowsAdded: number;
  /** Rows only in the old table. */
  rowsRemoved: number;
  /** Rows in both tables with at least one changed cell. */
  rowsModified: number;
  /** Rows in both tables with no changed cells. */
  rowsEqual: number;
  /** Changed cells across all modified rows. Cells of added/removed rows and columns aren't counted. */
  cellsChanged: number;
  /** Columns only in the new table. */
  columnsAdded: number;
  /** Columns only in the old table. */
  columnsRemoved: number;
  /** Columns matched by content under a new name. */
  columnsRenamed: number;
}

/**
 * The result of diffing two tables: a sequence of **aligned rows**, each
 * pointing at an old row, a new row, or both, over **aligned columns**.
 *
 * Per-row data lives in typed arrays, so a million-row diff costs a few
 * megabytes. To read aligned row `r`:
 *
 * ```ts
 * import type { TableDiff } from 'spreadiff';
 * declare const diff: TableDiff;
 * const r = 0;
 * const oldRow = diff.oldRows[r] >= 0 ? diff.oldTable.rows[diff.oldRows[r]] : null;
 * const newRow = diff.newRows[r] >= 0 ? diff.newTable.rows[diff.newRows[r]] : null;
 * const changedInColumn2 = diff.changedCells.get(r)?.[2] === 1;
 * ```
 */
export interface TableDiff {
  /** The table diffed from. */
  oldTable: Table;
  /** The table diffed to. */
  newTable: Table;
  /** Aligned columns, in display order (new order, removed columns where they were). */
  columns: AlignedColumn[];
  /** Number of aligned rows. */
  rowCount: number;
  /** Old row index per aligned row, -1 when the row was added. */
  oldRows: Int32Array;
  /** New row index per aligned row, -1 when the row was removed. */
  newRows: Int32Array;
  /** Status per aligned row, encoded as {@link ROW_STATUS} numbers. */
  status: Uint8Array;
  /**
   * Changed cells of modified rows: aligned row → mask with one entry per
   * aligned column, 1 where the cell changed. Rows without changes are absent.
   */
  changedCells: Map<number, Uint8Array>;
  /**
   * Cells whose formula changed in meaning: a subset of `changedCells`, in
   * the same shape. Formulas that only shifted because rows were inserted
   * above, or ranges that grew only by inserted rows, don't count.
   */
  formulaChanged: Map<number, Uint8Array>;
  /** Key columns used to match rows; empty when rows were matched by order. */
  keyColumns: string[];
  /** Counts of what changed. */
  stats: DiffStats;
  /** The options used, with defaults filled in. */
  options: Required<DiffOptions>;
}

/** Numeric encoding of {@link RowStatus} used in `TableDiff.status`. */
export const ROW_STATUS = { equal: 0, added: 1, removed: 2, modified: 3 } as const;

/** Decode `TableDiff.status` values: `ROW_STATUS_NAMES[diff.status[r]]`. */
export const ROW_STATUS_NAMES: RowStatus[] = ['equal', 'added', 'removed', 'modified'];

/** One sheet of a {@link WorkbookDiff}. */
export interface SheetDiff {
  /** Sheet name (the new name for sheets present in both). */
  name: string;
  /** `added` / `removed` for sheets in only one workbook; otherwise whether anything changed. */
  status: 'equal' | 'modified' | 'added' | 'removed';
  /** The sheet's diff. Added and removed sheets diff against an empty table. */
  diff: TableDiff;
}

/** The result of {@link diffWorkbooks}: one {@link SheetDiff} per sheet. */
export interface WorkbookDiff {
  /** Sheets in new-workbook order, with removed sheets where they were. */
  sheets: SheetDiff[];
  /** Defined names of the old and new workbooks, for formula tracing. */
  names?: { old?: Record<string, string>; new?: Record<string, string> };
}
