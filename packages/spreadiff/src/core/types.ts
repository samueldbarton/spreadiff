/** A single cell value. Dates are carried as ISO strings by the parsers. */
export type CellValue = string | number | boolean | null;

/** One sheet worth of tabular data. `rows` excludes the header row. */
export interface Table {
  name: string;
  /** Column names. When the source has no header row, use `A`, `B`, … */
  columns: string[];
  rows: CellValue[][];
  /**
   * Optional formulas, parallel to `rows` (`null` where a cell has none). A
   * row may be missing (sparse). Set this (even empty) for any source that
   * can hold formulas, so a formula replaced by a value counts as a change.
   */
  formulas?: (string | null)[][];
  /**
   * Where `rows[0][0]` sits in the source sheet (0-based), for showing real
   * spreadsheet row numbers and column letters. CSV with a header: `{ row: 1, col: 0 }`.
   */
  origin?: { row: number; col: number };
}

/** A workbook is an ordered list of sheets. CSV files are one-sheet workbooks. */
export interface Workbook {
  name: string;
  sheets: Table[];
}

export type RowStatus = 'equal' | 'added' | 'removed' | 'modified';
export type ColumnStatus = 'equal' | 'added' | 'removed' | 'renamed';

export interface AlignedColumn {
  /** Index in the old table's columns, or -1 when the column was added. */
  oldIndex: number;
  /** Index in the new table's columns, or -1 when the column was removed. */
  newIndex: number;
  name: string;
  /** Previous name, set when `status === 'renamed'`. */
  oldName?: string;
  status: ColumnStatus;
  /** Number of modified cells in this column. */
  changes: number;
}

export interface DiffOptions {
  /**
   * Columns that uniquely identify a row. When set, rows are matched by key
   * (detects edits regardless of reordering). `'auto'` picks a column whose
   * values are unique and non-empty in both tables, falling back to sequence
   * alignment. Defaults to `'auto'`.
   */
  keys?: string[] | 'auto' | 'none';
  /** Treat numbers within this absolute tolerance as equal. Default 0. */
  numericTolerance?: number;
  /** Ignore leading/trailing whitespace in text. Default true. */
  trimWhitespace?: boolean;
  /** Compare text case-insensitively. Default false. */
  ignoreCase?: boolean;
  /** Also count a cell as changed when only its formula changed. Default true. */
  compareFormulas?: boolean;
  /**
   * When pairing unmatched removed/added rows into "modified" rows, the
   * minimum share of equal cells required. Default 0.5.
   */
  rowSimilarity?: number;
  /**
   * Try to pair removed and added columns with near-identical contents as
   * renames. Default true.
   */
  detectRenames?: boolean;
}

export interface DiffStats {
  rowsAdded: number;
  rowsRemoved: number;
  rowsModified: number;
  rowsEqual: number;
  cellsChanged: number;
  columnsAdded: number;
  columnsRemoved: number;
  columnsRenamed: number;
}

/**
 * The result of diffing two tables. Row data lives in typed arrays so a
 * million-row diff costs a few megabytes, not a few hundred.
 */
export interface TableDiff {
  oldTable: Table;
  newTable: Table;
  columns: AlignedColumn[];
  /** Aligned row count. */
  rowCount: number;
  /** Old row index per aligned row, -1 when added. */
  oldRows: Int32Array;
  /** New row index per aligned row, -1 when removed. */
  newRows: Int32Array;
  /** `RowStatus` per aligned row, encoded with `ROW_STATUS`. */
  status: Uint8Array;
  /** Changed-cell bitmask per modified row, keyed by aligned row index. */
  changedCells: Map<number, Uint8Array>;
  /** Key columns used for matching, empty when aligned by sequence. */
  keyColumns: string[];
  stats: DiffStats;
  options: Required<DiffOptions>;
}

export const ROW_STATUS = { equal: 0, added: 1, removed: 2, modified: 3 } as const;
export const ROW_STATUS_NAMES: RowStatus[] = ['equal', 'added', 'removed', 'modified'];

export interface SheetDiff {
  name: string;
  status: 'equal' | 'modified' | 'added' | 'removed';
  diff: TableDiff;
}

export interface WorkbookDiff {
  sheets: SheetDiff[];
}
