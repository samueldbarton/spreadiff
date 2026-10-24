/**
 * Sheets as the readers see them (sparse grids in sheet coordinates) and
 * how they become tables: header detection, trimming, merged and array
 * blocks. Shared by the XLSX, XLSB, XLS and ODS readers; not public API.
 */
import { columnLabel, uniqueNames } from './csv.js';
import { formulaRefs, scopedNames } from './formula.js';
import type { CellValue, Table, TableDefinition } from './types.js';

/**
 * Decide, per sheet, whether its first row is a header: all text, no
 * formulas, and not read by any formula in the workbook. A model's first
 * row often holds inputs; treating them as column names would hide edits
 * to them and every change they cause.
 */
function detectHeaders(raws: RawSheet[], names: Record<string, string>, tables: TableDefinition[], localNames: Map<string, Record<string, string>>): boolean[] {
  const candidate = new Map<string, number>(); // sheet name (upper) → index
  const result = raws.map((r, i) => {
    const ok = headerLike(r);
    if (ok) candidate.set(r.name.toUpperCase(), i);
    return ok;
  });
  if (!candidate.size) return result;
  const sheets = raws.map((r) => r.name);
  // Only formulas that could read a candidate row need tokenizing: a
  // reference to sheet row n spells it after a column letter or $ (A1, $B$1,
  // A1:A9) or as a row range (1:1). Names and structured references can
  // hide it, so formulas that may use them are always read. Tokenizing every
  // formula made parsing 2× slower on a 10,000-formula model.
  // A range covering row n starts at n or above, so every row number up to
  // the lowest candidate counts; past 64 of them, read everything.
  const top = Math.max(...[...candidate.values()].map((i) => raws[i].first + 1));
  const rowNumbers = Array.from({ length: top }, (_, i) => i + 1).join('|');
  const spells = top <= 64 ? new RegExp(`(?:[A-Za-z$](?:${rowNumbers})(?!\\d))|(?:(?:^|[^\\w.$])\\$?(?:${rowNumbers}):)`) : /./;
  const nameList = [...Object.keys(names), ...[...localNames.values()].flatMap((m) => Object.keys(m))].map((n) => n.toUpperCase());
  const scoped = raws.map((r) => (localNames.size ? scopedNames(names, raws.map((x) => ({ name: x.name, names: localNames.get(x.name) })), r.name) : names));
  const mayHide = (f: string) => f.includes('[') || (nameList.length > 0 && (nameList.length > 64 || nameList.some((n) => f.toUpperCase().includes(n))));
  raws.forEach((r, ri) => {
    for (let row = 0; row < r.formulaGrid.length && candidate.size; row++) {
      const fr = r.formulaGrid[row];
      if (!fr) continue;
      for (const f of fr) {
        if (!f) continue;
        if (!spells.test(f) && !mayHide(f)) continue;
        for (const rg of formulaRefs(f, scoped[ri], { tables, row, sheets }).ranges) {
          const key = (rg.sheet ?? r.name).toUpperCase();
          const i = candidate.get(key);
          if (i === undefined || rg.r0 < 0) continue; // whole columns (A:A) don't single out the header
          const first = raws[i].first;
          if (rg.r0 <= first && first <= rg.r1) {
            result[i] = false;
            candidate.delete(key);
          }
        }
        if (!candidate.size) break;
      }
    }
  });
  return result;
}

const ISO_DATE = /^\d{4}-\d{2}-\d{2}/;

function headerLike(r: RawSheet): boolean {
  if (r.maxCol < 0) return false;
  const row = r.grid[r.first];
  if (r.formulaGrid[r.first]?.some(Boolean)) return false;
  let n = 0;
  for (let c = r.minCol; c <= r.maxCol; c++) {
    const v = row?.[c];
    if (v === null || v === undefined || v === '') continue;
    // Whole years next to labels are column names (Region | 2023 | 2024);
    // a formula reading them still makes the row data (detectHeaders).
    if (isYear(v)) continue;
    // Other numbers, booleans and dates make it a data row.
    if (typeof v !== 'string' || ISO_DATE.test(v)) return false;
    n++;
  }
  return n > 0;
}

const isYear = (v: unknown) => typeof v === 'number' && Number.isInteger(v) && v >= 1900 && v <= 2100;

/** Drop namespace prefixes (`<x:c>` → `<c>`), as written by e.g. the OpenXML SDK. */
/**
 * Tables from raw sheets: with `'auto'`, decide per sheet whether its first
 * row is a header (see `detectHeaders`), then build each table. Shared by
 * the XLSX, ODS, XLS and XLSB readers.
 */
export function tablesFromRaw(
  raws: RawSheet[],
  header: boolean | 'auto',
  names: Record<string, string> = {},
  tables: TableDefinition[] = [],
  localNames: Map<string, Record<string, string>> = new Map(),
): Table[] {
  const headers = header === 'auto' ? detectHeaders(raws, names, tables, localNames) : raws.map(() => header);
  return raws.map((r, i) => {
    const t = buildTable(r, headers[i], header === 'auto');
    const own = localNames.get(r.name);
    if (own) t.names = own;
    return t;
  });
}

/** A raw sheet from sparse grids of values and formulas (`=…`), in sheet coordinates. */
export function rawSheet(
  name: string,
  grid: CellValue[][],
  formulaGrid: (string | null)[][],
  extra: Partial<Pick<RawSheet, 'arrays' | 'merges' | 'labelStyles' | 'formats' | 'hidden' | 'filters'>> = {},
): RawSheet {
  let minCol = Infinity;
  let maxCol = -1;
  for (const g of [grid, formulaGrid]) {
    for (const row of g) {
      if (!row) continue;
      for (let c = 0; c < row.length; c++) {
        const v = row[c];
        if (v === undefined || v === null) continue;
        if (c < minCol) minCol = c;
        if (c > maxCol) maxCol = c;
      }
    }
  }
  let first = 0;
  while (first < grid.length && !grid[first]?.some((v) => v !== null && v !== undefined) && !formulaGrid[first]?.some((f) => f)) first++;
  return { name, grid, formulaGrid, minCol, maxCol, first, arrays: [], merges: [], labelStyles: [], hidden: [], filters: [], ...extra };
}

/** A sheet as a reader sees it, in sheet coordinates, before it becomes a table. */
export interface RawSheet {
  /** Sheet name. */
  name: string;
  /** Values by sheet row and column (sparse). */
  grid: CellValue[][];
  /** Formula text (`=…`) by sheet row and column (sparse). */
  formulaGrid: (string | null)[][];
  /** First used column. */
  minCol: number;
  /** Last used column. */
  maxCol: number;
  /** First non-empty row. */
  first: number;
  /** Array-formula blocks (sheet coordinates). */
  arrays: { r0: number; c0: number; r1: number; c1: number }[];
  /** Merged ranges (sheet coordinates). */
  merges: { r0: number; c0: number; r1: number; c1: number }[];
  /** Indented or bold text cells (sheet coordinates). */
  labelStyles: NonNullable<Table['labelStyles']>;
  /** Number formats: codes (1-based) and each cell's code number (sheet coordinates). */
  formats?: { codes: string[]; grid: number[][]; date1904: boolean };
  /** Hidden rows (sheet rows, ascending). */
  hidden: number[];
  /** Data rows of active filters (sheet rows, inclusive), from the sheet's autoFilter; tables' are added later. */
  filters: { r0: number; r1: number }[];
}

/** Data rows under an autoFilter that filters something (it has a filterColumn). */
/**
 * A first row of labels that 'auto' kept as data (formulas read it, or a
 * parameter sits among the labels): mostly text, at least two labels.
 */
function labelRow(r: RawSheet): boolean {
  const row = r.grid[r.first];
  let text = 0;
  let filled = 0;
  for (let c = r.minCol; c <= r.maxCol; c++) {
    const v = row?.[c];
    if (v === null || v === undefined || v === '') continue;
    filled++;
    if (typeof v === 'string' && !ISO_DATE.test(v) && !r.formulaGrid[r.first]?.[c]) text++;
  }
  return text >= 2 && text * 2 >= filled;
}

function buildTable(raw: RawSheet, header: boolean, auto = false): Table {
  const { name, grid, formulaGrid, minCol, maxCol, first, arrays, merges, labelStyles, formats, hidden, filters } = raw;
  if (maxCol < 0) return { name, columns: [], rows: [], formulas: [], origin: { row: 0, col: 0 }, ...(merges.length ? { merges } : {}) };
  const width = maxCol - minCol + 1;
  const dense = <T,>(src: (T | null)[] | undefined): (T | null)[] => {
    const out = new Array<T | null>(width);
    for (let c = 0; c < width; c++) out[c] = src?.[minCol + c] ?? null;
    return out;
  };
  // A row of labels kept as data still names the columns, so key detection
  // and the grid see "id", "price" rather than A, B: it stays a data row too.
  const columns = header
    ? uniqueNames(dense(grid[first]).map((v) => (v === null ? '' : String(v))))
    : auto && labelRow(raw)
      ? uniqueNames(dense(grid[first]).map((v, c) => (typeof v === 'string' && v.trim() && !formulaGrid[first]?.[minCol + c] ? v : columnLabel(minCol + c))))
      : Array.from({ length: width }, (_, c) => columnLabel(minCol + c));
  const start = header ? first + 1 : first;
  const rows: CellValue[][] = [];
  // Formulas are always present (sparse) for XLSX, so a formula replaced by
  // a constant counts as a change.
  const formulas: (string | null)[][] = [];
  for (let r = start; r < grid.length; r++) {
    rows.push(dense(grid[r]));
    if (formulaGrid[r]) formulas[r - start] = dense(formulaGrid[r]);
  }
  // Trim trailing empty rows (rows holding only a formula are not empty).
  while (rows.length && !formulas[rows.length - 1] && rows[rows.length - 1].every((v) => v === null)) rows.pop();
  formulas.length = Math.min(formulas.length, rows.length);
  let numberFormats: Table['numberFormats'];
  if (formats) {
    const fr: (Uint16Array | undefined)[] = [];
    for (let r = start; r < start + rows.length; r++) {
      const src = formats.grid[r];
      if (!src) continue;
      const out = new Uint16Array(width);
      for (let c = 0; c < width; c++) out[c] = src[minCol + c] ?? 0;
      fr[r - start] = out;
    }
    if (fr.length) numberFormats = { codes: formats.codes, rows: fr, ...(formats.date1904 ? { date1904: true } : {}) };
  }
  return {
    name,
    columns,
    rows,
    formulas,
    origin: { row: start, col: minCol },
    ...(header ? { header: dense(grid[first]) } : {}),
    ...(arrays.length ? { arrays } : {}),
    ...(merges.length ? { merges } : {}),
    ...(labelStyles.length ? { labelStyles } : {}),
    ...(numberFormats ? { numberFormats } : {}),
    ...(hidden.length ? { hiddenRows: { rows: hidden, filtered: hidden.filter((r) => filters.some((f) => r >= f.r0 && r <= f.r1)) } } : {}),
  };
}
