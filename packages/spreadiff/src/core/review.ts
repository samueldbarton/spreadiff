/**
 * Accepting and rejecting changes. A review keeps the new version by
 * default; each rejected change goes back to the old one. Apply the
 * decisions to get the merged table ({@link applyDecisions}), or the
 * merged .xlsx (`applyDecisionsToXlsx` in `spreadiff/xlsx`).
 */
import { ROW_STATUS, type CellValue, type Table, type TableDiff } from './types.js';
import { changeKind } from './kind.js';
import { externalSource } from './external.js';

/**
 * Whether a changed cell is a decision to make: an edited value or
 * formula, or a value from outside data (a Bloomberg refresh). A
 * downstream result isn't: it follows its inputs, so to undo it, reject
 * the edits that moved it.
 */
export function isDecidable(diff: TableDiff, row: number, col: number): boolean {
  if (changeKind(diff, row, col) !== 'computed') return true;
  const ni = diff.newRows[row];
  const c = diff.columns[col];
  const f = ni >= 0 && c.newIndex >= 0 ? diff.newTable.formulas?.[ni]?.[c.newIndex] : null;
  return !!f && externalSource(f) !== null;
}

/** What to do with a change: keep the new version, or go back to the old. */
export type Decision = 'accept' | 'reject';

/**
 * Decisions for one sheet's diff. Anything not listed is accepted.
 * - `cells`: changed cells of rows in both tables, by aligned cell key `row * diff.columns.length + col`;
 * - `rows`: added and removed rows, by aligned row (rejecting an added row drops it, a removed one restores it);
 * - `columns`: added, removed and renamed columns, by index into `diff.columns`.
 */
export interface SheetDecisions {
  /** Changed cells, by aligned cell key `row * diff.columns.length + col`. */
  cells: Map<number, Decision>;
  /** Added and removed rows, by aligned row. */
  rows: Map<number, Decision>;
  /** Added, removed and renamed columns, by index into `diff.columns`. */
  columns: Map<number, Decision>;
}

/** No decisions yet: everything accepted. */
export const noDecisions = (): SheetDecisions => ({ cells: new Map(), rows: new Map(), columns: new Map() });

/**
 * Decide every change in a block of aligned rows and columns at once (a
 * selected range): its edited cells (not downstream results; see
 * {@link isDecidable}), and its added or removed rows. Returns new
 * decisions; the input isn't changed. `decision: undefined` clears them.
 */
export function decideRange(diff: TableDiff, decisions: SheetDecisions, rows: number[], columns: number[], decision: Decision | undefined): SheetDecisions {
  const next: SheetDecisions = { cells: new Map(decisions.cells), rows: new Map(decisions.rows), columns: new Map(decisions.columns) };
  const w = diff.columns.length;
  const set = <K,>(m: Map<K, Decision>, k: K) => (decision ? m.set(k, decision) : m.delete(k));
  for (const r of rows) {
    const st = diff.status[r];
    if (st === ROW_STATUS.added || st === ROW_STATUS.removed) set(next.rows, r);
    else if (st === ROW_STATUS.modified) {
      const mask = diff.changedCells.get(r);
      for (const c of columns) if (mask?.[c] && isDecidable(diff, r, c)) set(next.cells, r * w + c);
    }
  }
  // Every row of a column selected: decide the column too, if it was added, removed or renamed.
  if (rows.length === diff.rowCount) for (const c of columns) if (diff.columns[c].status !== 'equal') set(next.columns, c);
  return next;
}

/** How a cell, row or column was decided (accepted when not listed). */
export function decisionOf(diff: TableDiff, decisions: SheetDecisions | undefined, row: number, col: number): Decision {
  if (!decisions) return 'accept';
  const st = diff.status[row];
  if (st === ROW_STATUS.added || st === ROW_STATUS.removed) return decisions.rows.get(row) ?? 'accept';
  const c = diff.columns[col];
  if (c && c.status !== 'equal' && c.status !== 'renamed' && decisions.columns.has(col)) return decisions.columns.get(col)!;
  return decisions.cells.get(row * diff.columns.length + col) ?? 'accept';
}

/** Counts of decided changes, and of changes there are to decide (edits, rows, columns). */
export function decisionStats(diff: TableDiff, decisions: SheetDecisions): { accepted: number; rejected: number; changes: number } {
  let changes = 0;
  for (let r = 0; r < diff.rowCount; r++) {
    const st = diff.status[r];
    if (st === ROW_STATUS.added || st === ROW_STATUS.removed) changes++;
    else if (st === ROW_STATUS.modified) diff.changedCells.get(r)?.forEach((v, c) => (changes += v && isDecidable(diff, r, c) ? 1 : 0));
  }
  for (const c of diff.columns) if (c.status === 'added' || c.status === 'removed' || c.status === 'renamed') changes++;
  let rejected = 0;
  let accepted = 0;
  for (const m of [decisions.cells, decisions.rows, decisions.columns]) for (const d of m.values()) d === 'reject' ? rejected++ : accepted++;
  return { accepted, rejected, changes };
}

/**
 * The table with the decisions applied: the new table, with rejected cell
 * changes back to their old values (and formulas), rejected added rows
 * dropped, rejected removed rows restored where they were, and rejected
 * column changes undone. Formula results aren't recalculated: for a
 * workbook, `applyDecisionsToXlsx` writes the merged file with its results
 * recalculated.
 *
 * ```ts
 * import { applyDecisions, diffTables, noDecisions, parseCsv } from 'spreadiff';
 *
 * const diff = diffTables(parseCsv('id,price\n1,10'), parseCsv('id,price\n1,12'));
 * const d = noDecisions();
 * d.cells.set(0 * diff.columns.length + 1, 'reject'); // keep the old price
 * applyDecisions(diff, d).rows; // [[1, 10]]
 * ```
 */
export function applyDecisions(diff: TableDiff, decisions: SheetDecisions): Table {
  const w = diff.columns.length;
  const keepCol = diff.columns.map((c, ci) => {
    const d = decisions.columns.get(ci);
    if (c.status === 'added') return d !== 'reject';
    if (c.status === 'removed') return d === 'reject';
    return true;
  });
  // New-table order; a restored column goes back after the old column it followed.
  const cols = diff.columns.map((_, i) => i).filter((i) => keepCol[i] && diff.columns[i].newIndex >= 0);
  cols.sort((x, y) => diff.columns[x].newIndex - diff.columns[y].newIndex);
  const restored = diff.columns.map((_, i) => i).filter((i) => keepCol[i] && diff.columns[i].newIndex < 0);
  for (const ci of restored.sort((x, y) => diff.columns[x].oldIndex - diff.columns[y].oldIndex)) {
    const k = diff.columns[ci].oldIndex;
    let at = 0;
    cols.forEach((cj, j) => {
      const o = diff.columns[cj].oldIndex;
      if (o >= 0 && o < k) at = j + 1;
    });
    cols.splice(at, 0, ci);
  }
  const columns = cols.map((ci) => {
    const c = diff.columns[ci];
    return (c.status === 'renamed' || c.status === 'removed') && decisions.columns.get(ci) === 'reject' ? (c.oldName ?? c.name) : c.name;
  });
  const oldF = diff.oldTable.formulas;
  const newF = diff.newTable.formulas;
  const hasFormulas = !!(oldF || newF);
  const rows: CellValue[][] = [];
  const formulas: (string | null)[][] = [];
  const cell = (side: 'old' | 'new', r: number, ci: number): [CellValue, string | null] => {
    const c = diff.columns[ci];
    const ri = side === 'old' ? diff.oldRows[r] : diff.newRows[r];
    const idx = side === 'old' ? c.oldIndex : c.newIndex;
    if (ri < 0 || idx < 0) return [null, null];
    const t = side === 'old' ? diff.oldTable : diff.newTable;
    return [t.rows[ri][idx] ?? null, (side === 'old' ? oldF : newF)?.[ri]?.[idx] ?? null];
  };
  for (let r = 0; r < diff.rowCount; r++) {
    const st = diff.status[r];
    const rd = decisions.rows.get(r);
    let side: 'old' | 'new';
    if (st === ROW_STATUS.added) {
      if (rd === 'reject') continue;
      side = 'new';
    } else if (st === ROW_STATUS.removed) {
      if (rd !== 'reject') continue;
      side = 'old';
    } else side = 'new';
    const values: CellValue[] = [];
    const fs: (string | null)[] = [];
    for (const ci of cols) {
      const c = diff.columns[ci];
      // A column that exists only on the other side reads from there.
      let from: 'old' | 'new' = side;
      if (c.status === 'removed') from = 'old';
      else if (c.status === 'added') from = 'new';
      else if (st === ROW_STATUS.modified && diff.changedCells.get(r)?.[ci] && decisions.cells.get(r * w + ci) === 'reject') from = 'old';
      const [v, f] = cell(from, r, ci);
      values.push(v);
      fs.push(f);
    }
    rows.push(values);
    formulas.push(fs);
  }
  const t: Table = { name: diff.newTable.name || diff.oldTable.name, columns, rows };
  if (hasFormulas) t.formulas = formulas;
  return t;
}

/**
 * A table as CSV (RFC 4180): quoted where needed, `\r\n` between rows.
 * Formulas are written as their values. `delimiter` `'\t'` writes TSV.
 */
export function tableToCsv(table: Table, delimiter = ','): string {
  const needs = new RegExp(`["\r\n${delimiter === '\t' ? '\t' : delimiter.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}]`);
  const esc = (v: CellValue) => {
    const s = v === null ? '' : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : String(v);
    return needs.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
  };
  return [table.columns, ...table.rows].map((r) => r.map(esc).join(delimiter)).join('\r\n') + '\r\n';
}
