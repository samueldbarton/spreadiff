/**
 * Formula values a file doesn't have. Files written by tools that don't
 * calculate (openpyxl, many exporters) store formulas without cached
 * results; diffed as they are, every formula would read as blank.
 */
import { Engine, Unsupported } from './engine/evaluate.js';
import { tableHost } from './engine/host.js';
import { isError, type Scalar } from './engine/values.js';
import type { CellValue, Workbook } from './types.js';

/** What {@link computeMissingValues} did. */
export interface MissingValues {
  /** Formula cells that had no cached value and were computed. */
  computed: number;
  /** Formula cells with no cached value that couldn't be computed (unsupported function, circular reference); left empty. */
  failed: number;
}

/**
 * Compute, in place, the value of every formula cell that has none (a
 * formula with a null value). Values the file does have are never changed.
 * `parseXlsx` calls this by default; see `ParseXlsxOptions.computeMissing`.
 */
export function computeMissingValues(workbook: Workbook): MissingValues {
  const result = { computed: 0, failed: 0 };
  const missing = new Set<string>();
  workbook.sheets.forEach((t, s) => {
    t.formulas?.forEach((row, r) => row?.forEach((f, c) => {
      if (f && (t.rows[r]?.[c] ?? null) === null) missing.add(`${s},${r},${c}`);
    }));
  });
  if (!missing.size) return result;
  const host = tableHost(workbook.sheets, { names: workbook.names, tables: workbook.tables, date1904: workbook.date1904 });
  const engine = new Engine(host);
  const state = new Map<string, 1 | 2>();
  const failed = new Set<string>();
  let depth = 0;
  const base = host.get;
  host.getRange = undefined; // read cell by cell, so a missing precedent is computed first
  const compute = (s: number, r: number, c: number): void => {
    const k = `${s},${r},${c}`;
    const st = state.get(k);
    if (st === 2) return;
    if (st === 1) throw new Unsupported('circular reference');
    if (depth > 2000) throw new Unsupported('dependency chain too deep');
    state.set(k, 1);
    depth++;
    const t = workbook.sheets[s];
    const o = t.origin ?? { row: 1, col: 0 };
    try {
      const { value } = engine.evaluate(t.formulas![r]![c]!, t.name, r + o.row, c + o.col, false);
      t.rows[r][c] = toCellValue(value);
      host.override(sheetKeyOf(s, r + o.row, c + o.col), value);
      result.computed++;
    } catch (e) {
      if (!(e instanceof Unsupported)) throw e;
      failed.add(k);
      result.failed++;
    } finally {
      depth--;
      state.set(k, 2);
    }
  };
  host.get = (sheet, row, col) => {
    const s = host.sheetIndex(sheet);
    if (s >= 0) {
      const t = workbook.sheets[s];
      const o = t.origin ?? { row: 1, col: 0 };
      const k = `${s},${row - o.row},${col - o.col}`;
      if (missing.has(k)) {
        if (state.get(k) !== 2) compute(s, row - o.row, col - o.col);
        // Reading a value that couldn't be computed would treat it as blank: fail instead.
        if (failed.has(k)) throw new Unsupported('depends on a value that could not be computed');
      }
    }
    return base(sheet, row, col);
  };
  for (const k of missing) {
    if (state.get(k) === 2) continue;
    const [s, r, c] = k.split(',').map(Number);
    try {
      compute(s, r, c);
    } catch (e) {
      if (!(e instanceof Unsupported)) throw e;
    }
  }
  return result;
}

const sheetKeyOf = (s: number, r: number, c: number) => s * 2 ** 34 + r * 2 ** 14 + c;

function toCellValue(v: Scalar): CellValue {
  return isError(v) ? v.code : v;
}
