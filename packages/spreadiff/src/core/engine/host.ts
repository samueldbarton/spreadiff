/** An {@link EngineHost} over parsed tables, with an override layer for recalculated values. */
import type { CellValue, Table, TableDefinition } from '../types.js';
import { scopedNames } from '../formula.js';
import type { EngineHost } from './evaluate.js';
import { errorOf, isoToSerial, type Scalar } from './values.js';

const ROW_MUL = 2 ** 14;
const SHEET_MUL = 2 ** 34;

/** Key of a cell in sheet coordinates: sheet index × 2^34 + row × 2^14 + col. */
export const sheetCellKey = (sheet: number, row: number, col: number) => sheet * SHEET_MUL + row * ROW_MUL + col;

export interface TableHost extends EngineHost {
  /** Sheet index by name (case-insensitive), or -1. */
  sheetIndex(name: string): number;
  /** Values that replace cached ones (recalculated cells), by {@link sheetCellKey}. */
  readonly overrides: ReadonlyMap<number, Scalar>;
  /** Replace a cell's value (`undefined` restores the cached one). */
  override(key: number, value: Scalar | undefined): void;
  /** The cached value as stored (dates as serials). */
  cached(sheet: number, row: number, col: number): Scalar;
  /** The formula of a cell, or null. */
  formula(sheet: number, row: number, col: number): string | null;
}

const ERROR_TEXT = /^#(?:VALUE!|N\/A|REF!|DIV\/0!|NUM!|NAME\?|NULL!|SPILL!|CALC!)$/;

/**
 * A parsed cell as the engine sees it: error text becomes an error value,
 * and the parser's ISO rendering of a date cell goes back to its serial.
 */
export function engineValue(v: CellValue | undefined, date1904: boolean): Scalar {
  if (v === undefined || v === null || typeof v !== 'string') return v ?? null;
  const c = v.charCodeAt(0);
  if (c === 35 && ERROR_TEXT.test(v)) return errorOf(v);
  if (c >= 48 && c <= 57 && v.length >= 8) {
    const s = isoToSerial(v, date1904);
    if (s !== null) return s;
  }
  return v;
}

export function tableHost(
  tables: Table[],
  extra: { names?: Record<string, string>; tables?: TableDefinition[]; date1904?: boolean } = {},
): TableHost {
  const byName = new Map<string, number>();
  tables.forEach((t, i) => {
    byName.set(t.name, i);
    if (!byName.has(t.name.toUpperCase())) byName.set(t.name.toUpperCase(), i);
  });
  const date1904 = extra.date1904 ?? false;
  const overrides = new Map<number, Scalar>();
  const hiddenSets: (Map<number, 1 | 2> | null | undefined)[] = [];
  const scoped = new Map<string, Record<string, string>>();
  // Formula references usually spell the sheet exactly as named: try that first.
  const index = (name: string) => byName.get(name) ?? byName.get(name.toUpperCase()) ?? -1;
  const cachedIn = (t: Table, r: number, c: number): Scalar => {
    const o = t.origin ?? { row: 1, col: 0 };
    const tr = r - o.row;
    const tc = c - o.col;
    if (tc < 0 || tc >= t.columns.length) return null;
    if (tr === -1 && t.header) return engineValue(t.header[tc], date1904);
    if (tr < 0 || tr >= t.rows.length) return null;
    return engineValue(t.rows[tr][tc], date1904);
  };
  // Changes since a version, for patching cached ranges (bounded; older versions re-read).
  const LOG = 4096;
  const log: { sheet: string; row: number; col: number }[] = [];
  let logStart = 0; // version of log[0]
  const host: TableHost = {
    sheets: tables.map((t) => t.name),
    names: extra.names,
    namesFor: tables.some((t) => t.names)
      ? (sheet) => {
          let m = scoped.get(sheet);
          if (!m) scoped.set(sheet, (m = scopedNames(extra.names, tables, sheet) ?? {}));
          return m;
        }
      : undefined,
    tables: extra.tables,
    date1904,
    overrides,
    version: 0,
    override(k, v) {
      if (v === undefined) overrides.delete(k);
      else overrides.set(k, v);
      const s = Math.floor(k / 2 ** 34);
      const rest = k - s * 2 ** 34;
      if (log.length >= LOG) {
        // Too many to patch with: start over (older cached ranges get re-read).
        log.length = 0;
        logStart = host.version!;
      }
      log.push({ sheet: tables[s].name, row: Math.floor(rest / 2 ** 14), col: rest % 2 ** 14 });
      host.version!++;
    },
    changesSince(version) {
      if (version < logStart) return null;
      return log.slice(version - logStart);
    },
    sheetIndex: index,
    rowHidden: tables.some((t) => t.hiddenRows)
      ? (sheet, row) => {
          const s = index(sheet);
          if (s < 0) return 0;
          let h = hiddenSets[s];
          if (h === undefined) {
            const hr = tables[s].hiddenRows;
            h = hiddenSets[s] = hr ? new Map([...hr.rows.map((r) => [r, 1] as const), ...hr.filtered.map((r) => [r, 2] as const)]) : null;
          }
          return h?.get(row) ?? 0;
        }
      : undefined,
    formulaAt(sheet, r, c) {
      const s = index(sheet);
      return s < 0 ? null : host.formula(s, r, c);
    },
    cached: (s, r, c) => cachedIn(tables[s], r, c),
    formula(s, r, c) {
      const t = tables[s];
      const o = t.origin ?? { row: 1, col: 0 };
      return t.formulas?.[r - o.row]?.[c - o.col] ?? null;
    },
    get(sheet, r, c) {
      const s = index(sheet);
      if (s < 0) return null;
      if (overrides.size) {
        const k = sheetCellKey(s, r, c);
        const v = overrides.get(k);
        if (v !== undefined || overrides.has(k)) return v ?? null;
      }
      return cachedIn(tables[s], r, c);
    },
    getRange(sheet, r0, c0, r1, c1) {
      const s = index(sheet);
      const out: Scalar[] = new Array((r1 - r0 + 1) * (c1 - c0 + 1));
      if (s < 0) return out.fill(null);
      const t = tables[s];
      const o = t.origin ?? { row: 1, col: 0 };
      let i = 0;
      for (let r = r0; r <= r1; r++) {
        const tr = r - o.row;
        const row = tr >= 0 && tr < t.rows.length ? t.rows[tr] : null;
        for (let c = c0; c <= c1; c++) {
          const tc = c - o.col;
          let v: Scalar;
          if (tc < 0 || tc >= t.columns.length) v = null;
          else if (row) v = engineValue(row[tc], date1904);
          else v = tr === -1 && t.header ? engineValue(t.header[tc], date1904) : null;
          if (overrides.size) {
            const k = sheetCellKey(s, r, c);
            if (overrides.has(k)) v = overrides.get(k)!;
          }
          out[i++] = v;
        }
      }
      return out;
    },
    extent(sheet) {
      const s = index(sheet);
      if (s < 0) return null;
      const t = tables[s];
      const o = t.origin ?? { row: 1, col: 0 };
      return { rows: o.row + t.rows.length, cols: o.col + t.columns.length };
    },
    arrayBlock(sheet, r, c) {
      const s = index(sheet);
      const blocks = s >= 0 ? tables[s].arrays : undefined;
      if (!blocks) return null;
      for (const b of blocks) if (r >= b.r0 && r <= b.r1 && c >= b.c0 && c <= b.c1) return { r0: b.r0, c0: b.c0 };
      return null;
    },
  };
  return host;
}
