/** Values the formula engine computes with, and Excel's coercion rules. */

/** An Excel error value (`#DIV/0!`, `#N/A`, …). */
export class XlError {
  constructor(readonly code: string) {}
  toString(): string {
    return this.code;
  }
}

export const ERR = {
  div0: new XlError('#DIV/0!'),
  na: new XlError('#N/A'),
  value: new XlError('#VALUE!'),
  ref: new XlError('#REF!'),
  num: new XlError('#NUM!'),
  name: new XlError('#NAME?'),
  null: new XlError('#NULL!'),
};
const BY_CODE = new Map(Object.values(ERR).map((e) => [e.code, e]));
export const errorOf = (code: string): XlError => BY_CODE.get(code) ?? new XlError(code);

/** A single value. `null` is an empty cell. */
export type Scalar = number | string | boolean | XlError | null;

/** A rectangle of cells on a sheet, read lazily (so lookups touch only what they use). */
export interface RangeVal {
  kind: 'range';
  sheet: string;
  r0: number;
  c0: number;
  r1: number;
  c1: number;
}

/** An in-memory 2-D array (array constants, results of array arithmetic). */
export interface ArrayVal {
  kind: 'array';
  rows: number;
  cols: number;
  data: Scalar[];
}

export type Value = Scalar | RangeVal | ArrayVal;

export const isError = (v: unknown): v is XlError => v instanceof XlError;
export const isRange = (v: Value): v is RangeVal => typeof v === 'object' && v !== null && (v as RangeVal).kind === 'range';
export const isArray = (v: Value): v is ArrayVal => typeof v === 'object' && v !== null && (v as ArrayVal).kind === 'array';

/** Parse text as Excel would in arithmetic ("1,000"? no: plain, percent and exponent forms). */
export function textToNumber(s: string): number | null {
  const t = s.trim();
  if (t === '') return null;
  const pct = t.endsWith('%');
  const body = pct ? t.slice(0, -1) : t;
  if (!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(body)) return null;
  const n = Number(body);
  return pct ? n / 100 : n;
}

/** Number coercion for arithmetic: blank → 0, TRUE → 1, numeric text → number. */
export function toNumber(v: Scalar): number | XlError {
  if (typeof v === 'number') return v;
  if (v === null) return 0;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (isError(v)) return v;
  const n = textToNumber(v);
  return n === null ? ERR.value : n;
}

export function toBool(v: Scalar): boolean | XlError {
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v !== 0;
  if (v === null) return false;
  if (isError(v)) return v;
  const u = v.toUpperCase();
  if (u === 'TRUE') return true;
  if (u === 'FALSE') return false;
  return ERR.value;
}

/** Excel's "General" rendering of a number, as `&` and TEXT-less concatenation use it. */
export function numberToText(n: number): string {
  if (Number.isInteger(n) && Math.abs(n) < 1e15) return String(n);
  const s = String(Number(n.toPrecision(15)));
  if (!s.includes('e')) return s;
  // 1e+21 → 1E+21, 1e-7 → 1E-07 (Excel pads the exponent to two digits).
  const [mant, exp] = s.split('e');
  return `${mant}E${exp.startsWith('-') ? '-' : '+'}${exp.replace(/^[+-]/, '').padStart(2, '0')}`;
}

export function toText(v: Scalar): string | XlError {
  if (typeof v === 'string') return v;
  if (v === null) return '';
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (isError(v)) return v;
  return numberToText(v);
}

/** Type rank for comparisons: numbers < text < logicals. */
const rank = (v: Scalar) => (typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : 2);

/** Excel comparison: -1, 0, 1. Blank compares as 0, "" or FALSE to match the other side. Text ignores case. */
export function compare(a: Scalar, b: Scalar): number {
  if (a === null && b === null) return 0;
  if (a === null) a = typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0;
  if (b === null) b = typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0;
  const ra = rank(a);
  const rb = rank(b);
  if (ra !== rb) return ra < rb ? -1 : 1;
  if (typeof a === 'number') return a < (b as number) ? -1 : a > (b as number) ? 1 : 0;
  if (typeof a === 'string') {
    const x = a.toLowerCase();
    const y = (b as string).toLowerCase();
    return x < y ? -1 : x > y ? 1 : 0;
  }
  return a === b ? 0 : a ? 1 : -1;
}

/** Excel's floating point display rounding: results agree with cached values to 15 digits. */
export function round15(n: number): number {
  return Number.isFinite(n) ? Number(n.toPrecision(15)) : n;
}

// ---- Dates: Excel serials, with the 1900 leap-year bug -------------------

const DAY = 86400000;
const EPOCH = Date.UTC(1899, 11, 30); // serial 0 in the (bug-compatible) 1900 system, for serials ≥ 61

export function serialFromParts(y: number, m: number, d: number, date1904 = false): number {
  // DATE() normalizes month/day overflow like JS Date.
  const ms = Date.UTC(y, m - 1, d);
  let s = Math.round((ms - EPOCH) / DAY);
  if (date1904) return s - 1462;
  if (s < 61) s -= 1; // before 1900-03-01: Excel counts the fictitious 1900-02-29
  return s;
}

export function partsFromSerial(serial: number, date1904 = false): { y: number; m: number; d: number; dow: number } {
  let s = Math.floor(serial);
  if (date1904) s += 1462;
  else if (s < 61) {
    if (s === 60) return { y: 1900, m: 2, d: 29, dow: 3 };
    s += 1;
  }
  const dt = new Date(EPOCH + s * DAY);
  return { y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1, d: dt.getUTCDate(), dow: dt.getUTCDay() };
}

const ISO = /^(\d{4})-(\d{2})-(\d{2})(?: (\d{2}):(\d{2}):(\d{2}))?$/;
const TIME = /^(\d{2}):(\d{2}):(\d{2})$/;

/** The parser's ISO rendering of a date-formatted cell, back to its serial (or null if it isn't one). */
export function isoToSerial(s: string, date1904 = false): number | null {
  let m = ISO.exec(s);
  if (m) {
    const day = serialFromParts(Number(m[1]), Number(m[2]), Number(m[3]), date1904);
    const t = m[4] ? (Number(m[4]) * 3600 + Number(m[5]) * 60 + Number(m[6])) / 86400 : 0;
    return day + t;
  }
  m = TIME.exec(s);
  if (m) return (Number(m[1]) * 3600 + Number(m[2]) * 60 + Number(m[3])) / 86400;
  return null;
}

/**
 * Is a recalculated value the same as a cached one? Engines differ in the
 * last bits (inputs cached to 15 digits, other SIN implementations), so
 * numbers get 1e-9 relative slack; a stale value is off by far more.
 */
export function sameResult(a: Scalar, b: Scalar): boolean {
  if (a === b) return true;
  if (typeof a === 'number' && typeof b === 'number') {
    const d = Math.abs(a - b);
    return d <= 1e-9 * Math.max(Math.abs(a), Math.abs(b)) || d <= 1e-14;
  }
  if (isError(a) && isError(b)) return a.code === b.code;
  // An empty-text result is cached as "" or omitted.
  if ((a === '' && b === null) || (a === null && b === '')) return true;
  return false;
}
