/** Values the formula engine computes with, and Excel's coercion rules. */
import { Unsupported } from './parse.js';

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

/**
 * Parse text as Excel (en-US) does in arithmetic and VALUE: plain,
 * exponent and percent forms, thousands separators in their places
 * ("1,234.5"), a leading currency sign ("$1,234") and accounting
 * parentheses for negatives ("(5)").
 */
export function textToNumber(s: string): number | null {
  let t = s.trim();
  if (t === '') return null;
  let neg = false;
  if (t.startsWith('(') && t.endsWith(')')) {
    neg = true;
    t = t.slice(1, -1).trim();
  }
  const pct = t.endsWith('%');
  if (pct) t = t.slice(0, -1).trim();
  const m = /^([+-]?)\$?\s*((?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d*)?|\.\d+)(?:[eE]([+-]?\d+))?$/.exec(t);
  if (!m) return null;
  if (neg && m[1]) return null; // "(-5)" isn't a number
  let n = Number(`${m[1]}${m[2].replace(/,/g, '')}${m[3] !== undefined ? `e${m[3]}` : ''}`);
  if (!Number.isFinite(n)) return null;
  if (neg) n = -n;
  return pct ? n / 100 : n;
}

const MONTHS = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
const monthOf = (w: string) => {
  const i = MONTHS.indexOf(w.slice(0, 3).toLowerCase());
  return i >= 0 && (w.length === 3 || 'january february march april may june july august september october november december'.split(' ')[i].startsWith(w.toLowerCase())) ? i + 1 : 0;
};
const TIME_TEXT = /^(\d{1,4}):(\d{2})(?::(\d{2}(?:\.\d+)?))?\s*([AaPp][Mm])?$/;

/**
 * Date or time text as Excel reads it in arithmetic (`"2024-03-01" - "2024-02-01"`
 * is 29), as a serial in the 1900 date system: ISO dates (`2024-03-01`,
 * with an optional time), `15-Jan-2024`, `15 January 2024`, `Jan 15, 2024`,
 * and times (`14:30`, `2:30 PM`). Null for text that isn't a date. Slash
 * dates (`1/2/2024`) mean different days in different locales, so they
 * throw `Unsupported` rather than guess. (Arithmetic in a 1904-system
 * workbook would be 1,462 days off; VALUE and DATEVALUE correct for it.)
 */
export function dateTextToSerial(s: string): number | null {
  const t = s.trim().replace(/\s+/g, ' ');
  if (!t || !/\d/.test(t)) return null;
  const time = (x: string): number | null => {
    const m = TIME_TEXT.exec(x);
    if (!m) return null;
    let h = Number(m[1]);
    const min = Number(m[2]);
    const sec = m[3] ? Number(m[3]) : 0;
    if (m[4]) {
      if (h < 1 || h > 12) return null;
      h = (h % 12) + (m[4].toLowerCase() === 'pm' ? 12 : 0);
    }
    if (min > 59 || sec >= 60) return null; // hours past 23 are a duration: "24:00" is 1, "36:00" is 1.5
    return (h * 3600 + min * 60 + sec) / 86400;
  };
  const year = (y: string) => (y.length <= 2 ? (Number(y) < 30 ? 2000 : 1900) + Number(y) : Number(y));
  const date = (y: number, mo: number, d: number): number | null => {
    if (!(mo >= 1 && mo <= 12 && d >= 1 && y >= 1900 && y <= 9999)) return null;
    const last = new Date(Date.UTC(y, mo, 0)).getUTCDate() + (y === 1900 && mo === 2 ? 1 : 0); // Excel's 29-Feb-1900
    if (d > last) return null;
    return y === 1900 && mo === 2 && d === 29 ? 60 : serialFromParts(y, mo, d);
  };
  const onlyTime = time(t);
  if (onlyTime !== null) return onlyTime;
  if (/^\d{1,4}\/\d{1,2}\/\d{1,4}(?:\s|$)/.test(t)) throw new Unsupported('locale-dependent date text');
  let m: RegExpExecArray | null;
  let day: number | null = null;
  let rest = '';
  if ((m = /^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](.+))?$/.exec(t))) {
    day = date(Number(m[1]), Number(m[2]), Number(m[3]));
    rest = m[4] ?? '';
  } else if ((m = /^(\d{1,2})[- ]([A-Za-z]+)[- ](\d{2,4})(?: (.+))?$/.exec(t))) {
    day = monthOf(m[2]) ? date(year(m[3]), monthOf(m[2]), Number(m[1])) : null;
    rest = m[4] ?? '';
  } else if ((m = /^([A-Za-z]+) (\d{1,2}),? (\d{4})(?: (.+))?$/.exec(t))) {
    day = monthOf(m[1]) ? date(Number(m[3]), monthOf(m[1]), Number(m[2])) : null;
    rest = m[4] ?? '';
  } else return null;
  if (day === null) return null;
  if (!rest) return day;
  const tm = time(rest);
  return tm === null ? null : day + tm;
}

/** Number coercion for arithmetic: blank → 0, TRUE → 1, numeric or date text → number. */
export function toNumber(v: Scalar): number | XlError {
  if (typeof v === 'number') return v;
  if (v === null) return 0;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (isError(v)) return v;
  const n = textToNumber(v) ?? dateTextToSerial(v);
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
  // More than 15 digits before the point can't show in 15 significant digits: scientific (1E+20).
  const s = Math.abs(n) >= 1e15 ? Number(n.toPrecision(15)).toExponential() : String(Number(n.toPrecision(15)));
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

// Excel orders text with the system collation, not by code point: case is
// ignored, accents count, and punctuation sorts before digits and letters
// ("~" < "b"). For plain letters, digits and spaces both orders agree.
const PLAIN = /^[a-z0-9 ]*$/;
let collator: Intl.Collator | null = null;
/** Excel's text comparison: -1, 0 or 1. */
export function compareText(a: string, b: string): number {
  const x = a.toLowerCase();
  const y = b.toLowerCase();
  if (x === y) return 0;
  if (PLAIN.test(x) && PLAIN.test(y)) return x < y ? -1 : 1;
  collator ??= new Intl.Collator('en', { sensitivity: 'accent', usage: 'sort' });
  const c = collator.compare(x, y);
  return c < 0 ? -1 : c > 0 ? 1 : x < y ? -1 : 1; // the collator calls some distinct strings equal; keep a stable order
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
  if (typeof a === 'string') return compareText(a, b as string);
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
    if (s === 0) return { y: 1900, m: 1, d: 0, dow: 6 }; // Excel's "1900-01-00"
    s += 1;
  }
  const dt = new Date(EPOCH + s * DAY);
  return { y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1, d: dt.getUTCDate(), dow: dt.getUTCDay() };
}

const ISO = /^(\d{4})-(\d{2})-(\d{2})(?: (\d{2}):(\d{2}):(\d{2}(?:\.\d+)?))?$/;
const TIME = /^(\d{2}):(\d{2}):(\d{2}(?:\.\d+)?)$/;

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
 * Tables keep dates as ISO text, so text that reads as an ISO date (a TEXT
 * result, say) comes back from the table as a serial; either side matches.
 */
export function sameResult(a: Scalar, b: Scalar, date1904 = false): boolean {
  if (a === b) return true;
  if (typeof a === 'string' && typeof b === 'number') a = isoToSerial(a, date1904) ?? a;
  else if (typeof a === 'number' && typeof b === 'string') b = isoToSerial(b, date1904) ?? b;
  if (typeof a === 'number' && typeof b === 'number') {
    const d = Math.abs(a - b);
    return d <= 1e-9 * Math.max(Math.abs(a), Math.abs(b)) || d <= 1e-14;
  }
  if (isError(a) && isError(b)) return a.code === b.code;
  // An empty-text result is cached as "" or omitted.
  if ((a === '' && b === null) || (a === null && b === '')) return true;
  return false;
}
