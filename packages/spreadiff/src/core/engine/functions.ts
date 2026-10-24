/**
 * Excel functions for the engine. Each takes unevaluated argument nodes, so
 * IF/CHOOSE/IFERROR evaluate only the branch they use, and lookups read only
 * the cells they need (recorded through the context).
 *
 * Deliberately missing: volatile and environment functions (NOW, TODAY,
 * RAND, CELL, INFO), whose cached values can't be checked. A missing
 * function throws Unsupported and the cell falls back to static analysis.
 */
import type { Ctx } from './evaluate.js';
import { lookupName } from '../formula.js';
import { parseFormula, Unsupported, type Node } from './parse.js';
import {
  compare, dateTextToSerial, ERR, isArray, isError, isRange, partsFromSerial, serialFromParts, textToNumber, toBool, toNumber, toText,
  type RangeVal, type Scalar, type Value, XlError,
} from './values.js';

type Fn = (args: Node[], ctx: Ctx) => Value;

// ---- Argument helpers -------------------------------------------------------

const arity = (args: Node[], min: number, max = min) => {
  if (args.length < min || args.length > max) throw new Unsupported('argument count');
};
const sc = (ctx: Ctx, n: Node | undefined): Scalar => (n === undefined ? null : ctx.scalar(ctx.ev(n)));
const missing = (n: Node | undefined) => n === undefined || n.t === 'missing';
function num(ctx: Ctx, n: Node | undefined, dflt?: number): number | XlError {
  if (missing(n)) {
    if (dflt === undefined) return ERR.value;
    return dflt;
  }
  return toNumber(sc(ctx, n));
}
function bool(ctx: Ctx, n: Node | undefined, dflt: boolean): boolean | XlError {
  if (missing(n)) return dflt;
  return toBool(sc(ctx, n));
}
function text(ctx: Ctx, n: Node | undefined): string | XlError {
  return toText(sc(ctx, n));
}
/** First error among values, if any. */
const firstError = (...xs: unknown[]): XlError | undefined => xs.find(isError) as XlError | undefined;

/**
 * Numbers for aggregates: in ranges and arrays only numbers count (text,
 * logicals and blanks are skipped); direct arguments are coerced.
 */
function numbers(ctx: Ctx, args: Node[], opts: { countText?: boolean } = {}): number[] | XlError {
  const out: number[] = [];
  for (const a of args) {
    if (a.t === 'missing') continue;
    const v = ctx.ev(a);
    if (isRange(v) || isArray(v)) {
      for (const x of ctx.values(v)) {
        if (typeof x === 'number') out.push(x);
        else if (isError(x)) return x;
      }
    } else {
      if (isError(v)) return v;
      if (v === null) {
        if (opts.countText) continue;
        out.push(0);
        continue;
      }
      const n = toNumber(v);
      if (isError(n)) return n;
      out.push(n);
    }
  }
  return out;
}

const agg = (f: (xs: number[]) => number | XlError): Fn => (args, ctx) => {
  const xs = numbers(ctx, args);
  return isError(xs) ? xs : f(xs);
};

const sum = (xs: number[]) => {
  let s = 0;
  for (const x of xs) s += x;
  return s;
};
const mean = (xs: number[]) => (xs.length ? sum(xs) / xs.length : ERR.div0);
const variance = (xs: number[], sample: boolean) => {
  const n = xs.length;
  if (n < (sample ? 2 : 1)) return ERR.div0;
  const m = sum(xs) / n;
  let s = 0;
  for (const x of xs) s += (x - m) * (x - m);
  return s / (sample ? n - 1 : n);
};
const sqrtOr = (v: number | XlError) => (isError(v) ? v : Math.sqrt(v));

function roundHalfAway(x: number, digits: number): number {
  const f = 10 ** digits;
  // Excel rounds the decimal it displays: correct binary noise first (2.675 → 2.68).
  const y = Number((Math.abs(x) * f).toPrecision(15));
  return (Math.sign(x) * Math.round(y)) / f;
}

const math1 = (f: (x: number) => number): Fn => (args, ctx) => {
  arity(args, 1);
  const x = num(ctx, args[0]);
  if (isError(x)) return x;
  const r = f(x);
  return Number.isFinite(r) ? r : ERR.num;
};

// ---- Criteria (SUMIF, COUNTIFS, …) -------------------------------------------

function wildcard(pattern: string): RegExp {
  let re = '';
  for (let i = 0; i < pattern.length; i++) {
    const ch = pattern[i];
    if (ch === '~' && i + 1 < pattern.length) re += escapeRe(pattern[++i]);
    else if (ch === '*') re += '[\\s\\S]*';
    else if (ch === '?') re += '[\\s\\S]';
    else re += escapeRe(ch);
  }
  return new RegExp(`^${re}$`, 'i');
}
const escapeRe = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
/** Equality as `=` tests it: text ignores case, numbers and logicals compare exactly. Cheaper than ordering. */
const same = (a: Scalar, b: Scalar) => (typeof a === 'string' && typeof b === 'string' ? a === b || a.toLowerCase() === b.toLowerCase() : compare(a, b) === 0);

export function criterion(c: Scalar): (v: Scalar) => boolean {
  if (typeof c === 'number') return (v) => (typeof v === 'number' ? v === c : typeof v === 'string' && textToNumber(v) === c);
  if (typeof c === 'boolean') return (v) => v === c;
  if (c === null) return (v) => v === null || v === '';
  if (isError(c)) return (v) => isError(v) && v.code === c.code;
  const m = /^(<=|>=|<>|<|>|=)?([\s\S]*)$/.exec(c)!;
  const op = m[1] ?? '';
  const rhs = m[2];
  // A number, or a date written as text (">2024-01-01"); slash dates are locale-dependent and throw Unsupported.
  const n = textToNumber(rhs) ?? (/\d/.test(rhs) && !/[*?~]/.test(rhs) ? dateTextToSerial(rhs) : null);
  if (op === '' || op === '=') {
    if (rhs === '') return op === '=' ? (v) => v === null || v === '' : (v) => v === null || v === '';
    if (n !== null) return (v) => (typeof v === 'number' ? v === n : typeof v === 'string' && textToNumber(v) === n);
    const lower = rhs.toLowerCase();
    if (/[*?~]/.test(rhs)) {
      const re = wildcard(rhs);
      return (v) => typeof v === 'string' && re.test(v);
    }
    if (lower === 'true' || lower === 'false') return (v) => v === (lower === 'true');
    return (v) => typeof v === 'string' && v.toLowerCase() === lower;
  }
  if (op === '<>') {
    if (rhs === '') return (v) => v !== null && v !== '';
    if (n !== null) return (v) => !(typeof v === 'number' && v === n);
    if (/[*?~]/.test(rhs)) {
      const re = wildcard(rhs);
      return (v) => !(typeof v === 'string' && re.test(v));
    }
    const lower = rhs.toLowerCase();
    return (v) => !(typeof v === 'string' && v.toLowerCase() === lower);
  }
  const test = (cmp: number) => (op === '<' ? cmp < 0 : op === '>' ? cmp > 0 : op === '<=' ? cmp <= 0 : cmp >= 0);
  if (n !== null) return (v) => typeof v === 'number' && test(compare(v, n));
  return (v) => typeof v === 'string' && test(compare(v, rhs));
}

/** Rows (as flat indexes) where every (range, criterion) pair matches. Records the criteria ranges. */
function matchingIndexes(ctx: Ctx, pairs: [Node, Node][]): { dims: { rows: number; cols: number }; idx: number[] } | XlError {
  let dims: { rows: number; cols: number } | null = null;
  let idx: number[] | null = null;
  for (const [rn, cn] of pairs) {
    const rv = ctx.ev(rn);
    const d = ctx.dims(rv);
    if (dims && (d.rows !== dims.rows || d.cols !== dims.cols)) return ERR.value;
    dims = d;
    const vals = ctx.values(rv);
    const test = criterion(sc(ctx, cn));
    const next: number[] = [];
    if (idx === null) {
      for (let i = 0; i < vals.length; i++) if (test(vals[i])) next.push(i);
    } else {
      for (const i of idx) if (test(vals[i])) next.push(i);
    }
    idx = next;
  }
  return { dims: dims!, idx: idx! };
}

/** Values of `target` at matched positions, reading only those cells. */
function pickAt(ctx: Ctx, target: Value, dims: { cols: number }, idx: number[]): Scalar[] {
  return idx.map((i) => ctx.at(target, Math.floor(i / dims.cols), i % dims.cols));
}

function ifsAggregate(kind: 'sum' | 'count' | 'avg' | 'max' | 'min'): Fn {
  return (args, ctx) => {
    const hasTarget = kind !== 'count';
    if (args.length < (hasTarget ? 3 : 2) || (args.length - (hasTarget ? 1 : 0)) % 2) throw new Unsupported('argument count');
    const pairs: [Node, Node][] = [];
    for (let i = hasTarget ? 1 : 0; i < args.length; i += 2) pairs.push([args[i], args[i + 1]]);
    const m = matchingIndexes(ctx, pairs);
    if (isError(m)) return m;
    if (kind === 'count') return m.idx.length;
    const target = ctx.ev(args[0]);
    const td = ctx.dims(target);
    if (td.rows !== m.dims.rows || td.cols !== m.dims.cols) return ERR.value;
    const xs: number[] = [];
    for (const v of pickAt(ctx, target, m.dims, m.idx)) {
      if (isError(v)) return v;
      if (typeof v === 'number') xs.push(v);
    }
    if (kind === 'sum') return sum(xs);
    if (kind === 'avg') return mean(xs);
    if (!xs.length) return 0;
    return kind === 'max' ? Math.max(...xs) : Math.min(...xs);
  };
}

/** SUMIF/AVERAGEIF(range, criterion, [sum_range]): sum_range takes range's shape from its top-left. */
function ifAggregate(kind: 'sum' | 'avg'): Fn {
  return (args, ctx) => {
    arity(args, 2, 3);
    const m = matchingIndexes(ctx, [[args[0], args[1]]]);
    if (isError(m)) return m;
    let target = ctx.ev(missing(args[2]) ? args[0] : args[2]);
    if (isRange(target)) target = { ...target, r1: target.r0 + m.dims.rows - 1, c1: target.c0 + m.dims.cols - 1 };
    const xs: number[] = [];
    for (const v of pickAt(ctx, target, m.dims, m.idx)) {
      if (isError(v)) return v;
      if (typeof v === 'number') xs.push(v);
    }
    return kind === 'sum' ? sum(xs) : mean(xs);
  };
}

// ---- Lookups -----------------------------------------------------------------

const indexKey = (v: Scalar): string | null =>
  typeof v === 'number' ? `n${v}` : typeof v === 'string' ? `s${v.toLowerCase()}` : typeof v === 'boolean' ? `b${v}` : null;

/** Position of `key` in `vals` (exact, with wildcards for text), or -1. Long lists get a reusable index. */
function exactIndex(vals: Scalar[], key: Scalar, wild: boolean, ctx?: Ctx): number {
  if (typeof key === 'string' && wild && /[*?~]/.test(key)) {
    const re = wildcard(key);
    return vals.findIndex((v) => typeof v === 'string' && re.test(v));
  }
  if (ctx && vals.length > 32) {
    let idx = ctx.engine.indexes.get(vals);
    if (!idx) {
      idx = new Map();
      for (let i = 0; i < vals.length; i++) {
        const k = indexKey(vals[i]);
        if (k !== null && !idx.has(k)) idx.set(k, i);
      }
      ctx.engine.indexes.set(vals, idx);
    }
    const k = indexKey(key);
    return k === null ? -1 : idx.get(k) ?? -1;
  }
  for (let i = 0; i < vals.length; i++) {
    const v = vals[i];
    if (v === null) continue;
    if (typeof v === typeof key && same(v, key)) return i;
  }
  return -1;
}

/** Approximate match on ascending data: last value <= key of the same type (Excel's binary search on sorted data). */
function approxIndex(vals: Scalar[], key: Scalar, descending = false): number {
  let lo = 0;
  let hi = vals.length - 1;
  let found = -1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    const v = vals[mid];
    // Values of another type are skipped by narrowing toward same-typed ones.
    const c = v === null || typeof v !== typeof key ? (typeof v === 'number' ? -1 : 1) : compare(v, key);
    const ok = descending ? c >= 0 : c <= 0;
    if (ok) {
      if (v !== null && typeof v === typeof key) found = mid;
      lo = mid + 1;
    } else hi = mid - 1;
  }
  return found;
}

function vhlookup(horizontal: boolean): Fn {
  return (args, ctx) => {
    arity(args, 3, 4);
    const key = sc(ctx, args[0]);
    if (isError(key)) return key;
    const table = ctx.ev(args[1]);
    const index = num(ctx, args[2]);
    if (isError(index)) return index;
    const approx = bool(ctx, args[3], true);
    if (isError(approx)) return approx;
    const d = ctx.dims(table);
    const k = Math.trunc(index);
    if (k < 1) return ERR.value;
    if (k > (horizontal ? d.rows : d.cols)) return ERR.ref;
    const keys = horizontal ? sliceRow(ctx, table, 0) : sliceCol(ctx, table, 0);
    const i = approx ? approxIndex(keys, key) : exactIndex(keys, key, true, ctx);
    if (i < 0) return ERR.na;
    return horizontal ? ctx.at(table, k - 1, i) : ctx.at(table, i, k - 1);
  };
}

function sliceCol(ctx: Ctx, v: Value, c: number): Scalar[] {
  if (isRange(v)) return ctx.values({ ...v, c0: v.c0 + c, c1: v.c0 + c });
  const d = ctx.dims(v);
  const out: Scalar[] = [];
  for (let r = 0; r < d.rows; r++) out.push(ctx.at(v, r, c));
  return out;
}
function sliceRow(ctx: Ctx, v: Value, r: number): Scalar[] {
  if (isRange(v)) return ctx.values({ ...v, r0: v.r0 + r, r1: v.r0 + r });
  const d = ctx.dims(v);
  const out: Scalar[] = [];
  for (let c = 0; c < d.cols; c++) out.push(ctx.at(v, r, c));
  return out;
}

/** A 1-D view (row or column) of a range or array, as values. */
function vector(ctx: Ctx, v: Value): Scalar[] | XlError {
  const d = ctx.dims(v);
  if (d.rows !== 1 && d.cols !== 1) return ERR.na;
  return ctx.values(v);
}

function matchFn(args: Node[], ctx: Ctx): Value {
  arity(args, 2, 3);
  const key = sc(ctx, args[0]);
  if (isError(key)) return key;
  const vals = vector(ctx, ctx.ev(args[1]));
  if (isError(vals)) return vals;
  const type = num(ctx, args[2], 1);
  if (isError(type)) return type;
  const i = type === 0 ? exactIndex(vals, key, true, ctx) : approxIndex(vals, key, type < 0);
  return i < 0 ? ERR.na : i + 1;
}

/** INDEX returns a reference, so only the cell finally used is read. */
function indexFn(args: Node[], ctx: Ctx): Value {
  arity(args, 1, 3);
  const src = ctx.ev(args[0]);
  if (isError(src)) return src;
  const d = ctx.dims(src);
  let r = num(ctx, args[1], 0);
  if (isError(r)) return r;
  let c = num(ctx, args[2], 0);
  if (isError(c)) return c;
  r = Math.trunc(r);
  c = Math.trunc(c);
  // One-dimensional source: a single index picks along it.
  if (missing(args[2]) && d.rows === 1 && d.cols > 1) {
    c = r;
    r = 1;
  }
  if (r < 0 || c < 0 || r > d.rows || c > d.cols) return ERR.ref;
  if (isRange(src)) {
    const r0 = r === 0 ? src.r0 : src.r0 + r - 1;
    const r1 = r === 0 ? src.r1 : r0;
    const c0 = c === 0 ? src.c0 : src.c0 + c - 1;
    const c1 = c === 0 ? src.c1 : c0;
    if (r === 0 && c === 0 && d.cols > 1 && d.rows > 1) return src;
    return { kind: 'range', sheet: src.sheet, r0, c0, r1, c1 };
  }
  if (r === 0 && c === 0) return src;
  if (r === 0 || c === 0) {
    const a = ctx.array(src);
    if (r === 0) return { kind: 'array', rows: a.rows, cols: 1, data: Array.from({ length: a.rows }, (_, i) => a.data[i * a.cols + c - 1]) };
    return { kind: 'array', rows: 1, cols: a.cols, data: a.data.slice((r - 1) * a.cols, r * a.cols) };
  }
  return ctx.at(src, r - 1, c - 1);
}

function xlookup(args: Node[], ctx: Ctx): Value {
  arity(args, 3, 6);
  const key = sc(ctx, args[0]);
  if (isError(key)) return key;
  const look = ctx.ev(args[1]);
  const ret = ctx.ev(args[2]);
  const mode = num(ctx, args[4], 0);
  const search = num(ctx, args[5], 1);
  if (isError(mode) || isError(search)) return ERR.value;
  if (search !== 1 && search !== -1) throw new Unsupported('XLOOKUP binary search');
  const vals = vector(ctx, look);
  if (isError(vals)) return vals;
  const order = vals.map((_, i) => i);
  if (search === -1) order.reverse();
  let hit = -1;
  if (mode === 0 || mode === 2) {
    const re = mode === 2 && typeof key === 'string' ? wildcard(key) : null;
    hit = order.find((i) => (re ? typeof vals[i] === 'string' && re.test(vals[i] as string) : vals[i] !== null && typeof vals[i] === typeof key && same(vals[i], key))) ?? -1;
  } else if (mode === -1 || mode === 1) {
    // Exact, else next smaller (-1) or next larger (1).
    let best = -1;
    for (const i of order) {
      const v = vals[i];
      if (v === null || typeof v !== typeof key) continue;
      const c = compare(v, key);
      if (c === 0) {
        best = i;
        break;
      }
      if (mode === -1 ? c < 0 : c > 0) {
        if (best < 0 || (mode === -1 ? compare(v, vals[best]) > 0 : compare(v, vals[best]) < 0)) best = i;
      }
    }
    hit = best;
  } else throw new Unsupported('XLOOKUP match mode');
  if (hit < 0) return missing(args[3]) ? ERR.na : ctx.ev(args[3]);
  const ld = ctx.dims(look);
  const rd = ctx.dims(ret);
  if (ld.cols === 1) {
    if (isRange(ret)) return { kind: 'range', sheet: ret.sheet, r0: ret.r0 + hit, r1: ret.r0 + hit, c0: ret.c0, c1: ret.c1 };
    return rd.cols === 1 ? ctx.at(ret, hit, 0) : ERR.value;
  }
  if (isRange(ret)) return { kind: 'range', sheet: ret.sheet, r0: ret.r0, r1: ret.r1, c0: ret.c0 + hit, c1: ret.c0 + hit };
  return rd.rows === 1 ? ctx.at(ret, 0, hit) : ERR.value;
}

function lookupFn(args: Node[], ctx: Ctx): Value {
  arity(args, 2, 3);
  const key = sc(ctx, args[0]);
  if (isError(key)) return key;
  const look = ctx.ev(args[1]);
  const vals = vector(ctx, look);
  if (isError(vals)) throw new Unsupported('LOOKUP array form');
  const i = approxIndex(vals, key);
  if (i < 0) return ERR.na;
  const res = missing(args[2]) ? look : ctx.ev(args[2]);
  const d = ctx.dims(res);
  return d.cols === 1 ? ctx.at(res, i, 0) : ctx.at(res, 0, i);
}

// ---- References ----------------------------------------------------------------

function refArg(ctx: Ctx, n: Node | undefined): RangeVal | XlError {
  if (missing(n)) return { kind: 'range', sheet: ctx.sheet, r0: ctx.row, c0: ctx.col, r1: ctx.row, c1: ctx.col };
  const v = ctx.ev(n!);
  if (isError(v)) return v;
  if (!isRange(v)) return ERR.value;
  return v;
}

function offsetFn(args: Node[], ctx: Ctx): Value {
  arity(args, 3, 5);
  const base = refArg(ctx, args[0]);
  if (isError(base)) return base;
  const dr = num(ctx, args[1]);
  const dc = num(ctx, args[2]);
  const h = num(ctx, args[3], base.r1 - base.r0 + 1);
  const w = num(ctx, args[4], base.c1 - base.c0 + 1);
  const e = firstError(dr, dc, h, w);
  if (e) return e;
  const r0 = base.r0 + Math.trunc(dr as number);
  const c0 = base.c0 + Math.trunc(dc as number);
  const hh = Math.trunc(h as number);
  const ww = Math.trunc(w as number);
  if (r0 < 0 || c0 < 0 || hh < 1 || ww < 1) return ERR.ref;
  return { kind: 'range', sheet: base.sheet, r0, c0, r1: r0 + hh - 1, c1: c0 + ww - 1 };
}

function indirectFn(args: Node[], ctx: Ctx): Value {
  arity(args, 1, 2);
  const t = text(ctx, args[0]);
  if (isError(t)) return t;
  const a1 = bool(ctx, args[1], true);
  if (a1 !== true) throw new Unsupported('INDIRECT R1C1');
  let node: Node;
  try {
    node = parseFormula(t);
  } catch {
    // A defined name?
    if (lookupName(ctx.host.namesFor ? ctx.host.namesFor(ctx.sheet) : ctx.host.names, t) !== undefined) node = { t: 'name', name: t };
    else return ERR.ref;
  }
  if (node.t !== 'ref' && node.t !== 'name' && !(node.t === 'binary' && node.op === ':')) return ERR.ref;
  const v = ctx.ev(node);
  return isRange(v) ? v : ERR.ref;
}

// ---- Dates ----------------------------------------------------------------------

const date1904 = (ctx: Ctx) => ctx.host.date1904 ?? false;
function datePart(part: 'y' | 'm' | 'd' | 'dow'): Fn {
  return (args, ctx) => {
    arity(args, part === 'dow' ? 1 : 1, part === 'dow' ? 2 : 1);
    const s = num(ctx, args[0]);
    if (isError(s)) return s;
    if (s < 0) return ERR.num;
    const p = partsFromSerial(s, date1904(ctx));
    if (part !== 'dow') return p[part];
    const type = num(ctx, args[1], 1);
    if (isError(type)) return type;
    // Serial 0..59 predates the bug-compatible calendar; Excel's 1900-01-01 was a "Sunday".
    const dow = partsFromSerial(s, date1904(ctx)).dow;
    if (type === 1) return dow + 1;
    if (type === 2) return ((dow + 6) % 7) + 1;
    if (type === 3) return (dow + 6) % 7;
    throw new Unsupported('WEEKDAY type');
  };
}

function edate(endOfMonth: boolean): Fn {
  return (args, ctx) => {
    arity(args, 2);
    const s = num(ctx, args[0]);
    const m = num(ctx, args[1]);
    const e = firstError(s, m);
    if (e) return e;
    const p = partsFromSerial(s as number, date1904(ctx));
    const months = Math.trunc(m as number);
    if (endOfMonth) return serialFromParts(p.y, p.m + months + 1, 0, date1904(ctx));
    const last = partsFromSerial(serialFromParts(p.y, p.m + months + 1, 0, date1904(ctx)), date1904(ctx)).d;
    return serialFromParts(p.y, p.m + months, Math.min(p.d, last), date1904(ctx));
  };
}

// ---- TEXT -----------------------------------------------------------------------

const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

/** A subset of Excel number formats: 0 # , . % literals, and y m d h s for dates. */
export function formatValue(v: number, format: string, d1904 = false): string {
  const sections = splitSections(format);
  let fmt = sections[0];
  let value = v;
  if (v < 0 && sections.length > 1) {
    fmt = sections[1];
    value = -v;
  } else if (v === 0 && sections.length > 2) fmt = sections[2];
  if (fmt === '') return ''; // an empty section hides the value
  const isDate = /[yd]|h|s|(?<![#0.,])m/i.test(fmt.replace(/"[^"]*"|\\./g, ''));
  if (isDate) return formatDate(value, fmt, d1904);
  return formatNumber(value, fmt, v < 0 && sections.length === 1);
}

/** Text through a format's text section: the fourth, or a lone section with @; otherwise unchanged. */
function formatText(v: string, format: string): string {
  const sections = splitSections(format);
  const fmt = sections.length > 3 ? sections[3] : sections.length === 1 && /@/.test(format.replace(/"[^"]*"/g, '')) ? sections[0] : null;
  if (fmt === null) return v;
  return fmt.replace(/"([^"]*)"|\\(.)|_.|\*.|@/g, (t, q, e) => (q !== undefined ? q : e !== undefined ? e : t[0] === '_' ? ' ' : t[0] === '*' ? '' : v));
}

function splitSections(f: string): string[] {
  const out: string[] = [];
  let cur = '';
  let q = false;
  for (const ch of f) {
    if (ch === '"') q = !q;
    if (ch === ';' && !q) {
      out.push(cur);
      cur = '';
    } else cur += ch;
  }
  out.push(cur);
  return out;
}

function formatNumber(v: number, fmt: string, negative: boolean): string {
  if (/[?/E]|\[/.test(fmt.replace(/"[^"]*"/g, ''))) throw new Unsupported('number format');
  const literal = (s: string) => s.replace(/"([^"]*)"/g, '$1').replace(/\\(.)/g, '$1');
  const m = /^([^0#,.%]*)([0#,]*)(?:\.([0#]*))?([^0#]*)$/.exec(fmt);
  if (!m) throw new Unsupported('number format');
  const [, pre, intPart, decPart = '', post] = m;
  let x = Math.abs(v);
  if (fmt.includes('%')) x *= 100;
  // Trailing commas scale by thousands.
  const scale = /,+$/.exec(intPart)?.[0].length ?? 0;
  x /= 1000 ** scale;
  const decimals = decPart.length;
  const required = decPart.replace(/#/g, '').length;
  let s = roundHalfAway(x, decimals).toFixed(decimals);
  if (decimals > required) s = s.replace(new RegExp(`0{1,${decimals - required}}$`), '').replace(/\.$/, '');
  let [i, d] = s.split('.');
  const minInt = intPart.replace(/[#,]/g, '').length;
  if (i === '0' && minInt === 0) i = '';
  i = i.padStart(minInt, '0');
  if (intPart.replace(/,+$/, '').includes(',')) i = i.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return (negative ? '-' : '') + literal(pre) + i + (d ? `.${d}` : '') + literal(post);
}

function formatDate(v: number, fmt: string, d1904: boolean): string {
  const p = partsFromSerial(v, d1904);
  // "ss.000" shows fractions of a second: round to that many decimals instead of whole seconds.
  const fd = Math.min(3, Math.max(0, ...[...fmt.matchAll(/s+\.(0+)/gi)].map((m) => m[1].length)));
  const scale = 10 ** fd;
  let units = Math.round((v - Math.floor(v)) * 86400 * scale);
  if (units >= 86400 * scale) units = 86400 * scale - 1;
  const secs = Math.floor(units / scale);
  const frac = fd ? String(units % scale).padStart(fd, '0') : '';
  const hh = Math.floor(secs / 3600);
  const mm = Math.floor((secs % 3600) / 60);
  const ss = secs % 60;
  const ampm = /AM\/PM|A\/P/i.test(fmt);
  // AM/PM and A/P before the literal-run rule, which mustn't swallow them ("ss AM/PM"): "a" is its own literal.
  const tokens = fmt.match(/"[^"]*"|\\.|AM\/PM|A\/P|y+|m+|d+|h+|s+|[aA]|[^ymdhsaA"\\]+/gi) ?? [];
  let out = '';
  for (let k = 0; k < tokens.length; k++) {
    const t = tokens[k];
    const l = t.toLowerCase();
    const prev = tokens.slice(0, k).reverse().find((x) => /^[ymdhs]/i.test(x))?.toLowerCase() ?? '';
    const nextTok = tokens.slice(k + 1).find((x) => /^[ymdhs]/i.test(x))?.toLowerCase() ?? '';
    if (t.startsWith('"')) out += t.slice(1, -1);
    else if (t.startsWith('\\')) out += t.slice(1);
    else if (/^am\/pm$/i.test(t)) out += hh < 12 ? 'AM' : 'PM'; // always upper case, as Excel writes it
    else if (/^a\/p$/i.test(t)) out += hh < 12 ? t[0] : t[2]; // the format's own case: a/p → "p", A/P → "P"
    else if (l[0] === 'y') out += l.length <= 2 ? String(p.y % 100).padStart(2, '0') : String(p.y);
    else if (l[0] === 'd') out += l.length === 1 ? p.d : l.length === 2 ? String(p.d).padStart(2, '0') : l.length === 3 ? DAYS[p.dow].slice(0, 3) : DAYS[p.dow];
    else if (l[0] === 'h') {
      const h = ampm ? hh % 12 || 12 : hh;
      out += l.length === 1 ? h : String(h).padStart(2, '0');
    } else if (l[0] === 's') out += l.length === 1 ? ss : String(ss).padStart(2, '0');
    else if (fd && /^\.0+/.test(t) && tokens[k - 1]?.toLowerCase().startsWith('s')) {
      const zeros = /^\.(0+)/.exec(t)![1].length;
      out += '.' + frac.slice(0, zeros).padEnd(zeros, '0') + t.slice(1 + zeros);
    }
    else if (l[0] === 'm') {
      // m after h or before s is minutes.
      if (prev.startsWith('h') || nextTok.startsWith('s')) out += l.length === 1 ? mm : String(mm).padStart(2, '0');
      else out += l.length === 1 ? p.m : l.length === 2 ? String(p.m).padStart(2, '0') : l.length === 3 ? MONTHS[p.m - 1].slice(0, 3) : l.length === 5 ? MONTHS[p.m - 1][0] : MONTHS[p.m - 1];
    } else out += t;
  }
  return out;
}

// ---- Financial ------------------------------------------------------------------

function pmt(rate: number, nper: number, pv: number, fv: number, type: number): number {
  if (rate === 0) return -(pv + fv) / nper;
  const f = (1 + rate) ** nper;
  return (-(fv + pv * f) * rate) / ((1 + rate * type) * (f - 1));
}
function fvOf(rate: number, nper: number, p: number, pv: number, type: number): number {
  if (rate === 0) return -(pv + p * nper);
  const f = (1 + rate) ** nper;
  return -(pv * f + (p * (1 + rate * type) * (f - 1)) / rate);
}
function ipmt(rate: number, per: number, nper: number, pv: number, fv: number, type: number): number {
  if (type === 1 && per === 1) return 0;
  // Interest on the balance left after per - 1 payments (negative for a loan, like Excel).
  const i = fvOf(rate, per - 1, pmt(rate, nper, pv, fv, type), pv, type) * rate;
  return type === 1 ? i / (1 + rate) : i;
}

const fin = (min: number, max: number, f: (a: number[]) => number): Fn => (args, ctx) => {
  arity(args, min, max);
  const xs: number[] = [];
  for (let i = 0; i < max; i++) {
    const v = num(ctx, args[i], 0);
    if (isError(v)) return v;
    xs.push(v);
  }
  const r = f(xs);
  return Number.isFinite(r) ? r : ERR.num;
};

function newton(f: (r: number) => number, guess: number): number | XlError {
  let r = guess;
  for (let k = 0; k < 100; k++) {
    const y = f(r);
    const h = 1e-7 * Math.max(1, Math.abs(r));
    const dy = (f(r + h) - y) / h;
    if (!Number.isFinite(dy) || dy === 0) return ERR.num;
    const next = r - y / dy;
    if (Math.abs(next - r) < 1e-12) return next;
    r = next;
  }
  return ERR.num;
}

// ---- The table --------------------------------------------------------------------

const passthrough = (f: (s: string, ...rest: number[]) => Scalar, nums = 0): Fn => (args, ctx) => {
  arity(args, 1, 1 + nums);
  const s = text(ctx, args[0]);
  if (isError(s)) return s;
  const rest: number[] = [];
  for (let i = 1; i < args.length; i++) {
    const n = num(ctx, args[i]);
    if (isError(n)) return n;
    rest.push(n);
  }
  return f(s, ...rest);
};

const isType = (test: (v: Scalar) => boolean): Fn => (args, ctx) => {
  arity(args, 1);
  return test(sc(ctx, args[0]));
};

export const FUNCTIONS: Record<string, Fn> = {
  // Logic (lazy branches)
  IF: (args, ctx) => {
    arity(args, 1, 3);
    const c = toBool(sc(ctx, args[0]));
    if (isError(c)) return c;
    if (c) return missing(args[1]) ? (args.length > 1 ? 0 : true) : ctx.ev(args[1]);
    return args.length < 3 ? false : missing(args[2]) ? 0 : ctx.ev(args[2]);
  },
  IFS: (args, ctx) => {
    if (args.length < 2 || args.length % 2) throw new Unsupported('argument count');
    for (let i = 0; i < args.length; i += 2) {
      const c = toBool(sc(ctx, args[i]));
      if (isError(c)) return c;
      if (c) return ctx.ev(args[i + 1]);
    }
    return ERR.na;
  },
  IFERROR: (args, ctx) => {
    arity(args, 2);
    const v = ctx.ev(args[0]);
    const s = isRange(v) || isArray(v) ? ctx.scalar(v) : v;
    return isError(s) ? ctx.ev(args[1]) : v;
  },
  IFNA: (args, ctx) => {
    arity(args, 2);
    const v = ctx.ev(args[0]);
    const s = isRange(v) || isArray(v) ? ctx.scalar(v) : v;
    return isError(s) && s.code === '#N/A' ? ctx.ev(args[1]) : v;
  },
  CHOOSE: (args, ctx) => {
    if (args.length < 2) throw new Unsupported('argument count');
    const i = num(ctx, args[0]);
    if (isError(i)) return i;
    const k = Math.trunc(i);
    return k < 1 || k >= args.length ? ERR.value : ctx.ev(args[k]);
  },
  SWITCH: (args, ctx) => {
    if (args.length < 3) throw new Unsupported('argument count');
    const v = sc(ctx, args[0]);
    if (isError(v)) return v;
    let i = 1;
    for (; i + 1 < args.length; i += 2) {
      const c = sc(ctx, args[i]);
      if (!isError(c) && typeof v === typeof c && same(v, c)) return ctx.ev(args[i + 1]);
    }
    return i < args.length ? ctx.ev(args[i]) : ERR.na;
  },
  AND: (args, ctx) => logic(args, ctx, (xs) => xs.every(Boolean)),
  OR: (args, ctx) => logic(args, ctx, (xs) => xs.some(Boolean)),
  XOR: (args, ctx) => logic(args, ctx, (xs) => xs.filter(Boolean).length % 2 === 1),
  NOT: (args, ctx) => {
    arity(args, 1);
    const b = toBool(sc(ctx, args[0]));
    return isError(b) ? b : !b;
  },
  TRUE: () => true,
  FALSE: () => false,

  // Aggregates
  SUM: agg(sum),
  AVERAGE: agg(mean),
  MIN: agg((xs) => (xs.length ? Math.min(...xs) : 0)),
  MAX: agg((xs) => (xs.length ? Math.max(...xs) : 0)),
  PRODUCT: agg((xs) => xs.reduce((a, b) => a * b, 1)),
  SUMSQ: agg((xs) => sum(xs.map((x) => x * x))),
  MEDIAN: agg((xs) => {
    if (!xs.length) return ERR.num;
    const s = [...xs].sort((a, b) => a - b);
    const m = s.length >> 1;
    return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2;
  }),
  STDEV: agg((xs) => sqrtOr(variance(xs, true))),
  'STDEV.S': agg((xs) => sqrtOr(variance(xs, true))),
  STDEVP: agg((xs) => sqrtOr(variance(xs, false))),
  'STDEV.P': agg((xs) => sqrtOr(variance(xs, false))),
  VAR: agg((xs) => variance(xs, true)),
  'VAR.S': agg((xs) => variance(xs, true)),
  VARP: agg((xs) => variance(xs, false)),
  'VAR.P': agg((xs) => variance(xs, false)),
  COUNT: (args, ctx) => {
    let n = 0;
    for (const a of args) {
      const v = ctx.ev(a);
      if (isRange(v) || isArray(v)) for (const x of ctx.values(v)) n += typeof x === 'number' ? 1 : 0;
      else if (typeof v === 'number' || typeof v === 'boolean' || (typeof v === 'string' && textToNumber(v) !== null)) n++;
    }
    return n;
  },
  COUNTA: (args, ctx) => {
    let n = 0;
    for (const a of args) {
      if (a.t === 'missing') continue;
      const v = ctx.ev(a);
      if (isRange(v) || isArray(v)) for (const x of ctx.values(v)) n += x === null ? 0 : 1;
      else n++;
    }
    return n;
  },
  COUNTBLANK: (args, ctx) => {
    arity(args, 1);
    return ctx.values(ctx.ev(args[0])).filter((x) => x === null || x === '').length;
  },
  LARGE: (args, ctx) => nth(args, ctx, true),
  SMALL: (args, ctx) => nth(args, ctx, false),
  RANK: (args, ctx) => rank(args, ctx),
  'RANK.EQ': (args, ctx) => rank(args, ctx),
  SUMPRODUCT: (args, ctx) => {
    if (!args.length) throw new Unsupported('argument count');
    const arrays = args.map((a) => ctx.array(ctx.ev(a)));
    const { rows, cols } = arrays[0];
    if (arrays.some((a) => a.rows !== rows || a.cols !== cols)) return ERR.value;
    let s = 0;
    for (let i = 0; i < rows * cols; i++) {
      let p = 1;
      for (const a of arrays) {
        const v = a.data[i];
        if (isError(v)) return v;
        p *= typeof v === 'number' ? v : 0;
      }
      s += p;
    }
    return s;
  },
  SUBTOTAL: (args, ctx) => {
    if (args.length < 2) throw new Unsupported('argument count');
    const code = num(ctx, args[0]);
    if (isError(code)) return code;
    const fn = ({ 1: 'AVERAGE', 2: 'COUNT', 3: 'COUNTA', 4: 'MAX', 5: 'MIN', 6: 'PRODUCT', 7: 'STDEV', 8: 'STDEVP', 9: 'SUM', 10: 'VAR', 11: 'VARP' } as Record<number, string>)[code % 100];
    if (!fn) return ERR.value;
    return FUNCTIONS[fn](args.slice(1), ctx);
  },
  SUMIF: ifAggregate('sum'),
  AVERAGEIF: ifAggregate('avg'),
  COUNTIF: (args, ctx) => {
    arity(args, 2);
    return FUNCTIONS.COUNTIFS(args, ctx);
  },
  SUMIFS: ifsAggregate('sum'),
  COUNTIFS: ifsAggregate('count'),
  AVERAGEIFS: ifsAggregate('avg'),
  MAXIFS: ifsAggregate('max'),
  MINIFS: ifsAggregate('min'),

  // Math
  ABS: math1(Math.abs),
  INT: math1(Math.floor),
  SQRT: (args, ctx) => {
    arity(args, 1);
    const x = num(ctx, args[0]);
    return isError(x) ? x : x < 0 ? ERR.num : Math.sqrt(x);
  },
  EXP: math1(Math.exp),
  LN: (args, ctx) => {
    arity(args, 1);
    const x = num(ctx, args[0]);
    return isError(x) ? x : x <= 0 ? ERR.num : Math.log(x);
  },
  LOG10: (args, ctx) => {
    arity(args, 1);
    const x = num(ctx, args[0]);
    return isError(x) ? x : x <= 0 ? ERR.num : Math.log10(x);
  },
  LOG: (args, ctx) => {
    arity(args, 1, 2);
    const x = num(ctx, args[0]);
    const b = num(ctx, args[1], 10);
    const e = firstError(x, b);
    if (e) return e;
    return (x as number) <= 0 || (b as number) <= 0 || b === 1 ? ERR.num : Math.log(x as number) / Math.log(b as number);
  },
  SIGN: math1(Math.sign),
  PI: () => Math.PI,
  SIN: math1(Math.sin),
  COS: math1(Math.cos),
  TAN: math1(Math.tan),
  ASIN: math1(Math.asin),
  ACOS: math1(Math.acos),
  ATAN: math1(Math.atan),
  SINH: math1(Math.sinh),
  COSH: math1(Math.cosh),
  TANH: math1(Math.tanh),
  ATAN2: (args, ctx) => {
    arity(args, 2);
    const x = num(ctx, args[0]);
    const y = num(ctx, args[1]);
    const e = firstError(x, y);
    if (e) return e;
    return x === 0 && y === 0 ? ERR.div0 : Math.atan2(y as number, x as number);
  },
  RADIANS: math1((x) => (x * Math.PI) / 180),
  DEGREES: math1((x) => (x * 180) / Math.PI),
  POWER: (args, ctx) => {
    arity(args, 2);
    const x = num(ctx, args[0]);
    const y = num(ctx, args[1]);
    const e = firstError(x, y);
    if (e) return e;
    const r = (x as number) ** (y as number);
    return Number.isFinite(r) ? r : ERR.num;
  },
  MOD: (args, ctx) => {
    arity(args, 2);
    const x = num(ctx, args[0]);
    const y = num(ctx, args[1]);
    const e = firstError(x, y);
    if (e) return e;
    if (y === 0) return ERR.div0;
    return (x as number) - (y as number) * Math.floor((x as number) / (y as number));
  },
  ROUND: round((x, d) => roundHalfAway(x, d)),
  ROUNDUP: round((x, d) => {
    const f = 10 ** d;
    return (Math.sign(x) * Math.ceil(Number((Math.abs(x) * f).toPrecision(15)))) / f;
  }),
  ROUNDDOWN: round((x, d) => {
    const f = 10 ** d;
    return (Math.sign(x) * Math.floor(Number((Math.abs(x) * f).toPrecision(15)))) / f;
  }),
  TRUNC: (args, ctx) => {
    arity(args, 1, 2);
    const x = num(ctx, args[0]);
    const d = num(ctx, args[1], 0);
    const e = firstError(x, d);
    if (e) return e;
    const f = 10 ** Math.trunc(d as number);
    return Math.trunc((x as number) * f) / f;
  },
  MROUND: (args, ctx) => {
    arity(args, 2);
    const x = num(ctx, args[0]);
    const m = num(ctx, args[1]);
    const e = firstError(x, m);
    if (e) return e;
    if (m === 0) return 0;
    if (Math.sign(x as number) * Math.sign(m as number) < 0) return ERR.num;
    return roundHalfAway((x as number) / (m as number), 0) * (m as number);
  },
  CEILING: (args, ctx) => ceilFloor(args, ctx, true, false),
  'CEILING.MATH': (args, ctx) => ceilFloor(args, ctx, true, true),
  FLOOR: (args, ctx) => ceilFloor(args, ctx, false, false),
  'FLOOR.MATH': (args, ctx) => ceilFloor(args, ctx, false, true),
  EVEN: math1((x) => Math.sign(x) * Math.ceil(Math.abs(x) / 2) * 2),
  ODD: math1((x) => {
    const n = Math.ceil(Math.abs(x));
    return Math.sign(x || 1) * (n % 2 ? n : n + 1);
  }),
  FACT: math1((x) => {
    if (x < 0) return NaN;
    let r = 1;
    for (let i = 2; i <= Math.floor(x); i++) r *= i;
    return r;
  }),

  // Lookup and reference
  VLOOKUP: vhlookup(false),
  HLOOKUP: vhlookup(true),
  MATCH: matchFn,
  INDEX: indexFn,
  XLOOKUP: xlookup,
  LOOKUP: lookupFn,
  OFFSET: offsetFn,
  INDIRECT: indirectFn,
  ROW: (args, ctx) => {
    arity(args, 0, 1);
    const r = refArg(ctx, args[0]);
    return isError(r) ? r : r.r0 + 1;
  },
  COLUMN: (args, ctx) => {
    arity(args, 0, 1);
    const r = refArg(ctx, args[0]);
    return isError(r) ? r : r.c0 + 1;
  },
  ROWS: (args, ctx) => {
    arity(args, 1);
    const v = ctx.ev(args[0]);
    return isError(v) ? v : ctx.dims(v).rows;
  },
  COLUMNS: (args, ctx) => {
    arity(args, 1);
    const v = ctx.ev(args[0]);
    return isError(v) ? v : ctx.dims(v).cols;
  },

  // Arrays and matrices
  TRANSPOSE: (args, ctx) => {
    arity(args, 1);
    const a = ctx.array(ctx.ev(args[0]));
    const data: Scalar[] = new Array(a.rows * a.cols);
    for (let r = 0; r < a.rows; r++) for (let c = 0; c < a.cols; c++) data[c * a.rows + r] = a.data[r * a.cols + c];
    return { kind: 'array', rows: a.cols, cols: a.rows, data };
  },
  MMULT: (args, ctx) => {
    arity(args, 2);
    const a = matrix(ctx, args[0]);
    const b = matrix(ctx, args[1]);
    if (isError(a)) return a;
    if (isError(b)) return b;
    if (a.cols !== b.rows) return ERR.value;
    const data: Scalar[] = [];
    for (let r = 0; r < a.rows; r++) {
      for (let c = 0; c < b.cols; c++) {
        let s = 0;
        for (let k = 0; k < a.cols; k++) s += a.m[r * a.cols + k] * b.m[k * b.cols + c];
        data.push(s);
      }
    }
    return { kind: 'array', rows: a.rows, cols: b.cols, data };
  },
  MDETERM: (args, ctx) => {
    arity(args, 1);
    const a = matrix(ctx, args[0]);
    if (isError(a)) return a;
    if (a.rows !== a.cols) return ERR.value;
    return gauss(a.m, a.rows, false).det;
  },
  MINVERSE: (args, ctx) => {
    arity(args, 1);
    const a = matrix(ctx, args[0]);
    if (isError(a)) return a;
    if (a.rows !== a.cols) return ERR.value;
    const g = gauss(a.m, a.rows, true);
    if (g.det === 0) return ERR.num;
    return { kind: 'array', rows: a.rows, cols: a.cols, data: g.inverse! };
  },
  HEX2DEC: (args, ctx) => baseToDec(args, ctx, 16),
  BIN2DEC: (args, ctx) => baseToDec(args, ctx, 2),
  OCT2DEC: (args, ctx) => baseToDec(args, ctx, 8),

  // Information
  ISBLANK: isType((v) => v === null),
  ISNUMBER: isType((v) => typeof v === 'number'),
  ISTEXT: isType((v) => typeof v === 'string'),
  ISNONTEXT: isType((v) => typeof v !== 'string'),
  ISLOGICAL: isType((v) => typeof v === 'boolean'),
  ISERROR: isType((v) => isError(v)),
  ISERR: isType((v) => isError(v) && v.code !== '#N/A'),
  ISNA: isType((v) => isError(v) && v.code === '#N/A'),
  ISEVEN: (args, ctx) => {
    const x = num(ctx, args[0]);
    return isError(x) ? x : Math.trunc(x) % 2 === 0;
  },
  ISODD: (args, ctx) => {
    const x = num(ctx, args[0]);
    return isError(x) ? x : Math.abs(Math.trunc(x) % 2) === 1;
  },
  NA: () => ERR.na,
  N: (args, ctx) => {
    arity(args, 1);
    const v = sc(ctx, args[0]);
    return typeof v === 'number' ? v : typeof v === 'boolean' ? (v ? 1 : 0) : isError(v) ? v : 0;
  },
  T: (args, ctx) => {
    arity(args, 1);
    const v = sc(ctx, args[0]);
    return typeof v === 'string' ? v : isError(v) ? v : '';
  },

  // Text
  CONCATENATE: (args, ctx) => {
    let s = '';
    for (const a of args) {
      const t = text(ctx, a);
      if (isError(t)) return t;
      s += t;
    }
    return s;
  },
  CONCAT: (args, ctx) => {
    let s = '';
    for (const a of args) {
      const v = ctx.ev(a);
      for (const x of isRange(v) || isArray(v) ? ctx.values(v) : [v]) {
        const t = toText(x);
        if (isError(t)) return t;
        s += t;
      }
    }
    return s;
  },
  TEXTJOIN: (args, ctx) => {
    if (args.length < 3) throw new Unsupported('argument count');
    const delim = text(ctx, args[0]);
    const skip = bool(ctx, args[1], true);
    if (isError(delim)) return delim;
    if (isError(skip)) return skip;
    const parts: string[] = [];
    for (const a of args.slice(2)) {
      const v = ctx.ev(a);
      for (const x of isRange(v) || isArray(v) ? ctx.values(v) : [v]) {
        const t = toText(x);
        if (isError(t)) return t;
        if (skip && t === '') continue;
        parts.push(t);
      }
    }
    return parts.join(delim);
  },
  LEN: passthrough((s) => s.length),
  UPPER: passthrough((s) => s.toUpperCase()),
  LOWER: passthrough((s) => s.toLowerCase()),
  PROPER: passthrough((s) => s.toLowerCase().replace(/(^|[^A-Za-z\u00C0-\uFFFF])([a-z\u00E0-\uFFFF])/g, (_, a: string, b: string) => a + b.toUpperCase())),
  TRIM: passthrough((s) => s.replace(/ +/g, ' ').trim()),
  CLEAN: passthrough((s) => s.replace(/[\x00-\x1f]/g, '')),
  LEFT: (args, ctx) => {
    arity(args, 1, 2);
    const s = text(ctx, args[0]);
    const n = num(ctx, args[1], 1);
    const e = firstError(s, n);
    if (e) return e;
    return (n as number) < 0 ? ERR.value : (s as string).slice(0, Math.trunc(n as number));
  },
  RIGHT: (args, ctx) => {
    arity(args, 1, 2);
    const s = text(ctx, args[0]);
    const n = num(ctx, args[1], 1);
    const e = firstError(s, n);
    if (e) return e;
    const k = Math.trunc(n as number);
    return k < 0 ? ERR.value : k === 0 ? '' : (s as string).slice(-k);
  },
  MID: (args, ctx) => {
    arity(args, 3);
    const s = text(ctx, args[0]);
    const a = num(ctx, args[1]);
    const n = num(ctx, args[2]);
    const e = firstError(s, a, n);
    if (e) return e;
    if ((a as number) < 1 || (n as number) < 0) return ERR.value;
    const st = Math.trunc(a as number) - 1;
    return (s as string).slice(st, st + Math.trunc(n as number));
  },
  REPT: (args, ctx) => {
    arity(args, 2);
    const s = text(ctx, args[0]);
    const n = num(ctx, args[1]);
    const e = firstError(s, n);
    if (e) return e;
    return (n as number) < 0 ? ERR.value : (s as string).repeat(Math.trunc(n as number));
  },
  EXACT: (args, ctx) => {
    arity(args, 2);
    const a = text(ctx, args[0]);
    const b = text(ctx, args[1]);
    const e = firstError(a, b);
    return e ?? a === b;
  },
  SUBSTITUTE: (args, ctx) => {
    arity(args, 3, 4);
    const s = text(ctx, args[0]);
    const from = text(ctx, args[1]);
    const to = text(ctx, args[2]);
    const e = firstError(s, from, to);
    if (e) return e;
    if (from === '') return s;
    if (missing(args[3])) return (s as string).split(from as string).join(to as string);
    const k = num(ctx, args[3]);
    if (isError(k)) return k;
    let seen = 0;
    let idx = -1;
    let pos = 0;
    while ((idx = (s as string).indexOf(from as string, pos)) >= 0) {
      if (++seen === Math.trunc(k)) return (s as string).slice(0, idx) + to + (s as string).slice(idx + (from as string).length);
      pos = idx + (from as string).length;
    }
    return s;
  },
  REPLACE: (args, ctx) => {
    arity(args, 4);
    const s = text(ctx, args[0]);
    const a = num(ctx, args[1]);
    const n = num(ctx, args[2]);
    const t = text(ctx, args[3]);
    const e = firstError(s, a, n, t);
    if (e) return e;
    const st = Math.trunc(a as number) - 1;
    return (s as string).slice(0, st) + t + (s as string).slice(st + Math.trunc(n as number));
  },
  FIND: (args, ctx) => findFn(args, ctx, false),
  SEARCH: (args, ctx) => findFn(args, ctx, true),
  VALUE: (args, ctx) => {
    arity(args, 1);
    const v = sc(ctx, args[0]);
    if (typeof v === 'number') return v;
    if (isError(v)) return v;
    const t = String(v ?? '');
    const n = textToNumber(t);
    if (n !== null) return n;
    const d = dateTextToSerial(t);
    if (d === null) return ERR.value;
    return date1904(ctx) && d >= 1 ? d - 1462 : d;
  },
  TEXT: (args, ctx) => {
    arity(args, 2);
    const v = sc(ctx, args[0]);
    const f = text(ctx, args[1]);
    if (isError(v)) return v;
    if (isError(f)) return f;
    if (typeof v === 'boolean') return formatText(v ? 'TRUE' : 'FALSE', f); // Excel formats a logical as text
    if (typeof v === 'string') {
      const n = textToNumber(v);
      if (n === null) return formatText(v, f);
      return formatValue(n, f, ctx.host.date1904);
    }
    const n = toNumber(v);
    if (isError(n)) return n;
    return formatValue(n, f, ctx.host.date1904);
  },
  CHAR: (args, ctx) => {
    const n = num(ctx, args[0]);
    return isError(n) ? n : n < 1 || n > 255 ? ERR.value : String.fromCharCode(Math.trunc(n));
  },
  CODE: passthrough((s) => (s.length ? s.charCodeAt(0) : ERR.value)),
  FIXED: (args, ctx) => {
    arity(args, 1, 3);
    const x = num(ctx, args[0]);
    const d = num(ctx, args[1], 2);
    const noCommas = bool(ctx, args[2], false);
    const e = firstError(x, d, noCommas);
    if (e) return e;
    const s = roundHalfAway(x as number, d as number).toFixed(Math.max(0, d as number));
    return noCommas ? s : s.replace(/^(-?\d+)/, (m) => m.replace(/\B(?=(\d{3})+(?!\d))/g, ','));
  },

  // Dates
  DATE: (args, ctx) => {
    arity(args, 3);
    const y = num(ctx, args[0]);
    const m = num(ctx, args[1]);
    const d = num(ctx, args[2]);
    const e = firstError(y, m, d);
    if (e) return e;
    let year = Math.trunc(y as number);
    if (year < 1900) year += 1900;
    const s = serialFromParts(year, Math.trunc(m as number), Math.trunc(d as number), date1904(ctx));
    return s < 0 ? ERR.num : s;
  },
  TIME: (args, ctx) => {
    arity(args, 3);
    const h = num(ctx, args[0]);
    const m = num(ctx, args[1]);
    const s = num(ctx, args[2]);
    const e = firstError(h, m, s);
    if (e) return e;
    const t = ((h as number) * 3600 + (m as number) * 60 + (s as number)) / 86400;
    return t - Math.floor(t);
  },
  YEAR: datePart('y'),
  MONTH: datePart('m'),
  DAY: datePart('d'),
  WEEKDAY: datePart('dow'),
  HOUR: math1((x) => Math.floor(Math.round((x - Math.floor(x)) * 86400) / 3600) % 24),
  MINUTE: math1((x) => Math.floor((Math.round((x - Math.floor(x)) * 86400) % 3600) / 60)),
  SECOND: math1((x) => Math.round((x - Math.floor(x)) * 86400) % 60),
  EDATE: edate(false),
  EOMONTH: edate(true),
  DAYS: (args, ctx) => {
    arity(args, 2);
    const a = num(ctx, args[0]);
    const b = num(ctx, args[1]);
    const e = firstError(a, b);
    return e ?? Math.trunc(a as number) - Math.trunc(b as number);
  },
  DATEVALUE: (args, ctx) => {
    arity(args, 1);
    const t = text(ctx, args[0]);
    if (isError(t)) return t;
    const n = dateTextToSerial(t);
    if (n === null || n < 1) return ERR.value; // not a date (a time alone has no date part)
    return Math.floor(n) - (date1904(ctx) ? 1462 : 0);
  },

  // Financial
  PMT: fin(3, 5, ([r, n, pv, fv, t]) => pmt(r, n, pv, fv, t)),
  FV: fin(3, 5, ([r, n, p, pv, t]) => fvOf(r, n, p, pv, t)),
  PV: fin(3, 5, ([r, n, p, fv, t]) => (r === 0 ? -(fv + p * n) : -(fv + (p * (1 + r * t) * ((1 + r) ** n - 1)) / r) / (1 + r) ** n)),
  NPER: fin(3, 5, ([r, p, pv, fv, t]) => {
    if (r === 0) return -(pv + fv) / p;
    const a = p * (1 + r * t);
    return Math.log((a - fv * r) / (a + pv * r)) / Math.log(1 + r);
  }),
  IPMT: fin(4, 6, ([r, per, n, pv, fv, t]) => ipmt(r, per, n, pv, fv, t)),
  PPMT: fin(4, 6, ([r, per, n, pv, fv, t]) => pmt(r, n, pv, fv, t) - ipmt(r, per, n, pv, fv, t)),
  NPV: (args, ctx) => {
    if (args.length < 2) throw new Unsupported('argument count');
    const r = num(ctx, args[0]);
    if (isError(r)) return r;
    const xs = numbers(ctx, args.slice(1), { countText: true });
    if (isError(xs)) return xs;
    let s = 0;
    xs.forEach((x, i) => (s += x / (1 + r) ** (i + 1)));
    return s;
  },
  IRR: (args, ctx) => {
    arity(args, 1, 2);
    const xs = numbers(ctx, [args[0]]);
    if (isError(xs)) return xs;
    const g = num(ctx, args[1], 0.1);
    if (isError(g)) return g;
    return newton((r) => xs.reduce((s, x, i) => s + x / (1 + r) ** i, 0), g);
  },
  RATE: (args, ctx) => {
    arity(args, 3, 6);
    const v: number[] = [];
    for (let i = 0; i < 6; i++) {
      const x = num(ctx, args[i], i === 5 ? 0.1 : 0);
      if (isError(x)) return x;
      v.push(x);
    }
    const [n, p, pv, fv, t, g] = v;
    return newton((r) => (r === 0 ? pv + p * n + fv : pv * (1 + r) ** n + (p * (1 + r * t) * ((1 + r) ** n - 1)) / r + fv), g);
  },
};

function logic(args: Node[], ctx: Ctx, f: (xs: boolean[]) => boolean): Value {
  const xs: boolean[] = [];
  for (const a of args) {
    const v = ctx.ev(a);
    if (isRange(v) || isArray(v)) {
      for (const x of ctx.values(v)) {
        if (isError(x)) return x;
        if (typeof x === 'boolean') xs.push(x);
        else if (typeof x === 'number') xs.push(x !== 0);
      }
    } else {
      const b = toBool(v);
      if (isError(b)) return b;
      xs.push(b);
    }
  }
  return xs.length ? f(xs) : ERR.value;
}

function round(f: (x: number, d: number) => number): Fn {
  return (args, ctx) => {
    arity(args, 1, 2);
    const x = num(ctx, args[0]);
    const d = num(ctx, args[1], 0);
    const e = firstError(x, d);
    if (e) return e;
    return f(x as number, Math.trunc(d as number));
  };
}

/**
 * CEILING / FLOOR and their .MATH forms, with Excel's sign rules:
 * - CEILING(x, s): up for s > 0; for x < 0 and s < 0, away from zero
 *   (CEILING(-2.1, -1) = -3); x > 0 with s < 0 is #NUM!.
 * - FLOOR(x, s): down for s > 0 (FLOOR(-2.1, 1) = -3); for x < 0 and
 *   s < 0, toward zero (-2); x > 0 with s < 0 is #NUM!; s = 0 is #DIV/0!.
 * - .MATH: the sign of s is ignored; negative x rounds toward zero
 *   (CEILING.MATH) or away from it (FLOOR.MATH), reversed by a non-zero mode.
 */
function ceilFloor(args: Node[], ctx: Ctx, up: boolean, math: boolean): Value {
  arity(args, 1, math ? 3 : 2);
  const x = num(ctx, args[0]);
  const s = num(ctx, args[1], 1);
  const mode = math ? num(ctx, args[2], 0) : 0;
  const e = firstError(x, s, mode);
  if (e) return e;
  const xv = x as number;
  const sv = s as number;
  if (sv === 0) return !math && !up ? ERR.div0 : 0;
  if (!math && xv > 0 && sv < 0) return ERR.num;
  const step = Math.abs(sv);
  const q = Number((xv / step).toPrecision(15));
  let toCeil: boolean;
  if (math) toCeil = xv >= 0 ? up : up === (mode === 0);
  else toCeil = sv > 0 ? up : !up; // both negative: CEILING goes away from zero (floor), FLOOR toward it (ceil)
  const r = (toCeil ? Math.ceil(q) : Math.floor(q)) * step;
  return r === 0 ? 0 : r; // no -0
}

function nth(args: Node[], ctx: Ctx, largest: boolean): Value {
  arity(args, 2);
  const xs = numbers(ctx, [args[0]]);
  if (isError(xs)) return xs;
  const k = num(ctx, args[1]);
  if (isError(k)) return k;
  const i = Math.ceil(k) - 1;
  if (i < 0 || i >= xs.length) return ERR.num;
  const s = [...xs].sort((a, b) => (largest ? b - a : a - b));
  return s[i];
}

function rank(args: Node[], ctx: Ctx): Value {
  arity(args, 2, 3);
  const x = num(ctx, args[0]);
  if (isError(x)) return x;
  const xs = numbers(ctx, [args[1]]);
  if (isError(xs)) return xs;
  const asc = num(ctx, args[2], 0);
  if (isError(asc)) return asc;
  if (!xs.includes(x)) return ERR.na;
  return 1 + xs.filter((v) => (asc ? v < x : v > x)).length;
}

function findFn(args: Node[], ctx: Ctx, insensitive: boolean): Value {
  arity(args, 2, 3);
  const needle = text(ctx, args[0]);
  const hay = text(ctx, args[1]);
  const start = num(ctx, args[2], 1);
  const e = firstError(needle, hay, start);
  if (e) return e;
  const from = Math.trunc(start as number) - 1;
  if (from < 0 || from > (hay as string).length) return ERR.value;
  if (insensitive) {
    // SEARCH: case-insensitive, with wildcards.
    const src = wildcard(needle as string).source.slice(1, -1);
    const m = new RegExp(src, 'i').exec((hay as string).slice(from));
    return m ? from + m.index + 1 : ERR.value;
  }
  const i = (hay as string).indexOf(needle as string, from);
  return i < 0 ? ERR.value : i + 1;
}


function matrix(ctx: Ctx, n: Node): { rows: number; cols: number; m: number[] } | XlError {
  const a = ctx.array(ctx.ev(n));
  const m: number[] = [];
  for (const v of a.data) {
    if (isError(v)) return v;
    if (typeof v !== 'number') return ERR.value;
    m.push(v);
  }
  return { rows: a.rows, cols: a.cols, m };
}

/** Gauss-Jordan elimination with partial pivoting: determinant, and optionally the inverse. */
function gauss(src: number[], n: number, invert: boolean): { det: number; inverse?: number[] } {
  const a = [...src];
  const inv: number[] = invert ? Array.from({ length: n * n }, (_, i) => (i % (n + 1) === 0 ? 1 : 0)) : [];
  let det = 1;
  for (let c = 0; c < n; c++) {
    let p = c;
    for (let r = c + 1; r < n; r++) if (Math.abs(a[r * n + c]) > Math.abs(a[p * n + c])) p = r;
    if (a[p * n + c] === 0) return { det: 0 };
    if (p !== c) {
      det = -det;
      for (let k = 0; k < n; k++) {
        [a[p * n + k], a[c * n + k]] = [a[c * n + k], a[p * n + k]];
        if (invert) [inv[p * n + k], inv[c * n + k]] = [inv[c * n + k], inv[p * n + k]];
      }
    }
    const pivot = a[c * n + c];
    det *= pivot;
    for (let k = 0; k < n; k++) {
      a[c * n + k] /= pivot;
      if (invert) inv[c * n + k] /= pivot;
    }
    for (let r = 0; r < n; r++) {
      if (r === c) continue;
      const f = a[r * n + c];
      if (f === 0) continue;
      for (let k = 0; k < n; k++) {
        a[r * n + k] -= f * a[c * n + k];
        if (invert) inv[r * n + k] -= f * inv[c * n + k];
      }
    }
  }
  return { det, inverse: invert ? inv : undefined };
}

function baseToDec(args: Node[], ctx: Ctx, base: number): Value {
  arity(args, 1);
  const t = text(ctx, args[0]);
  if (isError(t)) return t;
  const s = t.trim();
  if (!s || s.length > 10 || !new RegExp(`^[${'0123456789ABCDEF'.slice(0, base)}]+$`, 'i').test(s)) return ERR.num;
  const n = parseInt(s, base);
  // 10-digit values are two's complement (FFFFFFFFFF = -1).
  return s.length === 10 && n >= base ** 10 / 2 ? n - base ** 10 : n;
}
