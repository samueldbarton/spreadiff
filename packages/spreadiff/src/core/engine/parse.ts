/**
 * Excel formula → syntax tree, for evaluation. Accepts the grammar Excel
 * stores in files: English function names, `,` between arguments, `;`
 * between array-constant rows, `_xlfn.` prefixes.
 */

/** A rectangle in 0-based sheet coordinates; -1 bounds are whole columns/rows. */
export interface RefRange {
  sheet: string | null;
  sheetTo?: string;
  r0: number;
  c0: number;
  r1: number;
  c1: number;
  /**
   * Which bounds are relative (no `$`): bits 1 r0, 2 c0, 4 r1, 8 c1. The
   * numbers are as written for the cell the formula was parsed at; a cached
   * tree reused for another cell shifts the relative ones.
   */
  rel?: number;
}

export type Node =
  | { t: 'num'; v: number }
  | { t: 'str'; v: string }
  | { t: 'bool'; v: boolean }
  | { t: 'err'; v: string }
  | { t: 'ref'; ref: RefRange }
  | { t: 'name'; name: string }
  | { t: 'struct'; table: string; spec: string }
  | { t: 'array'; rows: Node[][] }
  | { t: 'missing' }
  | { t: 'unary'; op: '-' | '+' | '%'; arg: Node }
  | { t: 'binary'; op: string; l: Node; r: Node }
  | { t: 'call'; name: string; args: Node[] };

export type Tok =
  | { k: 'num'; v: number }
  | { k: 'str'; v: string }
  | { k: 'err'; v: string }
  | { k: 'ref'; ref: RefRange }
  | { k: 'ident'; v: string }
  | { k: 'func'; v: string }
  | { k: 'struct'; table: string; spec: string }
  | { k: 'op'; v: string }
  | { k: '('; } | { k: ')' } | { k: ',' } | { k: ';' } | { k: '{' } | { k: '}' } | { k: 'eof' };

/** Thrown for formulas the engine can't read; the caller falls back to static analysis. */
export class Unsupported extends Error {}

// Sticky: matched at an index, without slicing the formula.
const ERROR = /#(?:REF!|NAME\?|VALUE!|DIV\/0!|N\/A|NUM!|NULL!|SPILL!|CALC!|GETTING_DATA)/y;
const CELL = /(\$?)([A-Z]{1,3})(\$?)(\d+)(?::(\$?)([A-Z]{1,3})(\$?)(\d+))?/y;
const COLS = /\$?([A-Z]{1,3}):\$?([A-Z]{1,3})/y;
const ROWS = /\$?(\d+):\$?(\d+)/y;
const QUOTED_SHEET = /'((?:[^']|'')+)'!/y;
const SHEET = /([A-Za-z_\u00C0-\uFFFF][A-Za-z0-9_.\u00C0-\uFFFF]*)(?::([A-Za-z_\u00C0-\uFFFF][A-Za-z0-9_.\u00C0-\uFFFF]*))?!/y;
const IDENT = /[A-Za-z_\\\u00C0-\uFFFF][A-Za-z0-9_.?\\\u00C0-\uFFFF]*/y;
const NUMBER = /(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?/y;
const MAX_ROW = 1048575;
const MAX_COL = 16383;

function at(re: RegExp, s: string, i: number): RegExpExecArray | null {
  re.lastIndex = i;
  return re.exec(s);
}

function col(letters: string): number {
  let c = 0;
  for (let i = 0; i < letters.length; i++) c = c * 26 + (letters.charCodeAt(i) - 64);
  return c - 1;
}

const IDENT_CHAR = /[A-Za-z0-9_.\u00C0-\uFFFF]/;
const identChar = (ch: string | undefined) => ch !== undefined && IDENT_CHAR.test(ch);

/** Split a formula into tokens. */
export function lex(src: string): Tok[] {
  const s = src.charCodeAt(0) === 61 ? src.slice(1) : src;
  const out: Tok[] = [];
  let i = 0;
  const refAt = (j: number, sheet: string | null, sheetTo?: string): { end: number; ref: RefRange } | null => {
    const okEnd = (e: number) => !identChar(s[e]) && s[e] !== '(' && s[e] !== '!';
    let m = at(CELL, s, j);
    if (m && okEnd(j + m[0].length)) {
      const a = { r: Number(m[4]) - 1, c: col(m[2]), rr: !m[3], cr: !m[1] };
      const b = m[8] ? { r: Number(m[8]) - 1, c: col(m[6]), rr: !m[7], cr: !m[5] } : a;
      if (a.r <= MAX_ROW && b.r <= MAX_ROW && a.c <= MAX_COL && b.c <= MAX_COL) {
        // Keep corners ordered; relative flags travel with their bound.
        const [top, bottom] = a.r <= b.r ? [a, b] : [b, a];
        const [left, right] = a.c <= b.c ? [a, b] : [b, a];
        const rel = (top.rr ? 1 : 0) | (left.cr ? 2 : 0) | (bottom.rr ? 4 : 0) | (right.cr ? 8 : 0);
        return { end: j + m[0].length, ref: { sheet, sheetTo, r0: top.r, c0: left.c, r1: bottom.r, c1: right.c, rel } };
      }
    }
    m = at(COLS, s, j);
    if (m && okEnd(j + m[0].length)) {
      const a = col(m[1]);
      const b = col(m[2]);
      return { end: j + m[0].length, ref: { sheet, sheetTo, r0: -1, r1: -1, c0: Math.min(a, b), c1: Math.max(a, b) } };
    }
    m = at(ROWS, s, j);
    if (m && okEnd(j + m[0].length)) {
      const a = Number(m[1]) - 1;
      const b = Number(m[2]) - 1;
      return { end: j + m[0].length, ref: { sheet, sheetTo, r0: Math.min(a, b), r1: Math.max(a, b), c0: -1, c1: -1 } };
    }
    return null;
  };
  while (i < s.length) {
    const ch = s[i];
    if (ch === ' ' || ch === '\n' || ch === '\r' || ch === '\t') {
      i++;
      continue;
    }
    if (ch === '"') {
      let j = i + 1;
      let v = '';
      for (;;) {
        if (j >= s.length) throw new Unsupported('unterminated string');
        if (s[j] === '"') {
          if (s[j + 1] === '"') {
            v += '"';
            j += 2;
            continue;
          }
          break;
        }
        v += s[j++];
      }
      out.push({ k: 'str', v });
      i = j + 1;
      continue;
    }
    if (ch === '#') {
      const m = at(ERROR, s, i);
      if (!m) throw new Unsupported('bad error literal');
      out.push({ k: 'err', v: m[0] });
      i += m[0].length;
      continue;
    }
    if (ch === '[') throw new Unsupported('external or unqualified structured reference');
    if (!(i > 0 && identChar(s[i - 1]))) {
      const q = at(QUOTED_SHEET, s, i) ?? at(SHEET, s, i);
      if (q) {
        let sheet = q[1].replace(/''/g, "'");
        let sheetTo = q[2];
        if (q[0].startsWith("'") && sheet.includes(':')) [sheet, sheetTo] = sheet.split(':');
        const r = refAt(i + q[0].length, sheet, sheetTo);
        if (!r) throw new Unsupported('sheet-scoped name');
        out.push({ k: 'ref', ref: r.ref });
        i = r.end;
        continue;
      }
      const r = refAt(i, null);
      if (r) {
        out.push({ k: 'ref', ref: r.ref });
        i = r.end;
        continue;
      }
      const n = at(NUMBER, s, i);
      if (n) {
        out.push({ k: 'num', v: Number(n[0]) });
        i += n[0].length;
        continue;
      }
      const id = at(IDENT, s, i);
      if (id) {
        const word = id[0];
        const end = i + word.length;
        if (s[end] === '(') {
          out.push({ k: 'func', v: word.toUpperCase().replace(/^_XLFN\.(?:_XLWS\.)?/, '').replace(/^_XLL\./, '') });
          i = end + 1;
          continue;
        }
        if (s[end] === '[') {
          let j = end;
          let depth = 0;
          for (; j < s.length; j++) {
            if (s[j] === "'") j++;
            else if (s[j] === '[') depth++;
            else if (s[j] === ']' && --depth === 0) break;
          }
          if (depth) throw new Unsupported('bad structured reference');
          out.push({ k: 'struct', table: word, spec: s.slice(end, j + 1) });
          i = j + 1;
          continue;
        }
        out.push({ k: 'ident', v: word });
        i = end;
        continue;
      }
    }
    const two = s.slice(i, i + 2);
    if (two === '<=' || two === '>=' || two === '<>') {
      out.push({ k: 'op', v: two });
      i += 2;
      continue;
    }
    if ('+-*/^&=<>%:'.includes(ch)) {
      out.push({ k: 'op', v: ch });
      i++;
      continue;
    }
    if (ch === '(' || ch === ')' || ch === ',' || ch === ';' || ch === '{' || ch === '}') {
      out.push({ k: ch } as Tok);
      i++;
      continue;
    }
    throw new Unsupported(`unexpected ${ch}`);
  }
  out.push({ k: 'eof' });
  return out;
}

/**
 * A key for "the same formula, relative to its cell": relative references
 * become offsets from (row, col). Formulas copied down a column share it.
 */
export function relativeKey(toks: Tok[], row: number, col: number): string {
  let k = '';
  for (const t of toks) {
    switch (t.k) {
      case 'ref': {
        const r = t.ref;
        const rel = r.rel ?? 0;
        const f = (v: number, bit: number, base: number) => (rel & bit ? `~${v - base}` : `${v}`);
        k += `\u0001${r.sheet ?? ''}:${r.sheetTo ?? ''}:${f(r.r0, 1, row)},${f(r.c0, 2, col)},${f(r.r1, 4, row)},${f(r.c1, 8, col)}`;
        break;
      }
      case 'num':
        k += `\u0002${t.v}`;
        break;
      case 'str':
        k += `\u0003${t.v.replace(/\u0003/g, '')}\u0003`;
        break;
      case 'err':
      case 'ident':
      case 'func':
      case 'op':
        k += `\u0004${t.k}${t.v}`;
        break;
      case 'struct':
        k += `\u0005${t.table}${t.spec}`;
        break;
      default:
        k += t.k;
    }
  }
  return k;
}

const BINARY: Record<string, number> = { '=': 1, '<>': 1, '<': 1, '>': 1, '<=': 1, '>=': 1, '&': 2, '+': 3, '-': 3, '*': 4, '/': 4, '^': 5 };

/** Parse a formula (with or without the leading `=`). Throws {@link Unsupported}. */
export function parseFormula(formula: string): Node {
  return parseTokens(lex(formula));
}

/** Parse tokens from {@link lex}. Throws {@link Unsupported}. */
export function parseTokens(toks: Tok[]): Node {
  let p = 0;
  const peek = () => toks[p];
  const next = () => toks[p++];
  const expect = (k: Tok['k']) => {
    const t = next();
    if (t.k !== k) throw new Unsupported(`expected ${k}`);
    return t;
  };

  const primary = (): Node => {
    const t = next();
    switch (t.k) {
      case 'num':
        return { t: 'num', v: t.v };
      case 'str':
        return { t: 'str', v: t.v };
      case 'err':
        return { t: 'err', v: t.v };
      case 'ref':
        return { t: 'ref', ref: t.ref };
      case 'struct':
        return { t: 'struct', table: t.table, spec: t.spec };
      case 'ident': {
        const u = t.v.toUpperCase();
        if (u === 'TRUE' || u === 'FALSE') return { t: 'bool', v: u === 'TRUE' };
        return { t: 'name', name: t.v };
      }
      case 'func': {
        const args: Node[] = [];
        if (peek().k === ')') {
          next();
          return { t: 'call', name: t.v, args };
        }
        for (;;) {
          if (peek().k === ',' || peek().k === ')') args.push({ t: 'missing' });
          else args.push(expr(0));
          const sep = next();
          if (sep.k === ')') break;
          if (sep.k !== ',') throw new Unsupported('bad argument list');
        }
        return { t: 'call', name: t.v, args };
      }
      case '(': {
        const e = expr(0);
        expect(')');
        return e;
      }
      case '{': {
        const rows: Node[][] = [[]];
        for (;;) {
          rows[rows.length - 1].push(unary());
          const sep = next();
          if (sep.k === '}') break;
          if (sep.k === ';') rows.push([]);
          else if (sep.k !== ',') throw new Unsupported('bad array constant');
        }
        return { t: 'array', rows };
      }
      case 'op':
        if (t.v === '-' || t.v === '+') return { t: 'unary', op: t.v, arg: unary() };
        throw new Unsupported(`unexpected ${t.v}`);
      default:
        throw new Unsupported(`unexpected ${t.k}`);
    }
  };

  // Range operator (A1:B2 written with a name or function on one side) binds tightest.
  const rangeOp = (): Node => {
    let l = primary();
    while (peek().k === 'op' && (peek() as { v: string }).v === ':') {
      next();
      const r = primary();
      l = { t: 'binary', op: ':', l, r };
    }
    return l;
  };
  const postfix = (): Node => {
    let n = rangeOp();
    while (peek().k === 'op' && (peek() as { v: string }).v === '%') {
      next();
      n = { t: 'unary', op: '%', arg: n };
    }
    return n;
  };
  // Excel: negation binds tighter than ^ (so -2^2 = 4).
  const unary = (): Node => {
    const t = peek();
    if (t.k === 'op' && (t.v === '-' || t.v === '+')) {
      next();
      return { t: 'unary', op: t.v, arg: unary() };
    }
    return postfix();
  };
  const expr = (min: number): Node => {
    let l = unary();
    for (;;) {
      const t = peek();
      if (t.k !== 'op') break;
      const prec = BINARY[t.v];
      if (prec === undefined || prec < min) break;
      next();
      // Left-associative, ^ included: Excel evaluates 2^3^2 as 64.
      const r = expr(prec + 1);
      l = { t: 'binary', op: t.v, l, r };
    }
    return l;
  };
  const tree = expr(0);
  if (peek().k !== 'eof') throw new Unsupported('trailing input');
  return tree;
}
