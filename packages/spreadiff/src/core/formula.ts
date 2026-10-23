/**
 * Formula reference parsing: just enough of Excel's grammar to find which
 * cells a formula reads. No evaluation.
 *
 * Handles A1 refs, ranges, whole columns/rows (`A:C`, `1:3`), absolute
 * markers, sheet qualifiers (`Data!A1`, `'Q&A 2026'!B2:B9`), defined
 * names (via a lookup), and flags what can't be resolved statically:
 * INDIRECT/OFFSET, external workbooks, `#REF!`, structured table refs.
 */

/** A rectangle in 0-based sheet coordinates. `-1` bounds mean "whole column/row". */
export interface SheetRange {
  /** Sheet name from a qualifier (`Data!A1` → `"Data"`), or null for the formula's own sheet. */
  sheet: string | null;
  /** First row, 0-based (A1 is row 0), or -1 for a whole-column range (`C:C`). */
  r0: number;
  /** First column, 0-based (A is 0), or -1 for a whole-row range (`3:3`). */
  c0: number;
  /** Last row, inclusive, or -1 for a whole-column range. */
  r1: number;
  /** Last column, inclusive, or -1 for a whole-row range. */
  c1: number;
}

/** A piece of a formula from {@link tokenizeFormula}: a reference, or anything else as text. */
export type FormulaToken =
  | { type: 'text'; text: string }
  | { type: 'ref'; text: string; range: SheetRange };

/** What a formula reads, from {@link formulaRefs}. */
export interface FormulaRefs {
  /** Every range read, including those reached through defined names. */
  ranges: SheetRange[];
  /** Precedents can't be fully known statically (INDIRECT, OFFSET, unknown names, …). */
  dynamic: boolean;
  /** Uses a volatile function (NOW, TODAY, RAND, …): may change with no input change. */
  volatile: boolean;
}

const DYNAMIC_FUNCS = new Set(['INDIRECT', 'OFFSET']);
const VOLATILE_FUNCS = new Set(['NOW', 'TODAY', 'RAND', 'RANDBETWEEN', 'RANDARRAY']);
const MAX_ROW = 1048575;
const MAX_COL = 16383;

function colIndex(letters: string): number {
  let c = 0;
  for (let i = 0; i < letters.length; i++) c = c * 26 + (letters.charCodeAt(i) - 64);
  return c - 1;
}

const IDENT_CHAR = /[A-Za-z0-9_.\\À-￿]/;

// Tried in order at each position (all sticky).
const STRING = /"(?:[^"]|"")*"/y;
const QUOTED_SHEET = /'((?:[^']|'')+)'!/y;
const BARE_SHEET = /([A-Za-z_À-￿][A-Za-z0-9_.À-￿]*)!/y;
const EXTERNAL = /\[[^\]]*\]/y;
const ERROR = /#(?:REF!|NAME\?|VALUE!|DIV\/0!|N\/A|NUM!|NULL!|SPILL!|CALC!|GETTING_DATA)/y;
const CELL_RANGE = /\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?/y;
const COL_RANGE = /\$?([A-Z]{1,3}):\$?([A-Z]{1,3})/y;
const ROW_RANGE = /\$?(\d+):\$?(\d+)/y;
const IDENT = /[A-Za-z_\\À-￿][A-Za-z0-9_.À-￿]*/y;
const NUMBER = /\d+(?:\.\d*)?(?:[eE][+-]?\d+)?|\.\d+(?:[eE][+-]?\d+)?/y;

function at(re: RegExp, s: string, i: number): RegExpExecArray | null {
  re.lastIndex = i;
  return re.exec(s);
}

/** Split a formula into reference and non-reference tokens. */
export function tokenizeFormula(formula: string): { tokens: FormulaToken[]; dynamic: boolean; volatile: boolean; names: string[] } {
  const f = formula.startsWith('=') ? formula.slice(1) : formula;
  const tokens: FormulaToken[] = [];
  const names: string[] = [];
  let dynamic = false;
  let volatile = false;
  let text = '';
  const flush = () => {
    if (text) tokens.push({ type: 'text', text });
    text = '';
  };
  const boundaryBefore = (i: number) => i === 0 || !IDENT_CHAR.test(f[i - 1]);
  const boundaryAfter = (j: number) => j >= f.length || (!IDENT_CHAR.test(f[j]) && f[j] !== '(' && f[j] !== '!' && f[j] !== '[');

  /** Try a reference (with optional sheet prefix already consumed) at i. */
  const tryRef = (i: number, sheet: string | null): { end: number; range: SheetRange } | null => {
    let m = at(CELL_RANGE, f, i);
    if (m && boundaryAfter(i + m[0].length)) {
      const r0 = Number(m[2]) - 1;
      const c0 = colIndex(m[1]);
      const r1 = m[4] ? Number(m[4]) - 1 : r0;
      const c1 = m[3] ? colIndex(m[3]) : c0;
      if (r0 <= MAX_ROW && r1 <= MAX_ROW && c0 <= MAX_COL && c1 <= MAX_COL) {
        return { end: i + m[0].length, range: { sheet, r0: Math.min(r0, r1), c0: Math.min(c0, c1), r1: Math.max(r0, r1), c1: Math.max(c0, c1) } };
      }
    }
    m = at(COL_RANGE, f, i);
    if (m && boundaryAfter(i + m[0].length)) {
      const a = colIndex(m[1]);
      const b = colIndex(m[2]);
      return { end: i + m[0].length, range: { sheet, r0: -1, r1: -1, c0: Math.min(a, b), c1: Math.max(a, b) } };
    }
    m = at(ROW_RANGE, f, i);
    if (m && boundaryAfter(i + m[0].length)) {
      const a = Number(m[1]) - 1;
      const b = Number(m[2]) - 1;
      return { end: i + m[0].length, range: { sheet, r0: Math.min(a, b), r1: Math.max(a, b), c0: -1, c1: -1 } };
    }
    return null;
  };

  let i = 0;
  while (i < f.length) {
    let m: RegExpExecArray | null;
    if ((m = at(STRING, f, i))) {
      text += m[0];
      i += m[0].length;
      continue;
    }
    if ((m = at(EXTERNAL, f, i))) {
      // [1]Sheet!A1 — another workbook; we can't see it.
      dynamic = true;
      text += m[0];
      i += m[0].length;
      continue;
    }
    if ((m = at(ERROR, f, i))) {
      if (m[0] === '#REF!') dynamic = true;
      text += m[0];
      i += m[0].length;
      continue;
    }
    if (boundaryBefore(i)) {
      // Sheet-qualified reference.
      const q = at(QUOTED_SHEET, f, i) ?? at(BARE_SHEET, f, i);
      if (q) {
        const sheet = q[1].replace(/''/g, "'");
        const ref = tryRef(i + q[0].length, sheet);
        if (ref) {
          flush();
          tokens.push({ type: 'ref', text: f.slice(i, ref.end), range: ref.range });
          i = ref.end;
          continue;
        }
        // Sheet-scoped name (Sheet1!Rate) or 3-D ref: can't resolve.
        dynamic = true;
        text += q[0];
        i += q[0].length;
        continue;
      }
      const ref = tryRef(i, null);
      if (ref) {
        flush();
        tokens.push({ type: 'ref', text: f.slice(i, ref.end), range: ref.range });
        i = ref.end;
        continue;
      }
      if ((m = at(IDENT, f, i))) {
        const word = m[0];
        const end = i + word.length;
        const upper = word.toUpperCase();
        if (f[end] === '(') {
          const name = upper.replace(/^_XLFN\./, '');
          if (DYNAMIC_FUNCS.has(name)) dynamic = true;
          if (VOLATILE_FUNCS.has(name)) volatile = true;
          text += upper;
        } else if (f[end] === '[') {
          dynamic = true; // structured table reference
          text += word;
        } else if (upper === 'TRUE' || upper === 'FALSE') {
          text += upper;
        } else {
          names.push(word);
          flush();
          tokens.push({ type: 'text', text: `\u0000name:${upper}\u0000` });
        }
        i = end;
        continue;
      }
      if ((m = at(NUMBER, f, i))) {
        text += m[0];
        i += m[0].length;
        continue;
      }
    }
    // Whitespace is insignificant outside strings.
    if (f[i] !== ' ' && f[i] !== '\n' && f[i] !== '\r' && f[i] !== '\t') text += f[i];
    i++;
  }
  flush();
  return { tokens, dynamic, volatile, names };
}

/**
 * All ranges a formula reads. `names` resolves defined names (e.g. from
 * `Workbook.names`) to their formula text; unresolved names make the
 * result `dynamic`.
 */
export function formulaRefs(formula: string, names?: Record<string, string>): FormulaRefs {
  return refsOf(formula, names, 0);
}

function refsOf(formula: string, names: Record<string, string> | undefined, depth: number): FormulaRefs {
  const t = tokenizeFormula(formula);
  const ranges: SheetRange[] = [];
  let dynamic = t.dynamic;
  let volatile = t.volatile;
  for (const tok of t.tokens) if (tok.type === 'ref') ranges.push(tok.range);
  for (const n of t.names) {
    const target = lookupName(names, n);
    if (target === undefined || depth > 8) {
      dynamic = true;
      continue;
    }
    const sub = refsOf(target, names, depth + 1);
    ranges.push(...sub.ranges);
    dynamic ||= sub.dynamic;
    volatile ||= sub.volatile;
  }
  return { ranges, dynamic, volatile };
}

export function lookupName(names: Record<string, string> | undefined, name: string): string | undefined {
  if (!names) return undefined;
  if (name in names) return names[name];
  const upper = name.toUpperCase();
  for (const k of Object.keys(names)) if (k.toUpperCase() === upper) return names[k];
  return undefined;
}

/**
 * A formula with every reference replaced by `map(range)`, for comparing
 * formulas by what they point at rather than by their text. Two formulas
 * that differ only because rows were inserted above them canonicalize the
 * same when `map` translates sheet coordinates into aligned diff rows.
 */
export function canonicalFormula(formula: string, map: (range: SheetRange) => string): string {
  const t = tokenizeFormula(formula);
  let out = '';
  for (const tok of t.tokens) out += tok.type === 'ref' ? `\u0001${map(tok.range)}\u0001` : tok.text;
  return out;
}
