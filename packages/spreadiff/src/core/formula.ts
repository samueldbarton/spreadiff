/**
 * Formula reference parsing: just enough of Excel's grammar to find which
 * cells a formula reads. No evaluation.
 *
 * Handles A1 refs, ranges, whole columns/rows (`A:C`, `1:3`), absolute
 * markers, sheet qualifiers (`Data!A1`, `'Q&A 2026'!B2:B9`), defined
 * names (via a lookup), and flags what can't be resolved statically:
 * INDIRECT/OFFSET, external workbooks and `#REF!`. Structured references
 * (`Sales[Amount]`) resolve when the workbook's tables are passed in.
 */
import type { TableDefinition } from './types.js';

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
  /** For a 3-D reference (`Jan:Mar!B2`), the last sheet; `sheet` is the first. */
  sheetTo?: string;
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

// TABLE: a what-if data table (its results come from re-running another formula).
const DYNAMIC_FUNCS = new Set(['INDIRECT', 'OFFSET', 'TABLE']);
// CELL and INFO read the environment (column widths, OS, recalc mode), so they behave like volatiles.
const VOLATILE_FUNCS = new Set(['NOW', 'TODAY', 'RAND', 'RANDBETWEEN', 'RANDARRAY', 'CELL', 'INFO']);
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
const BARE_3D = /([A-Za-z_À-￿][A-Za-z0-9_.À-￿]*):([A-Za-z_À-￿][A-Za-z0-9_.À-￿]*)!/y;
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
export function tokenizeFormula(formula: string): {
  tokens: FormulaToken[];
  dynamic: boolean;
  volatile: boolean;
  names: string[];
  /** Structured references (`Sales[[#This Row],[Amount]]`): table name and the bracketed specifier. */
  structured: { table: string; spec: string }[];
} {
  const f = formula.startsWith('=') ? formula.slice(1) : formula;
  const tokens: FormulaToken[] = [];
  const names: string[] = [];
  const structured: { table: string; spec: string }[] = [];
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
      // Sheet-qualified reference, possibly 3-D (Jan:Mar!B2; sheet names can't contain ':').
      const q3 = at(BARE_3D, f, i);
      const q = q3 ?? at(QUOTED_SHEET, f, i) ?? at(BARE_SHEET, f, i);
      if (q) {
        const spec = q3 ? `${q3[1]}:${q3[2]}` : q[1].replace(/''/g, "'");
        const colon = spec.indexOf(':');
        const sheet = colon >= 0 ? spec.slice(0, colon) : spec;
        const ref = tryRef(i + q[0].length, sheet);
        if (ref && colon >= 0) ref.range.sheetTo = spec.slice(colon + 1);
        if (ref) {
          flush();
          tokens.push({ type: 'ref', text: f.slice(i, ref.end), range: ref.range });
          i = ref.end;
          continue;
        }
        // Sheet-scoped name from another sheet (Sheet1!Rate), resolved as "Sheet1!Rate".
        const nm = at(IDENT, f, i + q[0].length);
        if (nm && f[i + q[0].length + nm[0].length] !== '(' && colon < 0) {
          const qualified = `${sheet}!${nm[0]}`;
          names.push(qualified);
          flush();
          tokens.push({ type: 'text', text: `\u0000name:${qualified.toUpperCase()}\u0000` });
          i += q[0].length + nm[0].length;
          continue;
        }
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
        if ((upper === 'TRUE' || upper === 'FALSE') && f[end] === '(' && f[end + 1] === ')') {
          // TRUE() and FALSE() are the constants (LibreOffice writes them that way).
          text += upper;
          i = end + 2;
          continue;
        }
        if (f[end] === '(') {
          const name = upper.replace(/^_XLFN\./, '');
          if (DYNAMIC_FUNCS.has(name)) dynamic = true;
          if (VOLATILE_FUNCS.has(name)) volatile = true;
          text += upper;
        } else if (f[end] === '[') {
          // Structured reference: Sales[Amount], Sales[[#This Row],[Amount]].
          // Brackets nest; an apostrophe escapes the next character.
          let j = end;
          let depth = 0;
          for (; j < f.length; j++) {
            const ch = f[j];
            if (ch === "'") j++;
            else if (ch === '[') depth++;
            else if (ch === ']' && --depth === 0) break;
          }
          if (depth !== 0) {
            dynamic = true;
            text += word;
          } else {
            structured.push({ table: word, spec: f.slice(end, j + 1) });
            text += f.slice(i, j + 1);
            i = j + 1;
            continue;
          }
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
  return { tokens, dynamic, volatile, names, structured };
}

/** Where a formula sits, so structured references like `[#This Row]` can resolve. */
export interface FormulaContext {
  /** The workbook's Excel tables. */
  tables?: TableDefinition[];
  /** 0-based sheet row of the formula's cell. */
  row?: number;
  /** Sheet names in workbook order, to expand 3-D references (`Jan:Mar!B2`). */
  sheets?: string[];
}

/**
 * All ranges a formula reads. `names` resolves defined names (e.g. from
 * `Workbook.names`) to their formula text; `context.tables` resolves
 * structured references. Anything unresolved makes the result `dynamic`.
 */
export function formulaRefs(formula: string, names?: Record<string, string>, context?: FormulaContext): FormulaRefs {
  return refsOf(formula, names, context, 0);
}

function refsOf(formula: string, names: Record<string, string> | undefined, context: FormulaContext | undefined, depth: number): FormulaRefs {
  const t = tokenizeFormula(formula);
  const ranges: SheetRange[] = [];
  let dynamic = t.dynamic;
  let volatile = t.volatile;
  for (const tok of t.tokens) {
    if (tok.type !== 'ref') continue;
    const rg = tok.range;
    if (rg.sheetTo === undefined) {
      ranges.push(rg);
      continue;
    }
    // 3-D: the same range on every sheet from `sheet` to `sheetTo`.
    const order = context?.sheets?.map((x) => x.toUpperCase()) ?? [];
    const a = order.indexOf(rg.sheet!.toUpperCase());
    const b = order.indexOf(rg.sheetTo.toUpperCase());
    if (a < 0 || b < 0) {
      dynamic = true;
      continue;
    }
    for (let k = Math.min(a, b); k <= Math.max(a, b); k++) ranges.push({ sheet: context!.sheets![k], r0: rg.r0, c0: rg.c0, r1: rg.r1, c1: rg.c1 });
  }
  for (const n of t.names) {
    const target = lookupName(names, n);
    if (target === undefined || depth > 8) {
      dynamic = true;
      continue;
    }
    const sub = refsOf(target, names, context, depth + 1);
    ranges.push(...sub.ranges);
    dynamic ||= sub.dynamic;
    volatile ||= sub.volatile;
  }
  for (const s of t.structured) {
    const r = structuredRange(s.table, s.spec, context);
    if (r) ranges.push(r);
    else dynamic = true;
  }
  return { ranges, dynamic, volatile };
}

/**
 * Resolve a structured reference (`Sales[Amount]`, `Sales[[#Totals],[Amount]]`,
 * `Sales[[#This Row],[Qty]:[Price]]`, `Sales[#All]`) to a sheet range, or
 * null when the table or column isn't known.
 */
export function structuredRange(table: string, spec: string, context?: FormulaContext): SheetRange | null {
  const def = context?.tables?.find((t) => t.name.toUpperCase() === table.toUpperCase());
  if (!def) return null;
  const inner = spec.slice(1, -1).trim();
  // Items are either one bare specifier (Sales[Amount]) or a list of [..] items.
  const items: string[] = [];
  if (inner.startsWith('[')) {
    let depth = 0;
    let cur = '';
    for (let i = 0; i < inner.length; i++) {
      const ch = inner[i];
      if (ch === "'") {
        cur += inner[++i] ?? '';
        continue;
      }
      if (ch === '[') {
        if (depth++ === 0) continue;
      } else if (ch === ']') {
        if (--depth === 0) {
          items.push(cur); // spaces inside brackets are part of the name ("Field 4 ")
          cur = '';
          continue;
        }
      } else if (depth === 0) {
        if (ch === ':') items.push(':');
        continue; // commas and spaces between items
      }
      cur += ch;
    }
  } else if (inner.startsWith('@')) {
    items.push('@');
    if (inner.length > 1) items.push(inner.slice(1).replace(/'(.)/g, '$1'));
  } else if (inner) {
    items.push(inner.replace(/'(.)/g, '$1'));
  }
  const { range, headerRows, totalsRows } = def;
  const dataTop = range.r0 + headerRows;
  const dataBottom = range.r1 - totalsRows;
  let r0 = dataTop;
  let r1 = dataBottom;
  const special = items.filter((x) => x.startsWith('#') || x === '@').map((x) => x.toUpperCase());
  if (special.length) {
    const has = (k: string) => special.includes(k);
    if (has('#ALL')) {
      r0 = range.r0;
      r1 = range.r1;
    } else if (has('#THIS ROW') || has('@')) {
      if (context?.row === undefined) return null;
      r0 = r1 = context.row;
    } else {
      r0 = has('#HEADERS') ? range.r0 : has('#DATA') ? dataTop : dataBottom + 1;
      r1 = has('#TOTALS') ? range.r1 : has('#DATA') ? dataBottom : range.r0 + headerRows - 1;
    }
  }
  const cols = items.filter((x) => !x.startsWith('#') && x !== '@');
  let c0 = range.c0;
  let c1 = range.c1;
  if (cols.length) {
    const idx = (name: string) => def.columns.findIndex((c) => c.toUpperCase() === name.replace(/^@/, '').toUpperCase());
    const names = cols.filter((x) => x !== ':');
    const a = idx(names[0]);
    const b = idx(names[names.length - 1]);
    if (a < 0 || b < 0) return null;
    c0 = range.c0 + Math.min(a, b);
    c1 = range.c0 + Math.max(a, b);
  }
  if (r0 > r1) return null;
  return { sheet: def.sheet, r0, r1, c0, c1 };
}

/**
 * The defined names a formula on `sheet` sees: the workbook's, then that
 * sheet's own (`Table.names`, which win), plus every sheet's own as
 * `Sheet!Name`. Pass the result as `names` to {@link formulaRefs}.
 */
export function scopedNames(
  workbookNames: Record<string, string> | undefined,
  sheets: readonly { name: string; names?: Record<string, string> }[],
  sheet: string,
): Record<string, string> | undefined {
  if (!sheets.some((s) => s.names)) return workbookNames;
  const out: Record<string, string> = { ...workbookNames };
  for (const s of sheets) {
    if (!s.names) continue;
    for (const [k, v] of Object.entries(s.names)) out[`${s.name}!${k}`] = qualify(v, s.name);
  }
  const own = sheets.find((s) => s.name === sheet)?.names;
  if (own) for (const [k, v] of Object.entries(own)) out[k] = qualify(v, sheet);
  return out;
}

/** A sheet-scoped name's unqualified references ($B$2) point at its own sheet. */
function qualify(formula: string, sheet: string): string {
  const prefix = /^[A-Za-z_][\w.]*$/.test(sheet) ? sheet : `'${sheet.replace(/'/g, "''")}'`;
  return formula.replace(/("[^"]*")|((?:'(?:[^']|'')+'|[A-Za-z_][\w.]*)!)?(\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)(?![\w(!])/g, (all, str, q, ref) => (str || q ? all : `${prefix}!${ref}`));
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
  for (const tok of t.tokens) {
    if (tok.type !== 'ref') out += tok.text;
    else if (tok.range.sheetTo === undefined) out += `\u0001${map(tok.range)}\u0001`;
    else {
      // 3-D references don't map through one table: compare what they cover (A1 ≡ A1:A1).
      const r = tok.range;
      out += `\u00013d:${r.sheet!.toUpperCase()}:${r.sheetTo!.toUpperCase()}:${r.r0},${r.c0},${r.r1},${r.c1}\u0001`;
    }
  }
  return out;
}
