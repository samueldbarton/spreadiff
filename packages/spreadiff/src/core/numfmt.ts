/**
 * Excel number formats: how a cell with a format code shows its value
 * (`#,##0.0`, `0%`, `$#,##0_);[Red]($#,##0)`, `d-mmm-yy`, `[h]:mm`, `# ?/?`).
 * Used to display cells as the sheet shows them, and by the engine's `TEXT`.
 * Locale is en-US: `.` decimal point, `,` thousands, English month names.
 */
import { partsFromSerial } from './serial.js';

/** Excel's built-in format codes, by `numFmtId` (en-US). */
export const BUILTIN_FORMATS: Readonly<Record<number, string>> = {
  0: 'General',
  1: '0',
  2: '0.00',
  3: '#,##0',
  4: '#,##0.00',
  5: '$#,##0_);($#,##0)',
  6: '$#,##0_);[Red]($#,##0)',
  7: '$#,##0.00_);($#,##0.00)',
  8: '$#,##0.00_);[Red]($#,##0.00)',
  9: '0%',
  10: '0.00%',
  11: '0.00E+00',
  12: '# ?/?',
  13: '# ??/??',
  14: 'm/d/yyyy',
  15: 'd-mmm-yy',
  16: 'd-mmm',
  17: 'mmm-yy',
  18: 'h:mm AM/PM',
  19: 'h:mm:ss AM/PM',
  20: 'h:mm',
  21: 'h:mm:ss',
  22: 'm/d/yyyy h:mm',
  37: '#,##0_);(#,##0)',
  38: '#,##0_);[Red](#,##0)',
  39: '#,##0.00_);(#,##0.00)',
  40: '#,##0.00_);[Red](#,##0.00)',
  41: '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)',
  42: '_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)',
  43: '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)',
  44: '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)',
  45: 'mm:ss',
  46: '[h]:mm:ss',
  47: 'mm:ss.0',
  48: '##0.0E+0',
  49: '@',
};

const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

/** Split a code into sections at `;`, outside quotes, escapes and brackets. */
export function splitSections(code: string): string[] {
  const out: string[] = [];
  let cur = '';
  for (let i = 0; i < code.length; i++) {
    const ch = code[i];
    if (ch === '"') {
      const end = code.indexOf('"', i + 1);
      const j = end < 0 ? code.length : end;
      cur += code.slice(i, j + 1);
      i = j;
    } else if (ch === '\\' || ch === '_' || ch === '*') {
      cur += code.slice(i, i + 2);
      i++;
    } else if (ch === '[') {
      const end = code.indexOf(']', i);
      const j = end < 0 ? code.length : end;
      cur += code.slice(i, j + 1);
      i = j;
    } else if (ch === ';') {
      out.push(cur);
      cur = '';
    } else cur += ch;
  }
  out.push(cur);
  return out;
}

/** The parts of a code outside quotes and escapes (to look for placeholders). */
const bare = (s: string) => s.replace(/"[^"]*"|\\.|_.|\*./g, '');

/** Whether a section formats dates or times. */
function sectionIsDate(s: string): boolean {
  const b = bare(s).replace(/\[(?![hms]+\])[^\]]*\]/gi, '');
  return /[yd]|[hs]|\[[hms]+\]|(?<![#0?.,])m|am\/pm|a\/p/i.test(b) && !/[#0?]/.test(b.replace(/\.0+|0+(?=\])/g, ''));
}

/** Whether a format code shows numbers as dates or times (for its first section). */
export function isDateFormat(code: string): boolean {
  return sectionIsDate(splitSections(code)[0]);
}

interface Section {
  body: string;
  cond?: { op: string; n: number };
  currency: string[];
}

/** A section's brackets: a condition (`[<100]`), a color (dropped), a currency or locale (`[$€-407]` → `€`). */
function parseSection(s: string): Section {
  let cond: Section['cond'];
  const currency: string[] = [];
  const body = s.replace(/"[^"]*"|\\.|\[([^\]]*)\]/g, (t, inner: string | undefined) => {
    if (inner === undefined) return t;
    const c = /^(<=|>=|<>|<|>|=)\s*(-?[\d.]+(?:[eE][-+]?\d+)?)$/.exec(inner);
    if (c) {
      cond = { op: c[1], n: Number(c[2]) };
      return '';
    }
    if (inner[0] === '$') {
      // [$€-407]: the currency text before the dash; [$-409] alone is just a locale.
      const sym = inner.slice(1).split('-')[0];
      if (sym) currency.push(sym);
      return sym ? `"${sym.replace(/"/g, '')}"` : '';
    }
    if (/^[hms]+$/i.test(inner)) return t; // elapsed time, for the date formatter
    return ''; // colors ([Red], [Color10]) and anything else
  });
  return { body, cond, currency };
}

const test = (c: Section['cond'], v: number) => {
  if (!c) return true;
  switch (c.op) {
    case '<': return v < c.n;
    case '<=': return v <= c.n;
    case '>': return v > c.n;
    case '>=': return v >= c.n;
    case '=': return v === c.n;
    default: return v !== c.n;
  }
};

/**
 * Format a number with an Excel format code. `general` formats a `General`
 * section (default: up to 11 significant characters, as a cell shows it).
 * Returns null for a code this doesn't understand, so callers can fall back.
 */
export function formatNumberCode(v: number, code: string, date1904 = false, general: (n: number) => string = generalText): string | null {
  if (!Number.isFinite(v)) return null;
  const raw = splitSections(code);
  const secs = raw.map(parseSection);
  // Pick the section. Sections: positive;negative;zero;text. A negative
  // number in the negative section shows unsigned (the section has its own
  // parentheses or minus); in a lone section it gets a minus sign.
  const numeric = secs.slice(0, 3);
  let sec: Section;
  let value = v;
  let minus = false;
  if (numeric.some((s) => s.cond)) {
    const [a, b, c] = numeric;
    if (test(a.cond ?? { op: '>=', n: 0 }, v)) sec = a;
    else if (b && test(b.cond ?? { op: '<', n: 0 }, v)) sec = b;
    else sec = c ?? b ?? a;
    if (v < 0) {
      if (sec === a || numeric.length === 1) minus = true;
      value = -v;
    }
  } else if (v < 0 && numeric.length >= 2) {
    sec = numeric[1];
    value = -v;
  } else if (v === 0 && numeric.length >= 3) sec = numeric[2];
  else {
    sec = numeric[0];
    if (v < 0) {
      minus = true;
      value = -v;
    }
  }
  const body = sec.body;
  if (body === '' && raw.length > 1) return '';
  if (/^\s*general\s*$/i.test(body)) return (minus ? '-' : '') + general(value);
  if (/general/i.test(bare(body))) return (minus ? '-' : '') + body.replace(/"([^"]*)"|\\(.)|_.|\*.|general/gi, (t, q, e) => (q !== undefined ? q : e !== undefined ? e : t[0] === '_' ? ' ' : t[0] === '*' ? '' : general(value)));
  if (sectionIsDate(body)) return minus ? null : formatDate(value, body, date1904);
  return formatDigits(value, body, minus);
}

/** Format text with a code's text section: the fourth, or a lone section with `@`; null when none applies. */
export function formatTextCode(v: string, code: string): string | null {
  const s = splitSections(code);
  const sec = s.length > 3 ? s[3] : s.length === 1 && /@/.test(bare(s[0])) ? s[0] : null;
  if (sec === null) return null;
  return parseSection(sec).body.replace(/"([^"]*)"|\\(.)|_.|\*.|@/g, (t, q, e) => (q !== undefined ? q : e !== undefined ? e : t[0] === '_' ? ' ' : t[0] === '*' ? '' : v));
}

/** Excel's General in a cell: up to 11 characters, switching to scientific when a number won't fit. */
export function generalText(n: number): string {
  if (n === 0) return '0';
  const a = Math.abs(n);
  if (a >= 1e-9 && a < 1e11) {
    // Up to 10 significant digits, trailing zeros dropped (cells show "0.333333333").
    const int = Math.floor(a).toString().length;
    const decimals = Math.max(0, Math.min(10 - int, 10));
    const s = Number(n.toFixed(decimals)).toString();
    if (!/e/i.test(s) && s.replace(/^-/, '').length <= 11) return s;
  }
  const [m, e] = n.toExponential(5).split('e');
  const mant = m.replace(/\.?0+$/, '');
  const exp = Number(e);
  return `${mant}E${exp < 0 ? '-' : '+'}${String(Math.abs(exp)).padStart(2, '0')}`;
}

function roundHalfAway(x: number, digits: number): number {
  const f = 10 ** digits;
  // Excel rounds the decimal it displays: correct binary noise first (2.675 → 2.68).
  const y = Number((Math.abs(x) * f).toPrecision(15));
  return (Math.sign(x) * Math.round(y)) / f;
}

type Tok = { k: 'lit'; s: string } | { k: 'dig'; c: '0' | '#' | '?' } | { k: 'dot' } | { k: 'comma' } | { k: 'pct' } | { k: 'exp'; sign: '+' | '-' } | { k: 'slash' } | { k: 'at' };

function tokenize(body: string): Tok[] | null {
  const out: Tok[] = [];
  const lit = (s: string) => {
    const last = out[out.length - 1];
    if (last?.k === 'lit') last.s += s;
    else out.push({ k: 'lit', s });
  };
  let seenDot = false;
  for (let i = 0; i < body.length; i++) {
    const ch = body[i];
    if (ch === '"') {
      const end = body.indexOf('"', i + 1);
      const j = end < 0 ? body.length : end;
      lit(body.slice(i + 1, j));
      i = j;
    } else if (ch === '\\') {
      lit(body[i + 1] ?? '');
      i++;
    } else if (ch === '_') {
      lit(' ');
      i++;
    } else if (ch === '*') i++; // fill to width: nothing at a fixed width
    else if (ch === '0' || ch === '#' || ch === '?') out.push({ k: 'dig', c: ch });
    else if (ch === '.' && !seenDot) {
      seenDot = true;
      out.push({ k: 'dot' });
    } else if (ch === ',') out.push({ k: 'comma' });
    else if (ch === '%') out.push({ k: 'pct' });
    else if ((ch === 'E' || ch === 'e') && (body[i + 1] === '+' || body[i + 1] === '-')) {
      out.push({ k: 'exp', sign: body[i + 1] as '+' | '-' });
      i++;
    } else if (ch === '/') out.push({ k: 'slash' });
    else if (ch === '@') out.push({ k: 'at' });
    else if (ch === '[') return null; // leftover elapsed-time brackets in a number section
    else lit(ch);
  }
  return out;
}

/** Fill placeholders right-aligned (integer part): digits beyond them go into the first one. */
function fillInt(digits: string, ph: ('0' | '#' | '?')[]): string[] {
  const out: string[] = new Array(ph.length).fill('');
  let d = digits.length - 1;
  for (let i = ph.length - 1; i >= 0; i--) {
    if (d >= 0) out[i] = digits[d--];
    else out[i] = ph[i] === '0' ? '0' : ph[i] === '?' ? ' ' : '';
  }
  if (d >= 0 && ph.length) out[0] = digits.slice(0, d + 1) + out[0];
  return out;
}

/** Fill placeholders left-aligned (decimals): trailing zeros under `#` vanish, under `?` become spaces. */
function fillDec(digits: string, ph: ('0' | '#' | '?')[]): string[] {
  const out = ph.map((_, i) => digits[i] ?? '0');
  for (let i = ph.length - 1; i >= 0 && out[i] === '0' && ph[i] !== '0'; i--) out[i] = ph[i] === '?' ? ' ' : '';
  return out;
}

function formatDigits(x: number, body: string, minus: boolean): string | null {
  const toks = tokenize(body);
  if (!toks) return null;
  if (toks.some((t) => t.k === 'at')) return null;
  const sign = minus ? '-' : '';
  const digIdx = toks.map((t, i) => (t.k === 'dig' ? i : -1)).filter((i) => i >= 0);
  if (!digIdx.length) {
    // No placeholders: the literal text alone (a percent still scales, but shows nothing).
    return sign + toks.map((t) => (t.k === 'lit' ? t.s : t.k === 'pct' ? '%' : t.k === 'comma' ? ',' : t.k === 'dot' ? '.' : t.k === 'slash' ? '/' : '')).join('');
  }
  const slash = toks.findIndex((t) => t.k === 'slash');
  if (slash >= 0 && toks.slice(0, slash).some((t) => t.k === 'dig') && toks.slice(slash + 1).some((t) => t.k === 'dig' || (t.k === 'lit' && /\d/.test(t.s)))) {
    return formatFraction(x, toks, slash, sign);
  }
  const pct = toks.filter((t) => t.k === 'pct').length;
  x *= 100 ** pct;
  const expAt = toks.findIndex((t) => t.k === 'exp');
  const dotAt = toks.findIndex((t) => t.k === 'dot');
  const mantEnd = expAt >= 0 ? expAt : toks.length;
  const intEnd = dotAt >= 0 && dotAt < mantEnd ? dotAt : mantEnd;
  const firstDig = digIdx[0];
  const lastIntDig = [...digIdx].reverse().find((i) => i < intEnd) ?? -1;
  // Commas between integer placeholders group thousands; commas straight
  // after the mantissa's last placeholder (#,##0,, or 0.0,,"m") divide by 1,000 each.
  let group = false;
  for (let i = firstDig + 1; i < lastIntDig; i++) if (toks[i].k === 'comma') group = true;
  const lastDig = [...digIdx].reverse().find((i) => i < mantEnd) ?? -1;
  let scale = 0;
  for (let i = lastDig + 1; i < mantEnd && (toks[i].k === 'comma' || (toks[i].k === 'dot' && i === intEnd)); i++) if (toks[i].k === 'comma') scale++;
  x /= 1000 ** scale;
  const intPh = toks.slice(0, intEnd).filter((t): t is { k: 'dig'; c: '0' | '#' | '?' } => t.k === 'dig').map((t) => t.c);
  const decPh = toks.slice(intEnd, mantEnd).filter((t): t is { k: 'dig'; c: '0' | '#' | '?' } => t.k === 'dig').map((t) => t.c);
  let exp = 0;
  let expPh: ('0' | '#' | '?')[] = [];
  if (expAt >= 0) {
    expPh = toks.slice(expAt + 1).filter((t): t is { k: 'dig'; c: '0' | '#' | '?' } => t.k === 'dig').map((t) => t.c);
    if (x !== 0) {
      const lg = Math.floor(Math.log10(x));
      // ##0.0E+0: exponents in steps of the integer placeholders (engineering); 0.00E+00: one integer digit.
      const step = intPh.includes('#') && intPh.length > 1 ? intPh.length : 0;
      exp = step ? Math.floor(lg / step) * step : lg - Math.max(0, intPh.filter((c) => c === '0').length - 1);
      x /= 10 ** exp;
      // Rounding can carry the mantissa past its width (9.996 → 10.00): step the exponent up.
      const width = step || Math.max(1, intPh.filter((c) => c === '0').length);
      if (roundHalfAway(x, decPh.length) >= 10 ** width) {
        exp += step || 1;
        x /= 10 ** (step || 1);
      }
    }
  }
  const r = roundHalfAway(x, decPh.length);
  const fixed = r.toFixed(decPh.length);
  const [ip, dp = ''] = fixed.split('.');
  const intDigits = ip === '0' ? '' : ip;
  let ints = fillInt(intDigits, intPh);
  if (group) {
    // Group the whole integer field at once: "1234567" → "1,234,567".
    const field = ints.join('');
    const lead = /^ */.exec(field)![0];
    const grouped = lead + field.slice(lead.length).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    ints = [grouped, ...new Array(intPh.length - 1).fill('')];
  }
  const decs = fillDec(dp, decPh);
  const expDigits = expAt >= 0 ? fillInt(String(Math.abs(exp)), expPh.length ? expPh : ['0']) : [];
  let out = '';
  let ii = 0;
  let di = 0;
  let ei = 0;
  let dotShown = false;
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i];
    if (t.k === 'lit') out += t.s;
    else if (t.k === 'pct') out += '%';
    else if (t.k === 'slash') out += '/';
    else if (t.k === 'comma') continue; // grouping (inserted above) or scaling
    else if (t.k === 'dot') {
      if (!dotShown) out += '.';
      dotShown = true;
    } else if (t.k === 'exp') out += 'E' + (exp < 0 ? '-' : t.sign === '+' ? '+' : '');
    else if (i < intEnd) out += ints[ii++] ?? '';
    else if (i < mantEnd) out += decs[di++] ?? '';
    else out += expDigits[ei++] ?? '';
  }
  // A lone section shows a negative number's sign before the digits, after any leading literal ($-1 shows as -$1).
  return sign + out;
}

function formatFraction(x: number, toks: Tok[], slash: number, sign: string): string {
  // # ?/? : an optional whole-number part, a space, numerator / denominator.
  const before = toks.slice(0, slash);
  const after = toks.slice(slash + 1);
  // The numerator is the last run of placeholders before the slash; anything earlier is the whole part.
  let numStart = before.length;
  while (numStart > 0 && before[numStart - 1].k === 'dig') numStart--;
  const wholeToks = before.slice(0, numStart);
  const numToks = before.slice(numStart);
  const hasWhole = wholeToks.some((t) => t.k === 'dig');
  const fixedDen = after.find((t) => t.k === 'lit' && /^\d+/.test(t.s)) as { k: 'lit'; s: string } | undefined;
  const denPh = after.filter((t) => t.k === 'dig').length;
  const whole = hasWhole ? Math.floor(x) : 0;
  const frac = x - whole;
  let n: number;
  let d: number;
  if (fixedDen) {
    d = Number(/^\d+/.exec(fixedDen.s)![0]);
    n = Math.round(frac * d);
  } else {
    [n, d] = bestFraction(frac, 10 ** Math.max(1, denPh) - 1);
  }
  let w = whole;
  if (n === d && hasWhole) {
    w++;
    n = 0;
  }
  const pad = (s: string, ph: Tok[], left: boolean) => {
    const width = ph.filter((t) => t.k === 'dig').length;
    return left ? s.padStart(width, ' ') : s.padEnd(width, ' ');
  };
  const wholeText = hasWhole ? (w === 0 && n !== 0 ? '' : String(w)) : '';
  const lits = (ts: Tok[]) => ts.filter((t) => t.k === 'lit').map((t) => (t as { s: string }).s).join('');
  if (hasWhole && n === 0) {
    // No fraction: the whole number, and blanks where the fraction would be.
    const width = numToks.length + 1 + (fixedDen ? fixedDen.s.length : denPh);
    return sign + wholeText + lits(wholeToks).replace(/^\s*/, (s) => s) + ' '.repeat(width);
  }
  const num = String(hasWhole ? n : Math.round(n + whole * d));
  const den = fixedDen ? fixedDen.s : pad(String(d), after, false);
  return sign + (wholeText ? wholeText + lits(wholeToks) : hasWhole ? ' '.repeat(wholeText.length) + lits(wholeToks).replace(/./g, ' ') : '') + pad(num, numToks, true) + '/' + den;
}

/** The closest fraction to x (0 ≤ x < 1) with a denominator up to max (Stern–Brocot). */
function bestFraction(x: number, max: number): [number, number] {
  let best: [number, number] = [Math.round(x), 1];
  let err = Math.abs(x - best[0]);
  for (let d = 1; d <= max && err > 1e-12; d++) {
    const n = Math.round(x * d);
    const e = Math.abs(x - n / d);
    if (e < err - 1e-15) {
      best = [n, d];
      err = e;
    }
  }
  return best;
}

function formatDate(v: number, fmt: string, d1904: boolean): string | null {
  // Excel shows ##### for negative dates and ones past 9999-12-31.
  if (v < 0 || v >= (d1904 ? 2957004 : 2958466)) return null;
  // "ss.000" shows fractions of a second: round to that many decimals instead of whole seconds.
  const fd = Math.min(3, Math.max(0, ...[...fmt.matchAll(/s+\.(0+)/gi)].map((m) => m[1].length)));
  const scale = 10 ** fd;
  const totalUnits = Math.round(v * 86400 * scale);
  // Rounding can carry into the next day (23:59:59.6 → next day 00:00:00).
  const dayCarry = Math.floor(totalUnits / (86400 * scale));
  const p = partsFromSerial(dayCarry, d1904);
  const units = totalUnits - dayCarry * 86400 * scale;
  const secs = Math.floor(units / scale);
  const frac = fd ? String(units % scale).padStart(fd, '0') : '';
  const hh = Math.floor(secs / 3600);
  const mm = Math.floor((secs % 3600) / 60);
  const ss = secs % 60;
  const elapsedSecs = Math.floor(totalUnits / scale);
  const ampm = /AM\/PM|A\/P/i.test(bare(fmt));
  // AM/PM and A/P before the literal-run rule, which mustn't swallow them ("ss AM/PM"): "a" is its own literal.
  const tokens = fmt.match(/"[^"]*"|\\.|_.|\*.|\[[hms]+\]|AM\/PM|A\/P|y+|m+|d+|h+|s+|[aA]|[^ymdhsaA"\\_*[]+|\[/gi) ?? [];
  const isUnit = (x: string) => /^(\[?[ymdhs])/i.test(x);
  let out = '';
  for (let k = 0; k < tokens.length; k++) {
    const t = tokens[k];
    const l = t.toLowerCase();
    let prev = '';
    for (let j = k - 1; j >= 0; j--) if (isUnit(tokens[j])) { prev = tokens[j].toLowerCase().replace('[', ''); break; }
    let nextTok = '';
    for (let j = k + 1; j < tokens.length; j++) if (isUnit(tokens[j])) { nextTok = tokens[j].toLowerCase().replace('[', ''); break; }
    if (t.startsWith('"')) out += t.slice(1, -1);
    else if (t.startsWith('\\')) out += t.slice(1);
    else if (t.startsWith('_')) out += ' ';
    else if (t.startsWith('*')) continue;
    else if (/^\[h+\]$/i.test(t)) out += String(Math.floor(elapsedSecs / 3600)).padStart(t.length - 2, '0');
    else if (/^\[m+\]$/i.test(t)) out += String(Math.floor(elapsedSecs / 60)).padStart(t.length - 2, '0');
    else if (/^\[s+\]$/i.test(t)) out += String(elapsedSecs).padStart(t.length - 2, '0');
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
    } else if (l[0] === 'm') {
      // m after h or before s is minutes.
      if (prev.startsWith('h') || nextTok.startsWith('s')) out += l.length === 1 ? mm : String(mm).padStart(2, '0');
      else out += l.length === 1 ? p.m : l.length === 2 ? String(p.m).padStart(2, '0') : l.length === 3 ? MONTHS[p.m - 1].slice(0, 3) : l.length === 5 ? MONTHS[p.m - 1][0] : MONTHS[p.m - 1];
    } else out += t;
  }
  return out;
}
