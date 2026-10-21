import type { CellValue, Table } from './types.js';

export interface ParseCsvOptions {
  /** Field delimiter. Detected from the first lines when omitted. */
  delimiter?: string;
  /** Treat the first row as column names. Default true. */
  header?: boolean;
  /**
   * Convert numeric fields to numbers. `'lossless'` (default) converts only
   * when the number prints back identically, so "007", "1.50" and "1e5"
   * stay text. Numbers make diffs, rendering and worker transfer much cheaper.
   */
  numbers?: boolean | 'lossless';
  name?: string;
}

/** Spreadsheet-style column label: 0 → A, 25 → Z, 26 → AA. */
export function columnLabel(index: number): string {
  let s = '';
  let n = index + 1;
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

export function detectDelimiter(text: string): string {
  const sample = text.slice(0, 64 * 1024);
  const lines = sample.split(/\r?\n/, 20).filter((l) => l.length > 0);
  let best = ',';
  let bestScore = -1;
  for (const d of [',', '\t', ';', '|']) {
    const counts = lines.map((l) => l.split(d).length - 1);
    if (counts.length === 0 || counts[0] === 0) continue;
    // Prefer delimiters that appear a consistent number of times per line.
    const consistent = counts.filter((c) => c === counts[0]).length;
    const score = consistent * 1000 + counts[0];
    if (score > bestScore) {
      bestScore = score;
      best = d;
    }
  }
  return best;
}

/** RFC 4180 parser. Single pass over char codes; no regex per field. */
export function parseCsv(text: string, options: ParseCsvOptions = {}): Table {
  if (text.charCodeAt(0) === 0xfeff) text = text.slice(1);
  const delim = (options.delimiter ?? detectDelimiter(text)).charCodeAt(0);
  const QUOTE = 34;
  const CR = 13;
  const LF = 10;
  const rows: string[][] = [];
  let row: string[] = [];
  let i = 0;
  const len = text.length;
  while (i < len) {
    let field: string;
    if (text.charCodeAt(i) === QUOTE) {
      i++;
      let start = i;
      let parts: string[] | null = null;
      for (;;) {
        const q = text.indexOf('"', i);
        if (q === -1) {
          i = len;
          field = (parts ? parts.join('') : '') + text.slice(start);
          break;
        }
        if (text.charCodeAt(q + 1) === QUOTE) {
          (parts ??= []).push(text.slice(start, q + 1));
          i = q + 2;
          start = i;
          continue;
        }
        field = (parts ? parts.join('') : '') + text.slice(start, q);
        i = q + 1;
        break;
      }
      // Skip anything between the closing quote and the next delimiter.
      while (i < len) {
        const c = text.charCodeAt(i);
        if (c === delim || c === LF || c === CR) break;
        i++;
      }
    } else {
      const start = i;
      while (i < len) {
        const c = text.charCodeAt(i);
        if (c === delim || c === LF || c === CR) break;
        i++;
      }
      field = text.slice(start, i);
    }
    row.push(field);
    const c = text.charCodeAt(i);
    if (c === delim) {
      i++;
      if (i === len) row.push('');
    } else {
      if (c === CR) i++;
      if (text.charCodeAt(i) === LF) i++;
      rows.push(row);
      row = [];
    }
  }
  if (row.length) rows.push(row);
  return toTable(rows, options);
}

const NUMERIC = /^[-+]?(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?$/;

/**
 * True when `f` is a decimal that JavaScript prints back verbatim
 * (`String(Number(f)) === f`) without paying for that round trip:
 * `-?(0|[1-9]\d*)(\.\d*[1-9])?`, at most 15 significant digits (exact in a
 * double) and not so small that it prints in exponent form (< 1e-6).
 */
export function isCanonicalNumber(f: string): boolean {
  const len = f.length;
  const neg = f.charCodeAt(0) === 45; // '-'
  let i = neg ? 1 : 0;
  if (i >= len) return false;
  let c = f.charCodeAt(i);
  if (c < 48 || c > 57) return false;
  let digits = 0;
  if (c === 48) i++;
  else {
    while (i < len && (c = f.charCodeAt(i)) >= 48 && c <= 57) {
      i++;
      digits++;
    }
  }
  if (i === len) return digits <= 15 && !(neg && digits === 0); // "-0" prints "0"
  if (f.charCodeAt(i) !== 46 || ++i === len) return false; // need ".d"
  const intZero = digits === 0;
  let leadingZeros = 0;
  let last = 0;
  for (; i < len; i++) {
    c = f.charCodeAt(i);
    if (c < 48 || c > 57) return false;
    if (intZero && digits === 0 && c === 48) leadingZeros++;
    else digits++;
    last = c;
  }
  return last !== 48 && digits <= 15 && leadingZeros < 6;
}

function toTable(raw: string[][], options: ParseCsvOptions): Table {
  // Drop trailing fully-empty rows (common at EOF).
  while (raw.length && raw[raw.length - 1].every((f) => f === '')) raw.pop();
  const width = raw.reduce((w, r) => Math.max(w, r.length), 0);
  const header = options.header ?? true;
  const columns = header
    ? uniqueNames(Array.from({ length: width }, (_, c) => raw[0]?.[c] ?? ''))
    : Array.from({ length: width }, (_, c) => columnLabel(c));
  const body = header ? raw.slice(1) : raw;
  const numbers = options.numbers ?? 'lossless';
  const rows: CellValue[][] = new Array(body.length);
  for (let r = 0; r < body.length; r++) {
    const src = body[r];
    const out: CellValue[] = new Array(width);
    for (let c = 0; c < width; c++) {
      const f = src[c];
      if (f === undefined || f === '') out[c] = null;
      else if (numbers === 'lossless') out[c] = isCanonicalNumber(f) ? Number(f) : f;
      else if (numbers && NUMERIC.test(f)) out[c] = Number(f);
      else out[c] = f;
    }
    rows[r] = out;
  }
  return { name: options.name ?? 'Sheet1', columns, rows, origin: { row: header ? 1 : 0, col: 0 } };
}

/** Fill blanks with column letters and suffix duplicates: `x`, `x (2)`. */
export function uniqueNames(names: string[]): string[] {
  const bases = names.map((raw, i) => (raw.trim() === '' ? columnLabel(i) : raw.trim()));
  // Reserve every original name first, so a generated "x (2)" never
  // collides with a real column called "x (2)".
  const taken = new Set<string>();
  const firstSeen = new Set<string>();
  const isDup = bases.map((b) => (firstSeen.has(b) ? true : (firstSeen.add(b), false)));
  bases.forEach((b, i) => { if (!isDup[i]) taken.add(b); });
  return bases.map((b, i) => {
    if (!isDup[i]) return b;
    let n = 2;
    while (taken.has(`${b} (${n})`)) n++;
    const name = `${b} (${n})`;
    taken.add(name);
    return name;
  });
}
