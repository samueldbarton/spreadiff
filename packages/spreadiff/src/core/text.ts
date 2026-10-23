import { alignSequences } from './sequence.js';
import type { CellValue } from './types.js';

/** A run of text in a {@link diffText} result. */
export interface TextSpan {
  /** The characters of this run. */
  text: string;
  /** `equal` in both strings, `removed` only in the old one, `added` only in the new one. */
  kind: 'equal' | 'added' | 'removed';
}

/**
 * Character-level diff of two cell values, for inline highlighting. Works on
 * words for long strings so the result stays readable.
 */
export function diffText(before: string, after: string): TextSpan[] {
  const tokenize = (s: string) =>
    s.length > 40 ? s.match(/\s+|[\p{L}\p{N}_]+|[^\s\p{L}\p{N}_]/gu) ?? [] : Array.from(s);
  const a = tokenize(before);
  const b = tokenize(after);
  const ids = new Map<string, number>();
  const id = (t: string) => {
    let v = ids.get(t);
    if (v === undefined) ids.set(t, (v = ids.size));
    return v;
  };
  const matches = alignSequences(Int32Array.from(a, id), Int32Array.from(b, id));
  const out: TextSpan[] = [];
  const push = (text: string, kind: TextSpan['kind']) => {
    if (!text) return;
    const last = out[out.length - 1];
    if (last && last.kind === kind) last.text += text;
    else out.push({ text, kind });
  };
  let pa = 0;
  let pb = 0;
  for (let k = 0; k <= matches.length; k += 2) {
    const ia = k < matches.length ? matches[k] : a.length;
    const ib = k < matches.length ? matches[k + 1] : b.length;
    push(a.slice(pa, ia).join(''), 'removed');
    push(b.slice(pb, ib).join(''), 'added');
    if (k < matches.length) push(a[ia], 'equal');
    pa = ia + 1;
    pb = ib + 1;
  }
  return out;
}

/** Numeric change between two cells, or null when either side is not a number. */
export function numericDelta(before: CellValue, after: CellValue): { delta: number; ratio: number | null } | null {
  const x = typeof before === 'number' ? before : typeof before === 'string' && before.trim() ? Number(before) : NaN;
  const y = typeof after === 'number' ? after : typeof after === 'string' && after.trim() ? Number(after) : NaN;
  if (!Number.isFinite(x) || !Number.isFinite(y)) return null;
  // A difference is only as precise as its operands: keep 15 significant
  // digits of the larger one (3878764.11 − 3887856.67 = −9092.56, not −9092.56000000006).
  const scale = Math.max(Math.abs(x), Math.abs(y));
  const decimals = scale ? 14 - Math.floor(Math.log10(scale)) : 0;
  const raw = y - x;
  const delta = decimals > 0 && decimals <= 100 ? Number(raw.toFixed(decimals)) : raw;
  return { delta, ratio: x === 0 ? null : delta / Math.abs(x) };
}

/**
 * Display text for a cell value: `''` for null, `TRUE`/`FALSE` for booleans,
 * and numbers with float noise trimmed (`0.1 + 0.2` → `"0.3"`) at 15
 * significant digits.
 */
export function formatCell(v: CellValue): string {
  if (v === null || v === undefined) return '';
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v === 'number') {
    if (Number.isInteger(v)) return String(v);
    // Trim float noise (0.1 + 0.2 → 0.3); 15 significant digits is the most
    // a double always round-trips, so real differences stay visible.
    return String(Number(v.toPrecision(15)));
  }
  return v;
}
