/**
 * Sequence alignment over interned integer ids.
 *
 * Myers' O((N+M)·D) algorithm with linear-space bisection does the work when
 * the edit distance is small, which is the common case for spreadsheets. When
 * the budget runs out (huge rewrites, sorted vs. unsorted data) we fall back
 * to patience anchoring: elements unique in both sides, aligned by longest
 * increasing subsequence. That keeps the worst case at O(N log N) instead of
 * O(N²) while still producing a sensible alignment.
 */

/** Pairs of indices `[a0, b0, a1, b1, …]` of elements matched as equal, in order. */
export type Matches = number[];

/** Total Myers work budget per bisection, in diagonal steps. */
const WORK_BUDGET = 40_000_000;

/**
 * Longest-common-subsequence alignment of two integer sequences: returns
 * matched index pairs, increasing in both. Optimal (Myers) for small edit
 * distances; falls back to patience anchoring when the work budget runs out,
 * which stays near-linear but may not be minimal.
 */
export function alignSequences(a: Int32Array, b: Int32Array): number[] {
  const out: Matches = [];
  diffRange(a, b, 0, a.length, 0, b.length, out);
  return out;
}

function diffRange(
  a: Int32Array,
  b: Int32Array,
  aLo: number,
  aHi: number,
  bLo: number,
  bHi: number,
  out: Matches,
): void {
  // Common prefix.
  while (aLo < aHi && bLo < bHi && a[aLo] === b[bLo]) {
    out.push(aLo++, bLo++);
  }
  // Common suffix (emitted after the middle, so remember it).
  let aEnd = aHi;
  let bEnd = bHi;
  while (aEnd > aLo && bEnd > bLo && a[aEnd - 1] === b[bEnd - 1]) {
    aEnd--;
    bEnd--;
  }
  if (aLo < aEnd && bLo < bEnd) {
    const n = aEnd - aLo;
    const m = bEnd - bLo;
    const cap = Math.max(64, Math.floor(WORK_BUDGET / (n + m)));
    const split = bisect(a, b, aLo, n, bLo, m, cap);
    if (split === NO_COMMON) {
      // Nothing in common: everything in the middle is an edit.
    } else if (split === OVER_BUDGET) {
      patience(a, b, aLo, aEnd, bLo, bEnd, out);
    } else {
      const x = split[0];
      const y = split[1];
      diffRange(a, b, aLo, aLo + x, bLo, bLo + y, out);
      diffRange(a, b, aLo + x, aEnd, bLo + y, bEnd, out);
    }
  }
  for (let i = 0; i < aHi - aEnd; i++) out.push(aEnd + i, bEnd + i);
}

const NO_COMMON = 0;
const OVER_BUDGET = 1;

/** Find the middle snake. Returns the split point relative to the range. */
function bisect(
  a: Int32Array,
  b: Int32Array,
  aLo: number,
  n: number,
  bLo: number,
  m: number,
  cap: number,
): [number, number] | typeof NO_COMMON | typeof OVER_BUDGET {
  const maxD = Math.ceil((n + m) / 2);
  const dLimit = Math.min(maxD, cap);
  const vOffset = dLimit + 1;
  const vLength = 2 * dLimit + 2;
  const v1 = new Int32Array(vLength).fill(-1);
  const v2 = new Int32Array(vLength).fill(-1);
  v1[vOffset + 1] = 0;
  v2[vOffset + 1] = 0;
  const delta = n - m;
  const front = (delta & 1) !== 0;
  let k1start = 0;
  let k1end = 0;
  let k2start = 0;
  let k2end = 0;
  for (let d = 0; d < dLimit; d++) {
    for (let k1 = -d + k1start; k1 <= d - k1end; k1 += 2) {
      const k1Offset = vOffset + k1;
      let x1 =
        k1 === -d || (k1 !== d && v1[k1Offset - 1] < v1[k1Offset + 1])
          ? v1[k1Offset + 1]
          : v1[k1Offset - 1] + 1;
      let y1 = x1 - k1;
      while (x1 < n && y1 < m && a[aLo + x1] === b[bLo + y1]) {
        x1++;
        y1++;
      }
      v1[k1Offset] = x1;
      if (x1 > n) k1end += 2;
      else if (y1 > m) k1start += 2;
      else if (front) {
        const k2Offset = vOffset + delta - k1;
        if (k2Offset >= 0 && k2Offset < vLength && v2[k2Offset] !== -1) {
          if (x1 >= n - v2[k2Offset]) return [x1, y1];
        }
      }
    }
    for (let k2 = -d + k2start; k2 <= d - k2end; k2 += 2) {
      const k2Offset = vOffset + k2;
      let x2 =
        k2 === -d || (k2 !== d && v2[k2Offset - 1] < v2[k2Offset + 1])
          ? v2[k2Offset + 1]
          : v2[k2Offset - 1] + 1;
      let y2 = x2 - k2;
      while (x2 < n && y2 < m && a[aLo + n - x2 - 1] === b[bLo + m - y2 - 1]) {
        x2++;
        y2++;
      }
      v2[k2Offset] = x2;
      if (x2 > n) k2end += 2;
      else if (y2 > m) k2start += 2;
      else if (!front) {
        const k1Offset = vOffset + delta - k2;
        if (k1Offset >= 0 && k1Offset < vLength && v1[k1Offset] !== -1) {
          const x1 = v1[k1Offset];
          const y1 = x1 - (k1Offset - vOffset);
          if (x1 >= n - x2) return [x1, y1];
        }
      }
    }
  }
  return dLimit === maxD ? NO_COMMON : OVER_BUDGET;
}

/** Anchor on elements unique to both sides, then diff between anchors. */
function patience(
  a: Int32Array,
  b: Int32Array,
  aLo: number,
  aHi: number,
  bLo: number,
  bHi: number,
  out: Matches,
): void {
  // count[id] packs (countA, lastIndexA, countB, lastIndexB).
  const seen = new Map<number, [number, number, number, number]>();
  for (let i = aLo; i < aHi; i++) {
    const e = seen.get(a[i]);
    if (e) e[0]++;
    else seen.set(a[i], [1, i, 0, -1]);
  }
  for (let j = bLo; j < bHi; j++) {
    const e = seen.get(b[j]);
    if (e) {
      e[2]++;
      e[3] = j;
    }
  }
  const candA: number[] = [];
  const candB: number[] = [];
  for (let i = aLo; i < aHi; i++) {
    const e = seen.get(a[i])!;
    if (e[0] === 1 && e[2] === 1) {
      candA.push(i);
      candB.push(e[3]);
    }
  }
  if (candA.length === 0) return;
  const lis = longestIncreasing(candB);
  let pa = aLo;
  let pb = bLo;
  for (const idx of lis) {
    const ia = candA[idx];
    const ib = candB[idx];
    if (pa < ia && pb < ib) diffGapNoPatience(a, b, pa, ia, pb, ib, out);
    out.push(ia, ib);
    pa = ia + 1;
    pb = ib + 1;
  }
  if (pa < aHi && pb < bHi) diffGapNoPatience(a, b, pa, aHi, pb, bHi, out);
}

/**
 * Between patience anchors, run Myers again but never recurse into another
 * patience pass on the same data; an over-budget gap is left as pure edits.
 */
function diffGapNoPatience(
  a: Int32Array,
  b: Int32Array,
  aLo: number,
  aHi: number,
  bLo: number,
  bHi: number,
  out: Matches,
): void {
  const n = aHi - aLo;
  const m = bHi - bLo;
  if ((n + m) * Math.min(n, m) <= WORK_BUDGET) diffRange(a, b, aLo, aHi, bLo, bHi, out);
}

/** Indices (into `xs`) of a longest strictly increasing subsequence. */
function longestIncreasing(xs: number[]): number[] {
  const tails: number[] = [];
  const prev = new Int32Array(xs.length);
  for (let i = 0; i < xs.length; i++) {
    let lo = 0;
    let hi = tails.length;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (xs[tails[mid]] < xs[i]) lo = mid + 1;
      else hi = mid;
    }
    prev[i] = lo > 0 ? tails[lo - 1] : -1;
    tails[lo] = i;
  }
  const out = new Array<number>(tails.length);
  let k = tails.length ? tails[tails.length - 1] : -1;
  for (let i = tails.length - 1; i >= 0; i--) {
    out[i] = k;
    k = prev[k];
  }
  return out;
}
