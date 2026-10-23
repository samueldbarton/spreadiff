import { describe, expect, it } from 'vitest';
import { alignSequences } from '../src/core/sequence';

function lcsLength(a: number[], b: number[]): number {
  const dp = new Array(b.length + 1).fill(0);
  for (let i = 1; i <= a.length; i++) {
    let prev = 0;
    for (let j = 1; j <= b.length; j++) {
      const tmp = dp[j];
      dp[j] = a[i - 1] === b[j - 1] ? prev + 1 : Math.max(dp[j], dp[j - 1]);
      prev = tmp;
    }
  }
  return dp[b.length];
}

function rng(seed: number) {
  return () => ((seed = (seed * 1103515245 + 12345) & 0x7fffffff) / 0x7fffffff);
}

describe('alignSequences', () => {
  it('finds an optimal alignment on random inputs', () => {
    const rand = rng(42);
    for (let t = 0; t < 400; t++) {
      const alpha = 1 + Math.floor(rand() * 6);
      const a = Array.from({ length: Math.floor(rand() * 40) }, () => Math.floor(rand() * alpha));
      const b = Array.from({ length: Math.floor(rand() * 40) }, () => Math.floor(rand() * alpha));
      const m = alignSequences(Int32Array.from(a), Int32Array.from(b));
      for (let k = 0; k < m.length; k += 2) {
        expect(a[m[k]]).toBe(b[m[k + 1]]);
        if (k > 0) {
          expect(m[k]).toBeGreaterThan(m[k - 2]);
          expect(m[k + 1]).toBeGreaterThan(m[k - 1]);
        }
      }
      expect(m.length / 2).toBe(lcsLength(a, b));
    }
  });

  it('stays fast and sensible on a large shuffled input (patience fallback)', () => {
    const n = 200_000;
    const a = Int32Array.from({ length: n }, (_, i) => i);
    const b = Int32Array.from(a);
    const rand = rng(7);
    for (let i = n - 1; i > 0; i--) {
      const j = Math.floor(rand() * (i + 1));
      [b[i], b[j]] = [b[j], b[i]];
    }
    const t0 = performance.now();
    const m = alignSequences(a, b);
    const ms = performance.now() - t0;
    expect(ms).toBeLessThan(5000);
    for (let k = 2; k < m.length; k += 2) expect(m[k + 1]).toBeGreaterThan(m[k - 1]);
    expect(m.length / 2).toBeGreaterThan(100); // LIS of a random permutation ~ 2√n ≈ 894
  });

  it('handles a large input with few edits quickly', () => {
    const n = 1_000_000;
    const a = Int32Array.from({ length: n }, (_, i) => i);
    const b = Int32Array.from(a);
    for (let i = 0; i < 50; i++) b[i * 19_997] = -1 - i;
    const t0 = performance.now();
    const m = alignSequences(a, b);
    expect(performance.now() - t0).toBeLessThan(3000);
    expect(m.length / 2).toBe(n - 50);
  });
});
