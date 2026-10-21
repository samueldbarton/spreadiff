import { describe, expect, it } from 'vitest';
import {
  diffTables, parseCsv, diffText, packTable, unpackTable, ROW_STATUS, cellsEqual,
  type CellValue, type DiffOptions, type Table,
} from '../src/core';
import { computeRowLayout, ROW_KIND } from '../src/render/layout';

function rng(seed: number) {
  return () => {
    seed = (seed + 0x6d2b79f5) | 0;
    let t = Math.imul(seed ^ (seed >>> 15), 1 | seed);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

const POOL: CellValue[] = [null, '', 'a', 'A', ' a', 'b', 'x y', '1', '1.0', 1, 1.0000001, 2, -0, 0, true, false, '=1+1', 'é', '😀', 'a,b', 'q"q', 'line\nbreak'];

function randomTable(r: () => number, cols: string[], n: number, keyed: boolean): Table {
  return {
    name: 's',
    columns: cols,
    rows: Array.from({ length: n }, (_, i) =>
      cols.map((c, j) => (keyed && j === 0 ? `k${Math.floor(r() * n * 1.5)}${i}` : POOL[Math.floor(r() * POOL.length)])),
    ),
  };
}

function mutate(r: () => number, t: Table): Table {
  const rows = t.rows.map((row) => [...row]);
  const ops = Math.floor(r() * 6);
  for (let k = 0; k < ops; k++) {
    const op = r();
    if (op < 0.3 && rows.length) rows.splice(Math.floor(r() * rows.length), 1);
    else if (op < 0.55) rows.splice(Math.floor(r() * (rows.length + 1)), 0, t.columns.map(() => POOL[Math.floor(r() * POOL.length)]));
    else if (op < 0.85 && rows.length) {
      const row = rows[Math.floor(r() * rows.length)];
      row[Math.floor(r() * row.length)] = POOL[Math.floor(r() * POOL.length)];
    } else if (rows.length > 1) {
      const i = Math.floor(r() * rows.length);
      const j = Math.floor(r() * rows.length);
      [rows[i], rows[j]] = [rows[j], rows[i]];
    }
  }
  let columns = [...t.columns];
  if (r() < 0.2) columns = columns.map((c, i) => (i === columns.length - 1 ? c + '_renamed' : c));
  if (r() < 0.15) {
    columns.push('extra');
    rows.forEach((row) => row.push(POOL[Math.floor(r() * POOL.length)]));
  }
  if (r() < 0.15 && columns.length > 2) {
    columns.splice(1, 1);
    rows.forEach((row) => row.splice(1, 1));
  }
  return { name: 's', columns, rows };
}

describe('fuzz: diffTables invariants', () => {
  it('holds over 3000 random table pairs and option sets', () => {
    const r = rng(1234);
    for (let iter = 0; iter < 3000; iter++) {
      const keyed = r() < 0.4;
      const cols = keyed ? ['id', 'a', 'b', 'c'] : ['a', 'b', 'c'].slice(0, 1 + Math.floor(r() * 3));
      const a = randomTable(r, cols, Math.floor(r() * 25), keyed);
      const b = mutate(r, a);
      const opts: DiffOptions = {
        keys: r() < 0.7 ? 'auto' : 'none',
        ignoreCase: r() < 0.3,
        trimWhitespace: r() < 0.7,
        numericTolerance: r() < 0.2 ? 0.001 : 0,
        detectRenames: r() < 0.8,
      };
      const ctx = `iter ${iter} ${JSON.stringify({ opts, a, b })}`;
      const d = diffTables(a, b, opts);
      // Every old and new row appears exactly once.
      const seenOld = new Uint8Array(a.rows.length);
      const seenNew = new Uint8Array(b.rows.length);
      let lastNew = -1;
      for (let i = 0; i < d.rowCount; i++) {
        const o = d.oldRows[i];
        const n = d.newRows[i];
        expect(o >= 0 || n >= 0, ctx).toBe(true);
        if (o >= 0) { expect(seenOld[o], ctx).toBe(0); seenOld[o] = 1; }
        if (n >= 0) {
          expect(seenNew[n], ctx).toBe(0);
          seenNew[n] = 1;
          expect(n, `new rows out of order: ${ctx}`).toBeGreaterThan(lastNew);
          lastNew = n;
        }
        const st = d.status[i];
        if (o < 0) expect(st, ctx).toBe(ROW_STATUS.added);
        else if (n < 0) expect(st, ctx).toBe(ROW_STATUS.removed);
        else {
          const mask = d.changedCells.get(i);
          d.columns.forEach((c, ci) => {
            if (c.oldIndex < 0 || c.newIndex < 0) return;
            const eq = cellsEqual(a.rows[o][c.oldIndex] ?? null, b.rows[n][c.newIndex] ?? null, opts);
            expect(!!mask?.[ci], `mask mismatch col ${c.name}: ${ctx}`).toBe(!eq);
          });
          expect(st, ctx).toBe(mask ? ROW_STATUS.modified : ROW_STATUS.equal);
        }
      }
      expect(seenOld.every((x) => x === 1), `old row missing: ${ctx}`).toBe(true);
      expect(seenNew.every((x) => x === 1), `new row missing: ${ctx}`).toBe(true);
      // Old rows stay in order in sequence mode.
      if (!d.keyColumns.length) {
        let last = -1;
        for (let i = 0; i < d.rowCount; i++) if (d.oldRows[i] >= 0) { expect(d.oldRows[i], ctx).toBeGreaterThan(last); last = d.oldRows[i]; }
      }
      // Columns: each old and new column exactly once.
      const oc = d.columns.filter((c) => c.oldIndex >= 0).map((c) => c.oldIndex).sort((x, y) => x - y);
      const nc = d.columns.filter((c) => c.newIndex >= 0).map((c) => c.newIndex).sort((x, y) => x - y);
      expect(oc, ctx).toEqual(a.columns.map((_, i) => i));
      expect(nc, ctx).toEqual(b.columns.map((_, i) => i));
      // Stats agree with the arrays.
      const count = (s: number) => Array.from(d.status).filter((x) => x === s).length;
      expect(d.stats.rowsAdded).toBe(count(ROW_STATUS.added));
      expect(d.stats.rowsRemoved).toBe(count(ROW_STATUS.removed));
      expect(d.stats.rowsModified).toBe(count(ROW_STATUS.modified));
      expect(d.stats.cellsChanged).toBe([...d.changedCells.values()].reduce((s, m) => s + m.reduce((x, y) => x + y, 0), 0));
      expect(d.columns.reduce((s, c) => s + c.changes, 0)).toBe(d.stats.cellsChanged);
      // Identical inputs → no changes.
      if (iter % 10 === 0) {
        const same = diffTables(a, a, opts);
        expect(same.stats.rowsAdded + same.stats.rowsRemoved + same.stats.rowsModified, ctx).toBe(0);
      }
    }
  });
});

describe('fuzz: CSV round-trip', () => {
  const serialize = (rows: string[][], d: string) =>
    rows.map((row) => row.map((f) => (/[",\n\r\t;|]/.test(f) || f === '' ? `"${f.replace(/"/g, '""')}"` : f)).join(d)).join('\r\n');
  it('parse(serialize(x)) === x for random fields', () => {
    const r = rng(99);
    const alphabet = ['a', 'b', ' ', '"', ',', '\n', '\r\n', '\t', ';', '|', 'é', '😀', '0', '7', '.', '-'];
    for (let iter = 0; iter < 1500; iter++) {
      const w = 1 + Math.floor(r() * 4);
      const h = 1 + Math.floor(r() * 6);
      const field = () => Array.from({ length: Math.floor(r() * 6) }, () => alphabet[Math.floor(r() * alphabet.length)]).join('');
      const header = Array.from({ length: w }, (_, i) => `h${i}`);
      const body = Array.from({ length: h }, () => Array.from({ length: w }, field));
      // Drop trailing all-empty rows (the parser trims them by design).
      while (body.length && body[body.length - 1].every((f) => f === '')) body.pop();
      const d = [',', '\t', ';'][Math.floor(r() * 3)];
      const t = parseCsv(serialize([header, ...body], d), { delimiter: d, numbers: false });
      const ctx = JSON.stringify({ d, body });
      expect(t.columns, ctx).toEqual(header);
      expect(t.rows.map((row) => row.map((v) => v ?? '')), ctx).toEqual(body);
    }
  });
});

describe('fuzz: diffText', () => {
  it('reconstructs both sides exactly', () => {
    const r = rng(5);
    const alphabet = ['a', 'b', 'c', ' ', '😀', 'é', '.', 'word', 'other'];
    for (let iter = 0; iter < 2000; iter++) {
      const s = () => Array.from({ length: Math.floor(r() * 30) }, () => alphabet[Math.floor(r() * alphabet.length)]).join('');
      const a = s();
      const b = s();
      const spans = diffText(a, b);
      expect(spans.filter((x) => x.kind !== 'added').map((x) => x.text).join(''), JSON.stringify({ a, b })).toBe(a);
      expect(spans.filter((x) => x.kind !== 'removed').map((x) => x.text).join(''), JSON.stringify({ a, b })).toBe(b);
    }
  });
});

describe('fuzz: row layout', () => {
  it('shows every changed row and accounts for every row exactly once', () => {
    const r = rng(77);
    for (let iter = 0; iter < 500; iter++) {
      const a = randomTable(r, ['a', 'b'], Math.floor(r() * 80), false);
      const b = mutate(r, a);
      const d = diffTables(a, b, { keys: 'none' });
      const context = Math.floor(r() * 4);
      const l = computeRowLayout({
        diff: d, collapse: r() < 0.8, context, expanded: new Set(), rowHeight: 30, modifiedHeight: 48,
        separatorHeight: 30, annotationsAfter: new Map(),
      });
      const covered = new Uint8Array(d.rowCount);
      for (let i = 0; i < l.count; i++) {
        if (l.kind[i] === ROW_KIND.data) { expect(covered[l.ref[i]]).toBe(0); covered[l.ref[i]] = 1; expect(l.displayOf[l.ref[i]]).toBe(i); }
        if (l.kind[i] === ROW_KIND.separator) {
          const g = l.regions[l.ref[i]];
          for (let k = g.start; k < g.end; k++) { expect(covered[k]).toBe(0); covered[k] = 1; expect(d.status[k]).toBe(ROW_STATUS.equal); }
        }
        expect(l.offsets[i + 1]).toBeGreaterThan(l.offsets[i]);
      }
      expect(covered.every((x) => x === 1)).toBe(true);
    }
  });
});

describe('fuzz: transfer', () => {
  it('round-trips random tables', () => {
    const r = rng(3);
    for (let iter = 0; iter < 500; iter++) {
      const t = randomTable(r, ['a', 'b', 'c'], Math.floor(r() * 20), false);
      const u = unpackTable(structuredClone(packTable(t)));
      expect(u.rows.map((row) => row.map((v) => (typeof v === 'string' ? v.replace(/�/g, '\u0000') : v)))).toEqual(t.rows);
    }
  });
});
