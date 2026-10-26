// Shapley contributions against LibreOffice, on the real-file fixtures.
// For each fixture with at most MAX_EDITS edits and no added or removed rows,
// writes the edits (to undo in LibreOffice) and spreadiff's Shapley values per
// computed cell; shapley_oracle.py recalculates every subset in LibreOffice.
// Usage: npm run build && node bench/trace-eval/shapley.mjs [MAX_EDITS] OUT.json
import { readFileSync, readdirSync, existsSync, writeFileSync, mkdirSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
const { parseXlsx, diffWorkbooks, traceWorkbook, columnLabel } = await import('../../packages/spreadiff/dist/core/index.js');
const root = fileURLToPath(new URL('../../test/hosted/fixtures', import.meta.url));
const MAX_EDITS = +process.argv[2] || 8;
const out = [];
const addr = (wd, ref, side) => {
  const d = wd.sheets[ref.sheet].diff, t = side === 'new' ? d.newTable : d.oldTable;
  const ri = side === 'new' ? d.newRows[ref.row] : d.oldRows[ref.row], ci = side === 'new' ? d.columns[ref.col].newIndex : d.columns[ref.col].oldIndex;
  if (ri < 0 || ci < 0) return null;
  const o = t.origin ?? { row: 1, col: 0 };
  return { sheet: t.name, cell: `${columnLabel(ci + o.col)}${ri + o.row + 1}`, value: t.rows[ri][ci], formula: t.formulas?.[ri]?.[ci] ?? null };
};
for (const dir of readdirSync(root)) {
  const o = join(root, dir, 'old.xlsx'), n = join(root, dir, 'new.xlsx');
  if (!existsSync(o) || !existsSync(n)) continue;
  const wd = diffWorkbooks(parseXlsx(readFileSync(o).buffer), parseXlsx(readFileSync(n).buffer));
  const tr = traceWorkbook(wd, { contributions: 'shapley' });
  const edits = [], cells = [];
  wd.sheets.forEach((s, si) => {
    const w = s.diff.columns.length;
    for (const [k, t] of tr.sheet(si)) {
      const ref = { sheet: si, row: Math.floor(k / w), col: k % w };
      if (t.kind === 'input' || t.kind === 'formula') {
        const a = addr(wd, ref, 'new'), b = addr(wd, ref, 'old');
        if (a && b) edits.push({ ...a, kind: t.kind, oldValue: b.value, oldFormula: b.formula, key: `${a.sheet}!${a.cell}` });
        else edits.push({ structural: true });
      }
      if (t.kind === 'computed' && t.attribution === 'shapley' && !t.stale) {
        const a = addr(wd, ref, 'new');
        const shares = {};
        t.causes.forEach((c, i) => { const ca = addr(wd, c, 'new'); if (ca) shares[`${ca.sheet}!${ca.cell}`] = t.contributions[i]; });
        if (Object.values(shares).every((v) => typeof v === 'number')) cells.push({ sheet: a.sheet, cell: a.cell, shares });
      }
    }
  });
  const structural = edits.some((e) => e.structural) || wd.sheets.some((s) => s.diff.stats.rowsAdded || s.diff.stats.rowsRemoved);
  console.log(dir, 'edits', edits.length, 'shapley cells', cells.length, structural ? '(structural: skipped)' : '');
  if (!structural && edits.length && edits.length <= MAX_EDITS && cells.length) out.push({ dir, file: n, edits, cells });
}
mkdirSync(dirname(process.argv[3]), { recursive: true });
writeFileSync(process.argv[3], JSON.stringify(out, null, 1));
