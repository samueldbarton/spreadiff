// The page `npm run test:hosted` builds into a fresh Vite app, as a site
// using the published package would. For every fixture it parses, diffs and
// traces in the library's worker, then compares each changed cell with the
// expectations exported from LibreOffice (bench/trace-eval/export-fixtures.ts).
import { createDiffWorker } from 'spreadiff/worker';
import { applyRecalculation, columnLabel, DiffGrid, traceWorkbook, type CellValue, type WorkbookDiff } from 'spreadiff';

interface Expected {
  source: string;
  files: Record<string, string>;
  cells: Record<string, { kind: string; causes: string[] | null }>;
  values: Record<string, CellValue>;
  staleCheck: string[];
}

// Main-thread responsiveness: the longest gap between ticks of a 10 ms timer.
// (The Long Tasks API reports nothing in headless Chromium.)
let lastTick = performance.now();
let longestGap = 0;
setInterval(() => {
  const now = performance.now();
  longestGap = Math.max(longestGap, now - lastTick);
  lastTick = now;
}, 10);

const worker = createDiffWorker(new Worker(new URL('spreadiff/worker/worker.js', import.meta.url), { type: 'module' }));
const bytes = async (url: string) => (await fetch(url)).arrayBuffer();

function addressOf(d: WorkbookDiff, sheet: number, row: number, col: number): string | null {
  const s = d.sheets[sheet].diff;
  const ri = s.newRows[row];
  const ci = s.columns[col].newIndex;
  if (ri < 0 || ci < 0) return null;
  const o = s.newTable.origin ?? { row: 1, col: 0 };
  return `${s.newTable.name}!${columnLabel(ci + o.col)}${ri + o.row + 1}`;
}

/** New-side value at an address. */
function valueAt(d: WorkbookDiff, address: string): CellValue | undefined {
  const bang = address.lastIndexOf('!');
  const name = address.slice(0, bang);
  const m = /^([A-Z]+)(\d+)$/.exec(address.slice(bang + 1))!;
  let c = 0;
  for (const ch of m[1]) c = c * 26 + ch.charCodeAt(0) - 64;
  const t = d.sheets.find((s) => s.diff.newTable.name === name)?.diff.newTable;
  if (!t) return undefined;
  const o = t.origin ?? { row: 1, col: 0 };
  return t.rows[Number(m[2]) - 1 - o.row]?.[c - 1 - o.col];
}

const same = (a: CellValue | undefined, b: CellValue | undefined) =>
  typeof a === 'number' && typeof b === 'number' ? Math.abs(a - b) <= 1e-9 * Math.max(1, Math.abs(a), Math.abs(b)) : (a ?? null) === (b ?? null) || String(a ?? '') === String(b ?? '');

const report: Record<string, unknown>[] = [];
const cases: string[] = await (await fetch('fixtures/index.json')).json();
for (const name of cases) {
  const base = `fixtures/${encodeURIComponent(name)}/`;
  const exp: Expected = await (await fetch(base + 'expected.json')).json();
  const problems: string[] = [];
  const old = await bytes(base + 'old.xlsx');

  // 1. Causes and kinds of every changed cell.
  const t0 = performance.now();
  const { result, trace, timings } = await worker.diffBuffers({ name: 'old.xlsx', data: old.slice(0) }, { name: 'new.xlsx', data: await bytes(base + 'new.xlsx') }, {}, { trace: true });
  const elapsed = performance.now() - t0;
  if (!trace) problems.push('no trace returned');
  let scored = 0;
  const seen = new Set<string>();
  result.sheets.forEach((s, si) => {
    for (const [row, mask] of s.diff.changedCells) for (let col = 0; col < mask.length; col++) {
      if (!mask[col]) continue;
      const a = addressOf(result, si, row, col);
      if (!a) continue;
      seen.add(a);
      const want = exp.cells[a];
      const got = trace?.get({ sheet: si, row, col });
      if (!want) { problems.push(`unexpected change ${a}`); continue; }
      if (got?.kind !== want.kind) { problems.push(`kind ${a}: got ${got?.kind}, want ${want.kind}`); continue; }
      if (want.kind !== 'computed' || !want.causes) continue;
      // Added/removed rows are causes too, but have no single-edit oracle: compare the edits.
      const causes = got.causes
        .filter((c) => { const k = trace!.get(c)?.kind; return k !== 'added' && k !== 'removed'; })
        .map((c) => addressOf(result, c.sheet, c.row, c.col))
        .sort();
      if (got.causeCount > got.causes.length) problems.push(`causes of ${a} truncated`);
      if (JSON.stringify(causes) !== JSON.stringify(want.causes)) problems.push(`causes ${a}: got ${causes.join(' ')}, want ${want.causes.join(' ')}`);
      else scored++;
    }
  });
  for (const a of Object.keys(exp.cells)) if (!seen.has(a)) problems.push(`missed change ${a}`);

  // 2. Saved without recalculating: recalculated values must be LibreOffice's.
  let staleOk = 0;
  if (exp.files.stale) {
    const r = await worker.diffBuffers({ name: 'old.xlsx', data: old.slice(0) }, { name: 'stale.xlsx', data: await bytes(base + 'stale.xlsx') }, {}, { trace: true });
    if (!r.trace?.stats.stale) problems.push('stale file: nothing flagged stale');
    const shown = r.trace ? applyRecalculation(r.result, r.trace) : r.result;
    for (const a of exp.staleCheck) {
      if (same(valueAt(shown, a), exp.values[a])) staleOk++;
      else problems.push(`stale ${a}: recalculated ${valueAt(shown, a)}, LibreOffice ${exp.values[a]}`);
    }
  }

  // 3. Written without values: computed on parse must equal LibreOffice's.
  let uncachedOk = 0;
  if (exp.files.uncached) {
    const r = await worker.diffBuffers({ name: 'old.xlsx', data: old.slice(0) }, { name: 'uncached.xlsx', data: await bytes(base + 'uncached.xlsx') });
    for (const [a, v] of Object.entries(exp.values)) {
      if (same(valueAt(r.result, a), v)) uncachedOk++;
      else if (problems.length < 40) problems.push(`uncached ${a}: computed ${valueAt(r.result, a)}, LibreOffice ${v}`);
    }
  }

  // 4. The grid renders the traced diff.
  const host = document.getElementById('grid')!;
  const grid = new DiffGrid({ trace: trace ? { trace, workbook: result, sheet: 0 } : undefined });
  grid.render({ diff: result.sheets[0].diff, container: host });
  grid.cleanUp();

  report.push({
    name, source: exp.source, cells: Object.keys(exp.cells).length, scored, staleOk, staleChecked: exp.staleCheck.length,
    uncachedOk, uncachedChecked: exp.files.uncached ? Object.keys(exp.values).length : 0,
    ms: Math.round(elapsed), workerTrace: Math.round(timings.trace), stats: trace?.stats, problems,
  });
}
// Control: the same trace on the main thread must show up as a stall, or the measurement above means nothing.
const workerGap = longestGap;
const inv = await worker.diffBuffers({ name: 'old.xlsx', data: await bytes('fixtures/inventory/old.xlsx') }, { name: 'new.xlsx', data: await bytes('fixtures/inventory/new.xlsx') });
await new Promise((r) => setTimeout(r, 50));
longestGap = 0;
lastTick = performance.now();
traceWorkbook(inv.result);
await new Promise((r) => setTimeout(r, 50));
(window as unknown as { done: unknown }).done = { report, longestGap: workerGap, controlGap: longestGap };
