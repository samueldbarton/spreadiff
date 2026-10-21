// Renderer fuzzer: random operation sequences against DiffGrid in Chromium.
// Usage: (cd apps/viewer && npx vite --port 5199) & node bench/fuzz-render.mjs [iterations]
import { chromium } from 'playwright';

const ITER = Number(process.argv[2] ?? 300);
const SEED = Number(process.argv[3] ?? 7);
const browser = await chromium.launch({ executablePath: process.env.CHROMIUM ?? '/opt/pw-browsers/chromium', ignoreDefaultArgs: ['--hide-scrollbars'] });
const page = await browser.newPage({ viewport: { width: 1000, height: 700 } });
const errors = [];
page.on('pageerror', (e) => errors.push(e.stack || e.message));
await page.goto(process.env.URL ?? 'http://localhost:5199/test.html');
await page.waitForFunction(() => window.ready);

const result = await page.evaluate(async ([ITER, SEED]) => {
  const { parseCsv, diffTables, DiffGrid } = window.sd;
  let seed = SEED;
  const r = () => ((seed = (seed * 1103515245 + 12345) & 0x7fffffff) / 0x7fffffff);
  const pick = (xs) => xs[Math.floor(r() * xs.length)];
  const vals = ['', 'a', 'b', '1', '2.5', 'long text that will need truncation in a narrow column', 'é😀', 'x\ny'];
  const mkTable = (cols, rows) => ({
    name: 's',
    columns: cols,
    rows: Array.from({ length: rows }, (_, i) => cols.map((c, j) => (j === 0 && c === 'id' ? String(i) : pick(vals) || null))),
    formulas: r() < 0.2 ? Array.from({ length: rows }, () => cols.map(() => (r() < 0.1 ? '=A1' : null))) : undefined,
  });
  const mkDiff = () => {
    const kind = r();
    if (kind < 0.08) return diffTables({ name: 's', columns: [], rows: [] }, { name: 's', columns: [], rows: [] });
    if (kind < 0.14) return diffTables({ name: 's', columns: [], rows: [] }, mkTable(['a', 'b'], 5));
    if (kind < 0.2) return diffTables(mkTable(['a'], 0), mkTable(['a'], 0));
    const cols = r() < 0.5 ? ['id', 'a', 'b', 'c', 'd'] : ['a', 'b'];
    const a = mkTable(cols, Math.floor(r() * 400));
    const b = mkTable(r() < 0.2 ? [...cols.slice(0, -1), 'z'] : cols, Math.floor(r() * 400));
    // Make b a near-copy of a most of the time, for realistic sparse diffs.
    if (r() < 0.7) b.rows = a.rows.map((row) => row.map((v) => (r() < 0.03 ? pick(vals) : v)));
    return diffTables(a, b, { keys: pick(['auto', 'none']) });
  };
  const container = document.getElementById('grid');
  const log = [];
  let g;
  const fresh = () => {
    g?.cleanUp();
    g = new DiffGrid({ tooltips: true });
    g.render({ diff: mkDiff(), container });
  };
  fresh();
  const ops = {
    setDiff: () => g.setDiff(mkDiff()),
    layout: () => g.setOptions({ layout: pick(['unified', 'split']) }),
    collapse: () => g.setOptions({ collapseUnchanged: r() < 0.5, context: Math.floor(r() * 5) }),
    hide: () => g.setOptions({ hideUnchangedColumns: r() < 0.5 }),
    frozen: () => g.setOptions({ frozenColumns: pick(['keys', 0, 1, 3, 99, -1]) }),
    minimap: () => g.setOptions({ minimap: r() < 0.5 }),
    widths: () => g.setOptions({ columnWidths: { a: Math.floor(r() * 400), b: 1 } }),
    theme: () => g.setOptions({ theme: pick(['light', 'dark', 'system']) }),
    annotations: () => g.setOptions({
      annotations: Array.from({ length: Math.floor(r() * 4) }, () => ({ row: Math.floor(r() * 500) - 50, column: pick(['a', 'b', 'nope']), data: 1 })),
      renderAnnotation: r() < 0.8 ? () => Object.assign(document.createElement('div'), { textContent: 'note', style: `height:${Math.floor(r() * 80)}px` }) : undefined,
    }),
    size: () => { container.style.width = `${pick([0, 1, 50, 300, 900, 1400])}px`; container.style.height = `${pick([0, 1, 40, 200, 500])}px`; },
    scroll: () => { const s = g.scroller; s.scrollTop = r() * s.scrollHeight; s.scrollLeft = r() * s.scrollWidth; },
    next: () => { for (let i = 0; i < 5; i++) g.nextChange(); },
    prev: () => { for (let i = 0; i < 5; i++) g.prevChange(); },
    expand: () => (r() < 0.5 ? g.expandAll() : g.collapseAll()),
    scrollToRow: () => g.scrollToRow(Math.floor(r() * 600) - 50, pick(['start', 'center', 'nearest'])),
    select: () => { g.options.onSelectionChange = (c) => c && c.column.name; g.select(r() < 0.2 ? null : Math.floor(r() * 500), Math.floor(r() * 6)); },
    keys: () => { g.scroller.focus(); for (let i = 0; i < 4; i++) g.scroller.dispatchEvent(new KeyboardEvent('keydown', { key: pick(['ArrowDown', 'ArrowUp', 'ArrowLeft', 'ArrowRight', 'n', 'p', 'Escape']), bubbles: true })); },
    pointer: () => {
      const rect = g.scroller.getBoundingClientRect();
      for (let i = 0; i < 4; i++) {
        const x = rect.left + r() * (rect.width + 20) - 10;
        const y = rect.top + r() * (rect.height + 20) - 10;
        const type = pick(['pointermove', 'pointermove', 'pointerdown', 'pointerup', 'pointerleave']);
        g.scroller.dispatchEvent(new PointerEvent(type, { clientX: x, clientY: y, bubbles: true, pointerId: 1 }));
      }
    },
    fresh,
  };
  const names = Object.keys(ops);
  const failures = [];
  for (let i = 0; i < ITER; i++) {
    const op = names[Math.floor(r() * names.length)];
    log.push(op);
    try {
      ops[op]();
      // Let ResizeObserver / rAF settle, then draw synchronously.
      if (i % 5 === 0) await new Promise((res) => requestAnimationFrame(() => requestAnimationFrame(res)));
      g.drawNow();
      const rows = g.rows;
      if (rows) {
        for (let d = 0; d < rows.count; d++) if (!(rows.offsets[d + 1] >= rows.offsets[d])) throw new Error(`offsets not monotonic at ${d}`);
      }
      if (g.width && g.width !== g.scroller.clientWidth && i % 5 === 0) throw new Error(`canvas width ${g.width} != clientWidth ${g.scroller.clientWidth}`);
      if (g.selected && g.diff && (g.selected.row < 0 || g.selected.row >= g.diff.rowCount)) throw new Error(`selection out of range: ${JSON.stringify(g.selected)} of ${g.diff.rowCount}`);
      if (g.cols && g.cols.frozen < 0) throw new Error('negative frozen count');
    } catch (e) {
      failures.push({ i, op, error: String(e && e.stack || e), recent: log.slice(-8) });
      if (failures.length > 5) break;
      fresh();
    }
  }
  g.cleanUp();
  return failures;
}, [ITER, SEED]);

await browser.close();
const all = [...result, ...errors.map((e) => ({ error: e }))];
if (all.length) {
  console.log(JSON.stringify(all, null, 2));
  process.exit(1);
}
console.log(`fuzz: ${ITER} ops (seed ${SEED}), no failures`);
