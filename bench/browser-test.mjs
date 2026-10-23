// Browser regression tests for the canvas renderer.
// Usage: (cd apps/viewer && npx vite --port 5199) & node bench/browser-test.mjs
import { chromium } from 'playwright';

const URL = process.env.URL ?? 'http://localhost:5199/test.html';
// Headless Chromium hides scrollbars by default; the scrollbar test needs real ones.
const browser = await chromium.launch({ executablePath: process.env.CHROMIUM ?? '/opt/pw-browsers/chromium', ignoreDefaultArgs: ['--hide-scrollbars'] });
const page = await browser.newPage({ viewport: { width: 1000, height: 700 } });
const errors = [];
page.on('pageerror', (e) => errors.push(e.message));
await page.goto(URL);
await page.waitForFunction(() => window.ready);

let failed = 0;
async function test(name, fn) {
  try {
    await fn();
    console.log(`  ✓ ${name}`);
  } catch (e) {
    failed++;
    console.log(`  ✗ ${name}\n    ${e.message}`);
  }
}
const assert = (cond, msg) => { if (!cond) throw new Error(msg); };

// Fresh grid per test: N rows, one changed cell every 10 rows.
async function setup(opts = {}, rows = 300) {
  await page.evaluate(({ opts, rows }) => {
    window.G?.cleanUp();
    const { parseCsv, diffTables, DiffGrid } = window.sd;
    const mk = (bump) => 'id,a,b,c\n' + Array.from({ length: rows }, (_, i) => `${i},${i % 10 === 0 && bump ? 'x' : 'v'}${i},${i * 2},c`).join('\n');
    window.diff = diffTables(parseCsv(mk(false)), parseCsv(mk(true)));
    window.G = new DiffGrid({ collapseUnchanged: false, tooltips: false, ...opts });
    window.G.render({ diff: window.diff, container: document.getElementById('grid') });
  }, { opts, rows });
  await page.waitForTimeout(150);
}
const internals = () => page.evaluate(() => {
  const g = window.G;
  return { width: g.width, height: g.height, clientWidth: g.scroller.clientWidth, clientHeight: g.scroller.clientHeight, hover: g.hover, selected: g.selected };
});
const scrollerBox = () => page.evaluate(() => {
  const r = window.G.scroller.getBoundingClientRect();
  return { x: r.x, y: r.y };
});

await test('canvas excludes the scrollbar (observes scroller, not host)', async () => {
  // Start small (no scrollbar), then grow: the scrollbar appears later and
  // the host's own size never changes.
  await setup({}, 3);
  await page.evaluate(() => {
    const { parseCsv, diffTables } = window.sd;
    const mk = (v) => 'id,a\n' + Array.from({ length: 400 }, (_, i) => `${i},${v}${i}`).join('\n');
    window.G.setDiff(diffTables(parseCsv(mk('a')), parseCsv(mk('b'))));
  });
  await page.waitForTimeout(150);
  const s = await internals();
  assert(s.width === s.clientWidth, `canvas width ${s.width} != scroller clientWidth ${s.clientWidth}`);
  assert(s.clientWidth < 900, 'expected a vertical scrollbar to reduce clientWidth');
});

await test('setOptions with unchanged values does not relayout', async () => {
  await setup();
  const n = await page.evaluate(() => {
    let calls = 0;
    const g = window.G;
    const orig = g.relayout.bind(g);
    g.relayout = () => { calls++; orig(); };
    for (let i = 0; i < 50; i++) g.setOptions({ layout: 'unified', collapseUnchanged: false, columnWidths: { a: 120 }, onCellClick: () => {} });
    return calls;
  });
  assert(n <= 1, `relayout ran ${n} times for 50 no-op setOptions calls`);
});

await test('column resize recomputes columns only, not every row', async () => {
  await setup({}, 2000);
  const box = await scrollerBox();
  const edge = await page.evaluate(() => {
    const g = window.G;
    const p = g.panes()[0];
    return p.x + p.gutter + g.cols.offsets[2]; // right edge of 2nd visible column
  });
  const calls = await page.evaluate(() => {
    window.__rowLayouts = 0;
    const g = window.G;
    const orig = g.relayout.bind(g);
    g.relayout = () => { window.__rowLayouts++; orig(); };
    return 0;
  });
  void calls;
  await page.mouse.move(box.x + edge, box.y + 18);
  await page.mouse.down();
  for (let i = 1; i <= 10; i++) await page.mouse.move(box.x + edge + i * 5, box.y + 18);
  await page.mouse.up();
  const { rowLayouts, width } = await page.evaluate(() => ({ rowLayouts: window.__rowLayouts, width: window.G.cols.widths[1] }));
  assert(rowLayouts === 0, `row layout ran ${rowLayouts} times during a resize drag`);
  assert(width > 60, `column did not resize (width ${width})`);
});

await test('pointercancel ends a resize drag', async () => {
  await setup();
  const stuck = await page.evaluate(() => {
    const g = window.G;
    g.resizing = { vcol: 0, startX: 0, startW: 80 };
    // Simulate the listener wiring from onPointerDown by calling the real path:
    return null;
  });
  void stuck;
  const box = await scrollerBox();
  const edge = await page.evaluate(() => { const g = window.G; const p = g.panes()[0]; return p.x + p.gutter + g.cols.offsets[2]; });
  await page.evaluate(() => { window.G.resizing = null; });
  await page.mouse.move(box.x + edge, box.y + 18);
  await page.mouse.down();
  await page.evaluate(() => window.G.scroller.dispatchEvent(new PointerEvent('pointercancel', { bubbles: true })));
  const s = await page.evaluate(() => window.G.resizing);
  await page.mouse.up();
  assert(s === null, 'still resizing after pointercancel');
});

await test('hover follows the cursor after a wheel scroll', async () => {
  await setup();
  const box = await scrollerBox();
  await page.mouse.move(box.x + 300, box.y + 120);
  await page.waitForTimeout(50);
  const before = (await internals()).hover.display;
  await page.mouse.wheel(0, 300);
  await page.waitForTimeout(150);
  const after = (await internals()).hover.display;
  assert(after > before, `hover stayed on display row ${before} (now ${after}) after scrolling`);
});

await test('Tab is not trapped inside the grid', async () => {
  await setup();
  await page.focus('#before');
  await page.keyboard.press('Tab');
  const inGrid = await page.evaluate(() => document.activeElement?.className);
  await page.keyboard.press('Tab');
  const next = await page.evaluate(() => document.activeElement?.id);
  assert(inGrid === 'spreadiff', `focus did not enter grid (${inGrid})`);
  assert(next === 'after', `Tab did not leave the grid (active: ${next})`);
});

await test('ArrowDown with no selection selects a data row in view, not a separator ref', async () => {
  await setup({ collapseUnchanged: true }, 300);
  await page.evaluate(() => window.G.scroller.focus());
  await page.keyboard.press('ArrowDown');
  const s = await internals();
  const kind = await page.evaluate(() => window.G.rows.kind[window.G.rows.displayOf[window.G.selected.row]]);
  assert(s.selected && kind === 0, `selection ${JSON.stringify(s.selected)} is not a visible data row`);
});

await test('row numbers follow the table origin (header=false → starts at 1)', async () => {
  await page.evaluate(() => {
    window.G?.cleanUp();
    const { parseCsv, diffTables, DiffGrid } = window.sd;
    const d = diffTables(parseCsv('a,b\n1,2', { header: false }), parseCsv('a,b\n1,3', { header: false }));
    window.G = new DiffGrid({ collapseUnchanged: false });
    window.G.render({ diff: d, container: document.getElementById('grid') });
  });
  const off = await page.evaluate(() => [window.G.rowOffset('old'), window.G.rowOffset('new')]);
  assert(off[0] === 1 && off[1] === 1, `row offset ${off}`);
});

await test('formula-only change shows the formulas, not identical values', async () => {
  const texts = await page.evaluate(() => {
    window.G?.cleanUp();
    const { diffTables, DiffGrid } = window.sd;
    const t = (f) => ({ name: 's', columns: ['v'], rows: [[2]], formulas: [[f]] });
    const d = diffTables(t('=1+1'), t('=4/2'));
    window.G = new DiffGrid({ collapseUnchanged: false });
    const drawn = [];
    const orig = window.G.text.bind(window.G);
    window.G.text = (s, ...rest) => { drawn.push(s); return orig(s, ...rest); };
    window.G.render({ diff: d, container: document.getElementById('grid') });
    window.G.drawNow();
    return drawn;
  });
  assert(texts.includes('=1+1') && texts.includes('=4/2'), `drawn: ${texts.join(' | ')}`);
});

await test('cleanUp disconnects annotation observers and removes elements', async () => {
  const left = await page.evaluate(() => {
    window.G?.cleanUp();
    const { parseCsv, diffTables, DiffGrid } = window.sd;
    const d = diffTables(parseCsv('id,v\n1,a\n2,b'), parseCsv('id,v\n1,a\n2,c'));
    const g = new DiffGrid({ annotations: [{ row: 1, data: 1 }, { row: 99, data: 2 }], renderAnnotation: () => Object.assign(document.createElement('div'), { textContent: 'note' }) });
    g.render({ diff: d, container: document.getElementById('grid') });
    g.drawNow(); // out-of-range annotation row must not throw
    g.cleanUp();
    return { els: document.querySelectorAll('.sd-annotation').length, map: g.annotationEls.size };
  });
  assert(left.els === 0 && left.map === 0, JSON.stringify(left));

});

assert(true);
if (errors.length) {
  failed++;
  console.log(`  ✗ page errors:\n    ${errors.join('\n    ')}`);
}
await browser.close();
console.log(failed ? `\n${failed} failed` : '\nall passed');
process.exit(failed ? 1 : 0);
