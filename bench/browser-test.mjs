// Browser regression tests for the canvas renderer.
// Usage: (cd apps/viewer && npx vite --port 5199) & node bench/browser-test.mjs
// BROWSER=firefox or BROWSER=webkit runs the same tests in Playwright's Firefox or WebKit.
import { readFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import playwright from 'playwright';

const URL = process.env.URL ?? 'http://localhost:5199/test.html';
const engine = process.env.BROWSER ?? 'chromium';
const browser = await playwright[engine].launch(
  engine === 'chromium'
    ? // Headless Chromium hides scrollbars by default; the scrollbar test needs real ones.
      { executablePath: process.env.CHROMIUM ?? '/opt/pw-browsers/chromium', ignoreDefaultArgs: ['--hide-scrollbars'] }
    : {},
);
console.log(`${engine} ${browser.version()}`);
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

await test('a horizontal merge draws as one cell; a change inside breaks it', async () => {
  const r = await page.evaluate(() => {
    window.G?.cleanUp();
    const { diffTables, DiffGrid } = window.sd;
    const mk = (v) => ({ name: 's', columns: ['a', 'b', 'c'], rows: [['Q1 report', null, null], [1, 2, v]], merges: [{ r0: 1, c0: 0, r1: 1, c1: 2 }], origin: { row: 1, col: 0 } });
    const d = diffTables(mk(3), mk(4), { keys: 'none' });
    window.G = new DiffGrid({ layout: 'unified', collapseUnchanged: false });
    window.G.render({ diff: d, container: document.getElementById('grid') });
    const drawn = [];
    const orig = window.G.text.bind(window.G);
    window.G.text = (s, x, cy, w, ...rest) => { drawn.push({ s, w }); return orig(s, x, cy, w, ...rest); };
    window.G.drawNow();
    const g = window.G;
    return {
      joined: g.mergeAt('new', 0, 0) !== undefined && g.mergeAt('new', 0, 0) === g.mergeAt('new', 0, 2),
      changedInside: g.mergeAt('new', 1, 2),
      title: drawn.find((x) => x.s === 'Q1 report'),
      widths: g.cols.widths.slice(0, 3),
    };
  });
  assert(r.joined, 'cells of the merge should share it');
  assert(r.changedInside === undefined, 'a changed cell is never drawn merged');
  const total = r.widths.reduce((a, b) => a + b, 0);
  assert(r.title && Math.abs(r.title.w - total) < 1, `title drawn over ${r.title?.w}px, merge is ${total}px`);
});

await test('a moved row is marked and says where it came from', async () => {
  const r = await page.evaluate(async () => {
    window.G?.cleanUp();
    const { parseCsv, diffTables, DiffGrid } = window.sd;
    const d = diffTables(parseCsv('id,v\n1,a\n2,b\n3,c\n4,d'), parseCsv('id,v\n1,a\n4,d\n2,b\n3,c'), { keys: ['id'] });
    window.G = new DiffGrid({ layout: 'unified', collapseUnchanged: false });
    window.G.render({ diff: d, container: document.getElementById('grid') });
    const marks = [];
    const ctx = window.G.ctx;
    const fill = ctx.fillText.bind(ctx);
    ctx.fillText = (s, ...rest) => { if (s === '↕') marks.push(s); return fill(s, ...rest); };
    window.G.drawNow();
    const row = Array.from({ length: d.rowCount }, (_, i) => i).find((i) => d.moved[i]);
    return { marks: marks.length, info: window.G.cellInfo(row, 1).moved, row, rowsMoved: d.stats.rowsMoved };
  });
  assert(r.rowsMoved === 1 && r.marks === 1 && r.info === true, JSON.stringify(r));
  // Hovering it shows where it moved from.
  const box = await page.evaluate((row) => {
    const g = window.G;
    const pane = g.panes()[0];
    const d = g.rows.displayOf[row];
    const s = g.scroller.getBoundingClientRect();
    return { x: s.x + pane.gutter + g.cols.widths[0] + 10, y: s.y + 32 + g.rows.offsets[d] + 10 };
  }, r.row);
  await page.mouse.move(box.x, box.y);
  await page.waitForFunction(() => window.G.tooltip.hasAttribute('data-open'), null, { timeout: 3000 }).catch(() => {});
  const tip = await page.evaluate(() => window.G.tooltip.textContent ?? '');
  assert(/Moved from row 5/.test(tip), `tooltip: ${tip}`);
  await page.mouse.move(1, 1);
});

await test('annotations keep their elements across equal arrays; data changes re-render in place', async () => {
  const r = await page.evaluate(() => {
    window.G?.cleanUp();
    const { parseCsv, diffTables, DiffGrid } = window.sd;
    const d = diffTables(parseCsv('id,v\n1,a\n2,b\n3,c'), parseCsv('id,v\n1,a\n2,x\n3,y'));
    let calls = 0;
    const renderAnnotation = (a) => { calls++; return Object.assign(document.createElement('div'), { textContent: String(a.data) }); };
    const g = new DiffGrid({ annotations: [{ row: 1, data: 'first' }, { row: 2, data: 'second' }], renderAnnotation });
    g.render({ diff: d, container: document.getElementById('grid') });
    const before = [...document.querySelectorAll('.sd-annotation')];
    g.setOptions({ annotations: [{ row: 1, data: 'first' }, { row: 2, data: 'second' }] }); // fresh, equal
    const same = [...document.querySelectorAll('.sd-annotation')].every((e, i) => e === before[i]);
    const callsAfterEqual = calls;
    g.setOptions({ annotations: [{ row: 1, data: 'first' }, { row: 2, data: 'changed' }] });
    const texts = [...document.querySelectorAll('.sd-annotation')].map((e) => e.textContent);
    const stillSame = [...document.querySelectorAll('.sd-annotation')].every((e, i) => e === before[i]);
    const callsAfterChange = calls;
    g.setOptions({ annotations: [{ row: 0, data: 'first', id: 'x' }] }); // id-keyed, moved row; the other removed
    const n = document.querySelectorAll('.sd-annotation').length;
    window.G = g;
    return { same, callsAfterEqual, calls: callsAfterChange, texts, stillSame, n };
  });
  assert(r.same && r.callsAfterEqual === 2, `equal array re-rendered: ${JSON.stringify(r)}`);
  assert(r.stillSame && r.calls === 3 && r.texts.join() === 'first,changed', `data change: ${JSON.stringify(r)}`);
  assert(r.n === 1, `removal: ${JSON.stringify(r)}`);
});

await test('the minimap shades by density instead of painting every pixel solid', async () => {
  const r = await page.evaluate(() => {
    window.G?.cleanUp();
    const { parseCsv, diffTables, DiffGrid } = window.sd;
    const mk = (every) => (bump) => 'id,v\n' + Array.from({ length: 20000 }, (_, i) => `${i},${bump && i % every === 0 ? 'x' : 'v'}${i}`).join('\n');
    const alphasFor = (every) => {
      const m = mk(every);
      const g = new DiffGrid({ collapseUnchanged: false, tooltips: false });
      g.render({ diff: diffTables(parseCsv(m(false)), parseCsv(m(true))), container: document.getElementById('grid') });
      g.drawNow();
      const a = Array.from(g.minimapCache.alphas).filter((x) => x > 0);
      g.cleanUp();
      return { max: Math.max(...a), min: Math.min(...a), painted: a.length };
    };
    return { sparse: alphasFor(10), dense: alphasFor(1), lone: alphasFor(100000) };
  });
  assert(r.dense.max === 1 && r.sparse.max < 0.6, `alphas: ${JSON.stringify(r)}`);
  assert(r.lone.painted >= 1 && r.lone.min >= 0.35, `a single change must stay visible: ${JSON.stringify(r)}`);
});

await test('screen readers get an ARIA grid of the visible cells, following the selection', async () => {
  const r = await page.evaluate(async () => {
    window.G?.cleanUp();
    const { parseCsv, diffTables, DiffGrid } = window.sd;
    const rows = (bump) => 'id,price,qty\n' + Array.from({ length: 400 }, (_, i) => `${i},${bump && i === 3 ? 12.5 : 10},${i}`).join('\n');
    const d = diffTables(parseCsv(rows(false)), parseCsv(rows(true)));
    const g = new DiffGrid({ collapseUnchanged: true });
    g.render({ diff: d, container: document.getElementById('grid') });
    g.drawNow();
    window.G = g;
    const root = g.scroller;
    const q = (sel) => [...root.querySelectorAll(sel)];
    g.select(3, 1);
    g.drawNow();
    const active = root.getAttribute('aria-activedescendant');
    const activeText = active ? g.shadow.getElementById(active)?.textContent : null;
    return {
      role: root.getAttribute('role'),
      label: root.getAttribute('aria-label'),
      rowcount: root.getAttribute('aria-rowcount'),
      headers: q('[role=columnheader]').map((e) => e.textContent),
      separators: q('[role=row] [role=gridcell]').map((e) => e.textContent).filter((t) => /unchanged/.test(t)).length,
      rowHeaders: q('[role=rowheader]').map((e) => e.textContent).slice(0, 6),
      activeText,
      hidden: getComputedStyle(root.querySelector('.sd-a11y')).clipPath,
    };
  });
  assert(r.role === 'grid' && r.label === 'Spreadsheet diff: 1 row modified, 1 cell changed' && Number(r.rowcount) > 3, JSON.stringify(r));
  assert(r.headers.join() === 'Row,id,price,qty', `headers: ${r.headers}`);
  assert(r.separators >= 1, 'collapsed rows are announced');
  assert(r.rowHeaders.includes('Row 5, modified'), `row headers: ${r.rowHeaders}`);
  assert(r.activeText === '12.5, was 10, up 2.5 (25.0%), edited', `active cell reads: ${r.activeText}`);
  assert(r.hidden && r.hidden !== 'none', 'mirror is visually hidden');
});

await test('axe-core finds no accessibility violations in the grid', async () => {
  await page.evaluate(() => {
    window.G?.cleanUp();
    const { parseCsv, diffTables, DiffGrid } = window.sd;
    const mk = (bump) => 'id,name,price\n' + Array.from({ length: 60 }, (_, i) => `${i},item ${i},${bump && i % 7 === 0 ? i + 0.5 : i}`).join('\n');
    window.G = new DiffGrid({ layout: 'split' });
    window.G.render({ diff: diffTables(parseCsv(mk(false)), parseCsv(mk(true))), container: document.getElementById('grid') });
    window.G.select(0, 2);
    window.G.drawNow();
  });
  const axe = readFileSync(createRequire(import.meta.url).resolve('axe-core/axe.min.js'), 'utf8');
  await page.addScriptTag({ content: axe });
  const res = await page.evaluate(async () => {
    const r = await window.axe.run(document.getElementById('grid'), { resultTypes: ['violations'] });
    return r.violations.map((v) => `${v.impact} ${v.id}: ${v.help} (${v.nodes.length})`);
  });
  assert(res.length === 0, res.join('\n    '));
});

await test('canvas text colors meet WCAG AA in light and dark', async () => {
  const fails = await page.evaluate(() => {
    const { parseCsv, diffTables, DiffGrid } = window.sd;
    // Composite a color over a background on a canvas and read back sRGB.
    const cv = document.createElement('canvas');
    cv.width = cv.height = 1;
    const cx = cv.getContext('2d', { willReadFrequently: true });
    const rgb = (color, over = '#fff') => {
      cx.globalAlpha = 1;
      cx.fillStyle = over;
      cx.fillRect(0, 0, 1, 1);
      cx.fillStyle = color;
      cx.fillRect(0, 0, 1, 1);
      return [...cx.getImageData(0, 0, 1, 1).data.slice(0, 3)];
    };
    const lum = ([r, g, b]) => [r, g, b].map((v) => { v /= 255; return v <= 0.03928 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4; }).reduce((a, v, i) => a + v * [0.2126, 0.7152, 0.0722][i], 0);
    const ratio = (a, b) => { const [x, y] = [lum(a), lum(b)].sort((p, q) => q - p); return (x + 0.05) / (y + 0.05); };
    const out = [];
    for (const theme of ['light', 'dark']) {
      window.G?.cleanUp();
      const g = new DiffGrid({ theme });
      g.render({ diff: diffTables(parseCsv('a\n1'), parseCsv('a\n2')), container: document.getElementById('grid') });
      g.drawNow();
      const t = g.theme;
      const bg = rgb(t.bg, t.bg);
      const on = (layer) => rgb(layer, t.bg);
      // [text, background it's drawn on, minimum]: 4.5 for text, 3 for glyphs.
      const pairs = [
        ['fg', 'bg'], ['fg', 'addedBg'], ['fg', 'removedBg'], ['fg', 'addedEmphasis'], ['fg', 'removedEmphasis'], ['fg', 'computedEmphasis'], ['fg', 'headerBg'],
        ['fgMuted', 'bg'], ['fgMuted', 'headerBg'], ['fgMuted', 'separatorBg'],
        ['fgFaint', 'gutterBg'], ['fgFaint', 'bg'],
        ['addedFg', 'bg'], ['addedFg', 'addedBg'], ['addedFg', 'computedEmphasis'],
        ['removedFg', 'bg'], ['removedFg', 'removedBg'], ['removedFg', 'computedEmphasis'],
        // Glyphs and labels in the status hues: the dot, ↕, +/− signs, change counts, tooltip labels.
        ['modifiedFg', 'gutterBg'], ['modifiedFg', 'bg'], ['addedFg', 'gutterBg'], ['removedFg', 'gutterBg'], ['computedFg', 'bg'], ['computedFg', 'computedEmphasis'],
        // Outlines (graphics: 3:1). Bars and tints are backed up by signs, so they're not checked.
        ['computedFg', 'bg', 3], ['modifiedFg', 'bg', 3],
        // The selected cell's ring and the keyboard focus ring (WCAG 1.4.11).
        ['selection', 'bg', 3], ['selection', 'addedEmphasis', 3], ['selection', 'removedEmphasis', 3], ['selection', 'computedEmphasis', 3],
      ];
      for (const [fgKey, bgKey, min = 4.5] of pairs) {
        const back = bgKey === 'bg' ? bg : on(t[bgKey]);
        const r = ratio(rgb(t[fgKey], `rgb(${back.join(',')})`), back);
        if (r < min) out.push(`${theme}: ${fgKey} on ${bgKey} ${r.toFixed(2)} < ${min}`);
      }
      g.cleanUp();
    }
    return out;
  });
  assert(fails.length === 0, fails.join('\n    '));
});

await test('keyboard focus is visible: a ring when tabbed into, none after a click', async () => {
  await setup({}, 50);
  const ring = () => page.evaluate(() => {
    const r = window.G.shadow.querySelector('.sd-focus-ring');
    return { opacity: getComputedStyle(r).opacity, afterCanvas: !!(window.G.canvas.compareDocumentPosition(r) & Node.DOCUMENT_POSITION_FOLLOWING) };
  });
  assert((await ring()).opacity === '0', 'ring shows before focus');
  await page.focus('#before');
  await page.keyboard.press('Tab');
  const tabbed = await ring();
  assert(tabbed.opacity === '1' && tabbed.afterCanvas, `tabbed in: ${JSON.stringify(tabbed)}`);
  // Focus that arrives by mouse gets no ring (the selection shows where it is).
  const before = await page.locator('#before').boundingBox();
  await page.mouse.click(before.x + 5, before.y + 5);
  const box = await scrollerBox();
  await page.mouse.click(box.x + 200, box.y + 80);
  const clicked = await page.evaluate(() => window.G.shadow.activeElement === window.G.scroller);
  assert(clicked && (await ring()).opacity === '0', 'ring shows after a mouse click');
});

await test('Home/End, Ctrl+Home/End and PageUp/PageDown move the selection; Shift extends', async () => {
  await setup({}, 300);
  await page.evaluate(() => window.G.scroller.focus());
  const sel = () => page.evaluate(() => {
    const g = window.G;
    const s = g.selected;
    return { row: s.row, vi: g.cols.cols.indexOf(s.col), last: g.cols.cols.length - 1, rows: g.diff.rowCount, top: g.scroller.scrollTop, extent: g.extent };
  });
  await page.keyboard.press('ArrowDown');
  await page.keyboard.press('End');
  let s = await sel();
  assert(s.vi === s.last && s.row === 0, `End: ${JSON.stringify(s)}`);
  await page.keyboard.press('Home');
  s = await sel();
  assert(s.vi === 0, `Home: ${JSON.stringify(s)}`);
  await page.keyboard.press('PageDown');
  const paged = await sel();
  assert(paged.row > 5 && paged.row < 40 && paged.top > 0, `PageDown: ${JSON.stringify(paged)}`);
  await page.keyboard.press('PageUp');
  s = await sel();
  assert(s.row === 0, `PageUp: ${JSON.stringify(s)}`);
  await page.keyboard.press('Control+End');
  s = await sel();
  assert(s.row === s.rows - 1 && s.vi === s.last, `Ctrl+End: ${JSON.stringify(s)}`);
  await page.keyboard.press('Control+Home');
  s = await sel();
  assert(s.row === 0 && s.vi === 0 && s.top === 0, `Ctrl+Home: ${JSON.stringify(s)}`);
  await page.keyboard.press('Shift+End');
  s = await sel();
  assert(s.row === 0 && s.vi === 0 && s.extent && s.extent.row === 0, `Shift+End extends: ${JSON.stringify(s)}`);
});

await test('forced colors (Windows High Contrast): the canvas uses the system colors', async () => {
  await setup({}, 20);
  const read = () => page.evaluate(() => {
    const t = window.G.theme;
    return { bg: t.bg, fg: t.fg, addedEmphasis: t.addedEmphasis, selection: t.selection };
  });
  const normal = await read();
  await page.emulateMedia({ forcedColors: 'active' });
  await page.waitForTimeout(100);
  const forced = await read();
  // A fresh grid resolves the same way as one that followed the change.
  await setup({}, 20);
  const fresh = await read();
  await page.emulateMedia({ forcedColors: 'none' });
  await page.waitForTimeout(100);
  const back = await read();
  // Every token used to read as the one forced text color: a solid block.
  assert(forced.bg !== forced.fg, `background and text are the same color: ${JSON.stringify(forced)}`);
  assert(forced.addedEmphasis === forced.bg, `tints should fall back to the background: ${JSON.stringify(forced)}`);
  assert(JSON.stringify(forced) !== JSON.stringify(normal), 'the theme did not change under forced colors');
  assert(JSON.stringify(fresh) === JSON.stringify(forced), `fresh ${JSON.stringify(fresh)} != followed ${JSON.stringify(forced)}`);
  assert(JSON.stringify(back) === JSON.stringify(normal), 'the theme did not return when forced colors ended');
});

await test('strings and locale translate what the grid announces and draws', async () => {
  const r = await page.evaluate(() => {
    window.G?.cleanUp();
    const { parseCsv, diffTables, DiffGrid } = window.sd;
    const rows = (bump) => 'id,price\n' + Array.from({ length: 1500 }, (_, i) => `${i},${bump && i === 1499 ? 12.5 : 10}`).join('\n');
    const strings = {
      gridSummary: (parts) => `Tabellenvergleich: ${parts.join('; ')}`,
      summaryRowsModified: (c, n) => `${c} ${n === 1 ? 'Zeile' : 'Zeilen'} geändert`,
      summaryCellsChanged: (c, n) => `${c} ${n === 1 ? 'Zelle' : 'Zellen'} geändert`,
      rowHeaderColumn: 'Zeile',
      unchangedRowsCollapsed: (c) => `${c} unveränderte Zeilen`,
      rowHeader: (row) => `Zeile ${row}`,
      rowModified: 'geändert',
      was: (v) => `vorher ${v}`,
      up: (a) => `plus ${a}`,
      kindInput: 'bearbeitet',
    };
    const g = new DiffGrid({ collapseUnchanged: true, strings, locale: 'de-DE' });
    g.render({ diff: diffTables(parseCsv(rows(false)), parseCsv(rows(true))), container: document.getElementById('grid') });
    g.select(1499, 1);
    g.drawNow();
    window.G = g;
    const root = g.scroller;
    const cells = [...root.querySelectorAll('[role=gridcell]')].map((e) => e.textContent);
    const active = g.shadow.getElementById(root.getAttribute('aria-activedescendant'))?.textContent;
    const out = { label: root.getAttribute('aria-label'), header: root.querySelector('[role=columnheader]')?.textContent, separator: cells.find((t) => /unveränderte/.test(t)), active, rowHeaders: [...root.querySelectorAll('[role=rowheader]')].map((e) => e.textContent) };
    // Switching back to English rebuilds the words.
    g.setOptions({ strings: undefined, locale: 'en-US' });
    g.drawNow();
    out.english = root.getAttribute('aria-label');
    return out;
  });
  assert(r.label === 'Tabellenvergleich: 1 Zeile geändert; 1 Zelle geändert', `label: ${r.label}`);
  assert(r.header === 'Zeile', `header: ${r.header}`);
  assert(/^1\.4\d\d unveränderte Zeilen$/.test(r.separator ?? ''), `separator (de-DE grouping): ${r.separator}`);
  assert(r.active === '12.5, vorher 10, plus 2.5 (25.0%), bearbeitet', `active cell: ${r.active}`);
  assert(r.rowHeaders.includes('Zeile 1501, geändert'), `row headers: ${r.rowHeaders}`);
  assert(r.english === 'Spreadsheet diff: 1 row modified, 1 cell changed', `back to English: ${r.english}`);
});

await test('axe-core finds no violations in spreadiff/html HTML, light and dark', async () => {
  const axe = readFileSync(createRequire(import.meta.url).resolve('axe-core/axe.min.js'), 'utf8');
  await page.addScriptTag({ content: axe });
  const out = [];
  for (const scheme of ['light', 'dark']) {
    await page.emulateMedia({ colorScheme: scheme });
    const v = await page.evaluate(async () => {
      window.G?.cleanUp();
      const { parseCsv, diffTables, renderHTML, HTML_CSS } = window.sd;
      const mk = (b) => 'id,name,price,qty\n' + Array.from({ length: 30 }, (_, i) => `${i},item ${b && i === 4 ? 'renamed' : i},${b && i % 5 === 0 ? i + 0.5 : i},${i}`).join('\n') + (b ? '\n99,new,1,1' : '');
      const d = diffTables(parseCsv(mk(false)), parseCsv(mk(true)), { keys: ['id'] });
      const host = document.getElementById('grid');
      host.innerHTML = `<style>${HTML_CSS}</style>` + renderHTML(d, { layout: 'unified', display: 'delta' }) + renderHTML(d, { layout: 'split' });
      const r = await window.axe.run(host, { resultTypes: ['violations'] });
      host.innerHTML = '';
      return r.violations.map((x) => `${x.impact} ${x.id}: ${x.nodes.slice(0, 2).map((n) => n.any?.[0]?.message ?? n.failureSummary).join(' | ')}`);
    });
    out.push(...v.map((x) => `${scheme}: ${x}`));
  }
  await page.emulateMedia({ colorScheme: 'light' });
  assert(out.length === 0, out.join('\n    '));
});

await test('the stylesheet is adopted, not inline (works under a strict CSP)', async () => {
  const r = await page.evaluate(() => {
    window.G?.cleanUp();
    const { DiffGrid } = window.sd;
    const g = new DiffGrid();
    window.G = g;
    return { adopted: g.shadow.adoptedStyleSheets.length, inline: g.shadow.querySelectorAll('style').length };
  });
  assert(r.adopted === 1 && r.inline === 0, JSON.stringify(r));
});

await test('without light-dark() support, the fallback gives the same colors', async () => {
  const run = async (scheme, theme) => {
    await page.emulateMedia({ colorScheme: scheme });
    return page.evaluate((theme) => {
      window.G?.cleanUp();
      const { parseCsv, diffTables, DiffGrid } = window.sd;
      const g = new DiffGrid({ theme });
      g.render({ diff: diffTables(parseCsv('a\n1'), parseCsv('a\n2')), container: document.getElementById('grid') });
      const normal = { ...g.theme };
      // Pretend light-dark() is missing, as an old browser sees it: every use of it becomes an
      // unknown function, so the main tokens are invalid and the @supports-not fallback applies.
      const sheet = g.shadow.adoptedStyleSheets[0];
      const css = [...sheet.cssRules].map((r) => r.cssText).join('\n');
      sheet.replaceSync(css.replace(/light-dark\(/g, 'no-light-dark('));
      const broken = getComputedStyle(g.host).getPropertyValue('--sdi-added-bg');
      g.refreshTheme();
      const fallback = { ...g.theme };
      window.G = g;
      const out = Object.keys(normal).filter((k) => typeof normal[k] === 'string' && normal[k] !== fallback[k]).map((k) => `${k}: ${normal[k]} vs ${fallback[k]}`);
      if (/no-light-dark/.test(broken)) out.push('fallback did not apply');
      return out;
    }, theme);
  };
  const diffs = [...(await run('light', 'system')).map((x) => `light ${x}`), ...(await run('dark', 'system')).map((x) => `system dark ${x}`), ...(await run('light', 'dark')).map((x) => `forced dark ${x}`)];
  await page.emulateMedia({ colorScheme: 'light' });
  assert(diffs.length === 0, diffs.join('\n    '));
});

await test('accessibility: false skips the mirror', async () => {
  const n = await page.evaluate(() => {
    window.G?.cleanUp();
    const { parseCsv, diffTables, DiffGrid } = window.sd;
    const g = new DiffGrid({ accessibility: false });
    g.render({ diff: diffTables(parseCsv('a\n1'), parseCsv('a\n2')), container: document.getElementById('grid') });
    g.drawNow();
    window.G = g;
    return g.scroller.querySelectorAll('[role=gridcell]').length;
  });
  assert(n === 0, `cells: ${n}`);
});

// ---- Row and column labels ------------------------------------------------
const MODEL = (g, c) => [',FY2024,FY2025', 'India,,', '  Revenue,8300,' + 8300 * (1 + g), '  Costs,4150,4565', 'China,,', '  Revenue,' + c + ',' + c * 2, '  Costs,360,396'].join('\n');
async function labelled(opts = {}) {
  return page.evaluate(({ opts, a, b }) => {
    window.G?.cleanUp();
    const { parseCsv, diffTables, DiffGrid } = window.sd;
    const d = diffTables(parseCsv(a, { header: false }), parseCsv(b, { header: false }), { keys: 'none' });
    const g = new DiffGrid({ collapseUnchanged: false, ...opts });
    g.render({ diff: d, container: document.getElementById('grid') });
    const drawn = [];
    const ctx = g.ctx;
    const fill = ctx.fillText.bind(ctx);
    ctx.fillText = (s, ...rest) => { drawn.push(s); return fill(s, ...rest); };
    g.drawNow();
    ctx.fillText = fill;
    window.G = g;
    return {
      drawn,
      columns: Array.from(g.cols.cols, (ci) => d.columns[ci].name),
      rows: g.rows.count,
      gutter: g.panes()[0].gutter,
      rowHeaders: [...g.scroller.querySelectorAll('[role=rowheader]')].map((e) => e.textContent),
      colHeaders: [...g.scroller.querySelectorAll('[role=columnheader]')].map((e) => e.textContent),
    };
  }, { opts, a: MODEL(0.1, 720), b: MODEL(0.12, 800) });
}

await test('row headers show inferred labels; the label column and the year row fold into the headers', async () => {
  const r = await labelled();
  assert(r.drawn.includes('India › ') && r.drawn.some((s) => s === 'Revenue' || s.startsWith('Revenue')), `drawn: ${r.drawn.slice(0, 40)}`);
  assert(r.drawn.includes('FY2025') && !r.drawn.includes('C'), 'the column header shows the year, not the letter');
  assert(r.columns.join() === 'B,C', `the label column A is folded: ${r.columns}`);
  assert(r.rows === 6, `the unchanged year row is folded: ${r.rows} rows`);
  assert(r.rowHeaders.some((t) => t.startsWith('India › Revenue, row 3')), `aria: ${r.rowHeaders}`);
  assert(r.colHeaders.includes('FY2025'), `aria headers: ${r.colHeaders}`);
});

await test("rowHeaders: 'numbers' and columnHeaders: 'letters' show the sheet as it is", async () => {
  const r = await labelled({ rowHeaders: 'numbers', columnHeaders: 'letters' });
  assert(r.columns.join() === 'A,B,C' && r.rows === 7, JSON.stringify({ c: r.columns, rows: r.rows }));
  assert(r.drawn.includes('C') && !r.drawn.includes('India › '), 'letters and no labels');
  const both = await labelled({ rowHeaders: 'both', columnHeaders: 'both' });
  assert(both.drawn.includes('India › ') && both.drawn.includes('3') && both.drawn.includes('C') && both.drawn.includes('FY2025'), `both: ${both.drawn.slice(0, 60)}`);
  const none = await labelled({ rowHeaders: 'none' });
  assert(none.gutter < 30 && !none.drawn.includes('India › '), `none: gutter ${none.gutter}`);
  const kept = await labelled({ foldLabels: false });
  assert(kept.columns.join() === 'A,B,C' && kept.rows === 7, 'foldLabels: false keeps the label column and the year row');
});

await test('a label that does not fit drops outer sections, and shows in full on hover', async () => {
  await labelled({ rowHeaderWidth: 70 });
  const r = await page.evaluate(() => {
    const g = window.G;
    const keys = [...g.truncatedLabels];
    const d = g.rows.displayOf[2];
    const s = g.scroller.getBoundingClientRect();
    return { keys, x: s.x + 20, y: s.y + 36 + g.rows.offsets[d] + 8 };
  });
  assert(r.keys.length > 0, 'some labels are cut short at 70px');
  await page.mouse.move(r.x, r.y);
  await page.waitForFunction(() => window.G.tooltip.hasAttribute('data-open'), null, { timeout: 3000 }).catch(() => {});
  const tip = await page.evaluate(() => window.G.tooltip.textContent ?? '');
  assert(/India › /.test(tip) && /row/.test(tip), `tooltip: ${tip}`);
  await page.mouse.move(1, 1);
});

await test('labels for a large table come after the first frame; numbers until then', async () => {
  const r = await page.evaluate(async () => {
    window.G?.cleanUp();
    const { parseCsv, diffTables, DiffGrid } = window.sd;
    const mk = (b) => 'sku,name,qty\n' + Array.from({ length: 120000 }, (_, i) => `S${i},Item ${i},${b && i % 1000 === 0 ? 9 : i % 50}`).join('\n');
    const d = diffTables(parseCsv(mk(false)), parseCsv(mk(true)));
    const g = new DiffGrid();
    const t0 = performance.now();
    g.render({ diff: d, container: document.getElementById('grid') });
    g.drawNow();
    const first = { ms: performance.now() - t0, mode: g.rowHeaderMode() };
    window.G = g;
    await new Promise((res) => setTimeout(res, 400));
    return { first, later: g.rowHeaderMode(), columns: Array.from(g.cols.cols, (ci) => d.columns[ci].name) };
  });
  assert(r.first.mode === 'numbers', `first frame: ${JSON.stringify(r.first)}`);
  assert(r.later === 'labels' && r.columns.join() === 'qty', `after: ${JSON.stringify(r)}`);
});

// A screen point inside the cell (aligned row, column index), found through the grid's own hit testing.
const pointOf = (row, col) => page.evaluate(({ row, col }) => {
  const g = window.G;
  const r = g.scroller.getBoundingClientRect();
  for (let y = r.top + 40; y < r.bottom - 4; y += 6) {
    for (let x = r.left + 4; x < r.right - 20; x += 12) {
      const c = g.cellAt(g.hitTest(x, y));
      if (c && c.row === row && c.columnIndex === col) return { x: x + 2, y: y + 2 };
    }
  }
  return null;
}, { row, col });

await test('drag, shift-click and shift+arrows select a range; onRangeChange reports it once per change', async () => {
  await page.evaluate(() => {
    window.G?.cleanUp();
    const { parseCsv, diffTables, DiffGrid } = window.sd;
    const mk = (v) => 'id,a,b,c\n' + Array.from({ length: 30 }, (_, i) => `${i},${v}${i},${i},${i}`).join('\n');
    window.diff = diffTables(parseCsv(mk('a')), parseCsv(mk('b')));
    window.ranges = [];
    window.G = new DiffGrid({ collapseUnchanged: false, tooltips: false, onRangeChange: (r) => window.ranges.push(r && { rows: r.rows, columns: r.columns, a1: r.a1 }) });
    window.G.render({ diff: window.diff, container: document.getElementById('grid') });
  });
  await page.waitForTimeout(150);
  const a = await pointOf(2, 1);
  const b = await pointOf(4, 2);
  assert(a && b, 'cells not found on screen');
  await page.mouse.move(a.x, a.y);
  await page.mouse.down();
  await page.mouse.move((a.x + b.x) / 2, (a.y + b.y) / 2);
  await page.mouse.move(b.x, b.y);
  await page.mouse.move(b.x, b.y + 1); // same cell: no new event
  await page.mouse.up();
  let r = await page.evaluate(() => window.G.getRange());
  assert(r.rows.join() === '2,3,4' && r.columns.join() === '1,2', `drag: ${JSON.stringify(r)}`);
  assert(r.a1.new === 'B4:C6' && r.a1.old === 'B4:C6', `a1: ${JSON.stringify(r.a1)}`);
  const n = await page.evaluate(() => window.ranges.length);
  const last = await page.evaluate(() => window.ranges[window.ranges.length - 1]);
  assert(last.rows.join() === '2,3,4', `last event ${JSON.stringify(last)}`);
  assert(n <= 4, `${n} events for one drag`);

  // Shift-click moves the far corner, keeps the anchor.
  const c = await pointOf(2, 3);
  await page.keyboard.down('Shift');
  await page.mouse.click(c.x, c.y);
  await page.keyboard.up('Shift');
  r = await page.evaluate(() => window.G.getRange());
  assert(r.anchor.row === 2 && r.anchor.col === 1 && r.rows.join() === '2' && r.columns.join() === '1,2,3', `shift-click: ${JSON.stringify(r)}`);

  // Shift+Down extends; a plain arrow collapses to one cell.
  await page.keyboard.press('Shift+ArrowDown');
  r = await page.evaluate(() => window.G.getRange());
  assert(r.rows.join() === '2,3' && r.columns.join() === '1,2,3', `shift+down: ${JSON.stringify(r)}`);
  await page.keyboard.press('ArrowDown');
  r = await page.evaluate(() => window.G.getRange());
  assert(r.rows.join() === '3' && r.columns.join() === '1' && r.anchor.row === 3, `arrow: ${JSON.stringify(r)}`);

  // selectRange, and Escape clears.
  const ok = await page.evaluate(() => window.G.selectRange(5, 0, 7, 1));
  r = await page.evaluate(() => window.G.getRange());
  assert(ok && r.a1.new === 'A7:B9', `selectRange: ${JSON.stringify(r)}`);
  await page.keyboard.press('Escape');
  assert((await page.evaluate(() => window.G.getRange())) === null, 'Escape should clear');
  assert((await page.evaluate(() => window.ranges[window.ranges.length - 1])) === null, 'last event should be null');
});

await test('a range in A1 per side: an added row is skipped on the old side, moved rows scatter it', async () => {
  await page.evaluate(() => {
    window.G?.cleanUp();
    const { parseCsv, diffTables, DiffGrid } = window.sd;
    window.diff = diffTables(parseCsv('id,a\n1,x\n3,z'), parseCsv('id,a\n1,x\n2,y\n3,z'));
    window.G = new DiffGrid({ collapseUnchanged: false, tooltips: false });
    window.G.render({ diff: window.diff, container: document.getElementById('grid') });
  });
  await page.waitForTimeout(100);
  const r = await page.evaluate(() => (window.G.selectRange(0, 1, 2, 1), window.G.getRange()));
  assert(r.a1.new === 'B2:B4' && r.a1.old === 'B2:B3', JSON.stringify(r.a1));
  const one = await page.evaluate(() => (window.G.selectRange(1, 1, 1, 1), window.G.getRange()));
  assert(one.a1.new === 'B3' && one.a1.old === undefined, JSON.stringify(one.a1));
  // Rows matched by key in a new order: old rows 3,1 aren't one block in order.
  const moved = await page.evaluate(() => {
    const { parseCsv, diffTables } = window.sd;
    window.G.setDiff(diffTables(parseCsv('id,a\n1,x\n2,y\n3,z'), parseCsv('id,a\n3,z\n2,y\n1,x')));
    window.G.selectRange(0, 1, 1, 1);
    return window.G.getRange().a1;
  });
  assert(moved.new === 'B2:B3' && moved.old === undefined, JSON.stringify(moved));
});

await test('a sheet view brings its trace, and the grid finds the traced sheet from the diff', async () => {
  const r = await page.evaluate(() => {
    window.G?.cleanUp();
    const { compareWorkbooks, DiffGrid } = window.sd;
    const sheet = (name, rows, formulas) => ({ name, columns: ['A', 'B'], rows, formulas, origin: { row: 0, col: 0 } });
    // Sheet 2 reads sheet 1: an edit on S1 moves S2!B1.
    const a = { name: 'a', sheets: [sheet('S1', [[1, 2]], [[null, null]]), sheet('S2', [[0, 3]], [[null, '=S1!A1+2']])] };
    const b = { name: 'b', sheets: [sheet('S1', [[5, 2]], [[null, null]]), sheet('S2', [[0, 7]], [[null, '=S1!A1+2']])] };
    const cmp = compareWorkbooks(a, b);
    window.G = new DiffGrid({ tooltips: false });
    const container = document.getElementById('grid');
    window.G.render({ sheet: cmp.sheets[1], container });
    const kindOnS2 = window.G.cellInfo(0, 1).kind;
    // A wrong index next to the right diff: the diff decides.
    window.G.setOptions({ trace: { ...cmp.sheets[1].trace, sheet: 0 } });
    const stillS2 = window.G.cellInfo(0, 1).kind;
    window.G.render({ sheet: cmp.sheets[0], container });
    const kindOnS1 = window.G.cellInfo(0, 0).kind;
    return { kindOnS2, stillS2, kindOnS1 };
  });
  assert(r.kindOnS2 === 'computed' && r.stillS2 === 'computed' && r.kindOnS1 === 'input', JSON.stringify(r));
});

// ---- SpreadsheetDiff --------------------------------------------------------
const fixtureBytes = (name) => Array.from(readFileSync(new globalThis.URL(`../test/hosted/fixtures/commission/${name}`, import.meta.url)));
await test('SpreadsheetDiff: two files in, tabs, the traced grid and the outputs panel out', async () => {
  const r = await page.evaluate(async ({ a, b }) => {
    window.G?.cleanUp();
    window.W?.cleanUp();
    const { SpreadsheetDiff } = window.sd;
    const container = document.getElementById('grid');
    const sheets = [];
    window.W = new SpreadsheetDiff({ onSheetChange: (s) => sheets.push(s.name), tooltips: false });
    await window.W.render({ oldFile: { name: 'old.xlsx', data: new Uint8Array(a) }, newFile: { name: 'new.xlsx', data: new Uint8Array(b) }, container });
    const root = window.W.host.shadowRoot;
    const tabs = [...root.querySelectorAll('.sx-tab')].map((t) => [t.textContent, t.getAttribute('aria-selected')]);
    const outputs = root.querySelectorAll('.sx-out').length;
    const sideShown = !root.querySelector('.sx-side').hidden;
    const traced = !!window.W.grid.getOptions().trace;
    // Clicking an output reveals its cell (possibly on another sheet) and opens its waterfall.
    root.querySelector('.sx-out').click();
    await new Promise((res) => setTimeout(res, 50));
    const wf = root.querySelectorAll('.sx-wf-bar').length;
    const label = root.querySelector('.sx-wf')?.getAttribute('aria-label');
    const sel = window.W.grid.getRange();
    root.querySelectorAll('.sx-tab')[1].click();
    return { tabs, outputs, sideShown, traced, wf, label, sel: !!sel, sheet: window.W.sheet, sheets };
  }, { a: fixtureBytes('old.xlsx'), b: fixtureBytes('new.xlsx') });
  assert(r.tabs.length === 2 && r.tabs[0][0].startsWith('Reps') && r.tabs[0][1] === 'true', `tabs ${JSON.stringify(r.tabs)}`);
  assert(r.sideShown && r.outputs > 0 && r.traced, `outputs ${JSON.stringify(r)}`);
  assert(r.wf >= 3 && /Before .* After /.test(r.label ?? ''), `waterfall ${r.wf} ${r.label}`);
  assert(r.sel && r.sheet === 1 && r.sheets.includes('Rates'), `navigation ${JSON.stringify(r)}`);
});

await test('SpreadsheetDiff review: reject a selection, see it marked, download the merged file', async () => {
  const r = await page.evaluate(async ({ a, b }) => {
    window.W?.cleanUp();
    const { SpreadsheetDiff, parseXlsx } = window.sd;
    const container = document.getElementById('grid');
    let decided = 0;
    window.W = new SpreadsheetDiff({ review: true, tooltips: false, onDecisionsChange: () => decided++ });
    await window.W.render({ oldFile: { name: 'old.xlsx', data: new Uint8Array(a) }, newFile: { name: 'new.xlsx', data: new Uint8Array(b) }, container });
    const root = window.W.host.shadowRoot;
    const g = window.W.grid;
    const reject = () => [...root.querySelectorAll('.sx-btn')].find((x) => x.dataset.kind === 'reject');
    const disabledBefore = reject().disabled;
    // Select an edited input (not a downstream result), reject it with the button.
    const d0 = window.W.comparison.sheets[window.W.sheet].diff;
    let at = null;
    for (const [row, mask] of d0.changedCells) {
      const col = mask.findIndex((x, c) => x && g.cellInfo(row, c).kind === 'input');
      if (col >= 0) { at = { row, col }; break; }
    }
    g.select(at.row, at.col);
    const range = g.getRange();
    const enabled = !reject().disabled;
    reject().click();
    const info = g.cellInfo(range.anchor.row, range.anchor.col);
    const tally = root.querySelector('.sx-tally').textContent;
    const { file, applied, recalculated, skipped } = window.W.exportXlsx();
    // A whole-sheet selection decides edits only: downstream results follow them.
    g.selectRange(0, 0, d0.rowCount - 1, d0.columns.length - 1);
    window.W.decideSelection('reject');
    const all = root.querySelector('.sx-tally').textContent;
    window.W.decideSelection(undefined);
    g.select(at.row, at.col);
    window.W.decideSelection('reject');
    const merged = parseXlsx(file).sheets[window.W.sheet];
    const d = window.W.comparison.sheets[window.W.sheet].diff;
    const c = d.columns[range.anchor.col];
    const oldValue = d.oldTable.rows[d.oldRows[range.anchor.row]][c.oldIndex];
    const mergedValue = merged.rows[d.newRows[range.anchor.row]][c.newIndex];
    // 'a' in the grid accepts the selection instead.
    g.scroller.focus();
    g.scroller.dispatchEvent(new KeyboardEvent('keydown', { key: 'a', bubbles: true, composed: true }));
    const after = g.cellInfo(range.anchor.row, range.anchor.col).decision;
    return { disabledBefore, enabled, decision: info.decision, tally, all, applied, recalculated, skipped: skipped.length, oldValue, mergedValue, decided, after, download: [...root.querySelectorAll('.sx-btn')].some((x) => x.textContent === 'Download .xlsx') };
  }, { a: fixtureBytes('old.xlsx'), b: fixtureBytes('new.xlsx') });
  assert(r.disabledBefore && r.enabled, `buttons follow the selection ${JSON.stringify(r)}`);
  assert(r.decision === 'reject' && /^1 rejected/.test(r.tally), `marked ${JSON.stringify(r)}`);
  assert(r.applied === 1 && r.skipped === 0 && r.mergedValue === r.oldValue && r.recalculated > 0, `merged file ${JSON.stringify(r)}`);
  const [, rejected, , changes] = /^(\d+) rejected · (\d+) accepted · (\d+) changes/.exec(r.all) ?? [];
  assert(rejected === changes && Number(changes) < 60, `only edits are decided: ${r.all}`);
  assert(r.after === 'accept' && r.decided === 5 && r.download, `keyboard ${JSON.stringify(r)}`);
});

await test('SpreadsheetDiff (React): inline URL props that re-render on load compare once', async () => {
  const r = await page.evaluate(async ({ a, b }) => {
    window.G?.cleanUp();
    window.W?.cleanUp();
    const { SpreadsheetDiff, createElement: h, useState, createRoot } = window.sdReact;
    const blob = (bytes) => URL.createObjectURL(new Blob([new Uint8Array(bytes)]));
    const [oldHref, newHref] = [blob(a), blob(b)];
    let loads = 0;
    // The parent re-renders on every load, with new URL objects each time: the files must not be compared again.
    const App = () => {
      const [, setN] = useState(0);
      return h(SpreadsheetDiff, { oldFile: new URL(oldHref), newFile: new URL(newHref), style: { height: 400 }, onLoad: () => setN(++loads) });
    };
    const container = document.getElementById('grid');
    const root = createRoot(container);
    root.render(h(App));
    await new Promise((res) => setTimeout(res, 1500));
    const tabs = container.firstElementChild?.firstElementChild?.shadowRoot?.querySelectorAll('.sx-tab').length ?? 0;
    root.unmount();
    return { loads, tabs };
  }, { a: fixtureBytes('old.xlsx'), b: fixtureBytes('new.xlsx') });
  assert(r.loads === 1 && r.tabs === 2, `compared ${r.loads} times, ${r.tabs} tabs`);
});

await test('SpreadsheetDiff: an unreadable file shows the error and rejects', async () => {
  const r = await page.evaluate(async () => {
    window.W?.cleanUp();
    const { SpreadsheetDiff } = window.sd;
    let called = '';
    window.W = new SpreadsheetDiff({ onError: (e) => (called = e.message) });
    let threw = '';
    try {
      await window.W.render({ oldFile: { name: 'a.xlsx', data: new Uint8Array([80, 75, 3, 4, 0]) }, newFile: { name: 'b.xlsx', data: new Uint8Array([1]) }, container: document.getElementById('grid') });
    } catch (e) {
      threw = e.message;
    }
    return { threw, called, shown: window.W.host.shadowRoot.querySelector('.sx-msg').textContent };
  });
  assert(r.threw && r.called === r.threw && r.shown === r.threw, JSON.stringify(r));
  await page.evaluate(() => window.W.cleanUp());
});

await test('SpreadsheetDiff: strings translate the tabs, the outputs panel and the review bar', async () => {
  const r = await page.evaluate(async ({ a, b }) => {
    window.G?.cleanUp();
    window.W?.cleanUp();
    const { SpreadsheetDiff } = window.sd;
    const strings = { sheetsLabel: 'Blätter', outputsHeading: (c) => `Geänderte Ergebnisse · ${c}`, accept: 'Annehmen', waterfallBefore: 'Vorher', waterfallAfter: 'Nachher' };
    window.W = new SpreadsheetDiff({ review: true, tooltips: false, strings });
    await window.W.render({ oldFile: { name: 'old.xlsx', data: new Uint8Array(a) }, newFile: { name: 'new.xlsx', data: new Uint8Array(b) }, container: document.getElementById('grid') });
    const root = window.W.host.shadowRoot;
    root.querySelector('.sx-out').click();
    await new Promise((res) => setTimeout(res, 50));
    const out = {
      tabs: root.querySelector('.sx-tabs').getAttribute('aria-label'),
      heading: root.querySelector('.sx-side h2').textContent,
      buttons: [...root.querySelectorAll('.sx-btn')].map((x) => x.textContent),
      waterfall: root.querySelector('.sx-wf')?.getAttribute('aria-label'),
    };
    window.W.setOptions({ strings: undefined });
    out.englishTabs = root.querySelector('.sx-tabs').getAttribute('aria-label');
    out.englishButtons = [...root.querySelectorAll('.sx-btn')].map((x) => x.textContent);
    return out;
  }, { a: fixtureBytes('old.xlsx'), b: fixtureBytes('new.xlsx') });
  assert(r.tabs === 'Blätter' && /^Geänderte Ergebnisse · \d/.test(r.heading), JSON.stringify(r));
  assert(r.buttons.includes('Annehmen') && r.buttons.includes('Reject'), `buttons: ${r.buttons}`);
  assert(/^Vorher .* Nachher /.test(r.waterfall ?? ''), `waterfall: ${r.waterfall}`);
  assert(r.englishTabs === 'Sheets' && r.englishButtons.includes('Accept'), `back to English: ${JSON.stringify(r)}`);
});

await test('SpreadsheetDiff tabs: Home and End go to the first and last sheet', async () => {
  const r = await page.evaluate(async () => {
    const root = window.W.host.shadowRoot;
    const tabs = () => [...root.querySelectorAll('.sx-tab')];
    tabs()[0].focus();
    tabs()[0].dispatchEvent(new KeyboardEvent('keydown', { key: 'End', bubbles: true }));
    const end = { sheet: window.W.sheet, focused: tabs().indexOf(root.activeElement) };
    tabs()[end.focused].dispatchEvent(new KeyboardEvent('keydown', { key: 'Home', bubbles: true }));
    return { end, home: { sheet: window.W.sheet, focused: tabs().indexOf(root.activeElement) }, count: tabs().length };
  });
  assert(r.end.sheet === r.count - 1 && r.end.focused === r.count - 1, `End: ${JSON.stringify(r)}`);
  assert(r.home.sheet === 0 && r.home.focused === 0, `Home: ${JSON.stringify(r)}`);
});

await test('SpreadsheetDiff decisions: applied on load, replaced when the app passes new ones, no echo', async () => {
  const r = await page.evaluate(async ({ a, b }) => {
    window.W?.cleanUp();
    const { SpreadsheetDiff, decideRange, noDecisions, compareWorkbooks, parseXlsx } = window.sd;
    const cmp = compareWorkbooks(parseXlsx(new Uint8Array(a).buffer), parseXlsx(new Uint8Array(b).buffer));
    const sheet = cmp.firstChanged;
    const d = cmp.sheets[sheet].diff;
    const rejectAll = (n) => cmp.sheets.map((s, i) => (i === sheet ? decideRange(s.diff, noDecisions(), Array.from({ length: n }, (_, k) => k), d.columns.map((_, c) => c), 'reject') : noDecisions()));
    const initial = rejectAll(1);
    let calls = 0;
    window.W = new SpreadsheetDiff({ review: true, tooltips: false, decisions: initial, onDecisionsChange: () => calls++ });
    await window.W.render({ oldFile: { name: 'old.xlsx', data: new Uint8Array(a) }, newFile: { name: 'new.xlsx', data: new Uint8Array(b) }, container: document.getElementById('grid') });
    const applied = window.W.getDecisions()[sheet] === initial[sheet];
    const more = rejectAll(d.rowCount);
    window.W.setOptions({ decisions: more });
    const replaced = window.W.getDecisions()[sheet] === more[sheet];
    // Passing back what onDecisionsChange gave is a no-op.
    const own = window.W.getDecisions();
    window.W.setOptions({ decisions: own });
    return { applied, replaced, same: window.W.getDecisions() === own, calls };
  }, { a: fixtureBytes('old.xlsx'), b: fixtureBytes('new.xlsx') });
  assert(r.applied && r.replaced && r.same && r.calls === 0, JSON.stringify(r));
});

await test('SpreadsheetDiff (React): decisions kept in state round-trip without a loop', async () => {
  const r = await page.evaluate(async ({ a, b }) => {
    window.W?.cleanUp();
    const { SpreadsheetDiff, createElement: h, useState, createRoot } = window.sdReact;
    const container = document.getElementById('grid');
    let changes = 0;
    let handle;
    const files = { old: { name: 'old.xlsx', data: new Uint8Array(a) }, new: { name: 'new.xlsx', data: new Uint8Array(b) } };
    function App() {
      const [decisions, setDecisions] = useState(undefined);
      return h(SpreadsheetDiff, { ref: (x) => (handle = x), oldFile: files.old, newFile: files.new, review: true, decisions, onDecisionsChange: (d) => { changes++; setDecisions(d); }, style: { height: 500 } });
    }
    const root = createRoot(container);
    root.render(h(App));
    for (let i = 0; i < 100 && !handle?.instance?.cmp; i++) await new Promise((res) => setTimeout(res, 20));
    const w = handle.instance;
    w.grid.selectRange(0, 0, w.grid.diff.rowCount - 1, w.grid.diff.columns.length - 1);
    handle.decideSelection('reject');
    await new Promise((res) => setTimeout(res, 200));
    const out = { changes, rejected: w.getDecisions().some((x) => x.cells.size + x.rows.size + x.columns.size > 0) };
    root.unmount();
    return out;
  }, { a: fixtureBytes('old.xlsx'), b: fixtureBytes('new.xlsx') });
  assert(r.changes === 1 && r.rejected, JSON.stringify(r));
});

await test('SpreadsheetDiff: a newer render cancels the fetches of the one it replaces', async () => {
  let release;
  const held = new Promise((res) => (release = res));
  await page.route('**/held.csv', async (route) => { await held; await route.fulfill({ body: 'a\n1' }).catch(() => {}); });
  const r = await page.evaluate(async () => {
    window.W?.cleanUp();
    const { SpreadsheetDiff } = window.sd;
    const signals = [];
    const realFetch = window.fetch;
    window.fetch = (input, init) => { signals.push(init?.signal); return realFetch(input, init); };
    window.W = new SpreadsheetDiff({ tooltips: false });
    const container = document.getElementById('grid');
    const held = new URL('/held.csv', location.href);
    const first = window.W.render({ oldFile: held, newFile: held, container }).then(() => 'resolved', (e) => `rejected: ${e.message}`);
    await new Promise((res) => setTimeout(res, 50));
    await window.W.render({ oldFile: { name: 'a.csv', data: new TextEncoder().encode('a\n1') }, newFile: { name: 'b.csv', data: new TextEncoder().encode('a\n2') }, container });
    window.fetch = realFetch;
    return { aborted: signals.length > 0 && signals.every((s) => s?.aborted), shown: window.W.cmp?.sheets.length, first: await Promise.race([first, new Promise((res) => setTimeout(() => res('pending'), 300))]) };
  });
  release();
  await page.unroute('**/held.csv');
  assert(r.aborted && r.shown === 1, JSON.stringify(r));
  assert(r.first === 'resolved', `the replaced render should settle quietly: ${r.first}`);
});

assert(true);
if (errors.length) {
  failed++;
  console.log(`  ✗ page errors:\n    ${errors.join('\n    ')}`);
}
await browser.close();
console.log(failed ? `\n${failed} failed` : '\nall passed');
process.exit(failed ? 1 : 0);
