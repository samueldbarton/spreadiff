#!/usr/bin/env node
// Regenerates every image in docs/images from the current source.
// Usage: npm run docs:screenshots
//
// Component scenes come from apps/viewer/shots.html (the library alone, no
// app chrome); app views from the viewer itself. Pointer positions are
// computed from the grid's layout, so they survive layout changes.
import { spawn } from 'node:child_process';
import { mkdirSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { chromium } from 'playwright';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const OUT = join(ROOT, 'docs/images');
const PORT = 5299;
mkdirSync(OUT, { recursive: true });

const server = spawn('npx', ['vite', '--port', String(PORT), '--strictPort'], { cwd: join(ROOT, 'apps/viewer'), stdio: 'ignore' });
const errors = [];
try {
  await new Promise((r) => setTimeout(r, 2500));
  const browser = await chromium.launch({ executablePath: process.env.CHROMIUM ?? '/opt/pw-browsers/chromium' });

  async function stagePage(scheme = 'light') {
    const page = await browser.newPage({ viewport: { width: 1600, height: 720 }, deviceScaleFactor: 2, colorScheme: scheme });
    page.on('pageerror', (e) => errors.push(e.message));
    await page.goto(`http://localhost:${PORT}/shots.html`);
    await page.waitForFunction(() => window.ready);
    await page.evaluate(() => document.fonts.ready);
    return page;
  }
  /** Client point of the cell in the row whose first column reads `label`. */
  const cellPoint = (page, label, column) =>
    page.evaluate(
      ({ label, column }) => {
        const g = window.grid;
        const d = g.diff;
        let row = -1;
        for (let r = 0; r < d.rowCount; r++) {
          const ni = d.newRows[r];
          const v = ni >= 0 ? d.newTable.rows[ni][0] : d.oldTable.rows[d.oldRows[r]][0];
          if (v === label) row = r;
        }
        const ci = d.columns.findIndex((c) => c.name === column);
        const vi = Array.from(g.cols.cols).indexOf(ci);
        const pane = g.panes()[0];
        const disp = g.rows.displayOf[row];
        const box = g.scroller.getBoundingClientRect();
        const x = box.left + pane.x + pane.gutter + g.cols.offsets[vi] + g.cols.widths[vi] / 2 - (vi < g.cols.frozen ? 0 : g.scroller.scrollLeft);
        const y = box.top + 36 + (g.rows.offsets[disp] + g.rows.offsets[disp + 1]) / 2 - g.scroller.scrollTop;
        return { x, y, row, col: ci };
      },
      { label, column },
    );
  const shot = async (page, name) => {
    await page.waitForTimeout(250);
    await page.locator('#stage').screenshot({ path: join(OUT, `${name}.png`) });
    console.log(`  docs/images/${name}.png`);
  };
  const scene = (page, s) => page.evaluate((s) => window.scene(s), s);

  // ---- Grid scenes
  let page = await stagePage();
  await scene(page, { sample: 'inventory', sheet: 'Inventory', layout: 'unified', height: 540 });
  const p = await cellPoint(page, 'P-1003', 'price');
  await page.mouse.move(p.x, p.y);
  await page.waitForTimeout(700);
  await shot(page, 'grid-unified');
  await page.mouse.move(5, 5);
  await scene(page, { sample: 'inventory', sheet: 'Inventory', layout: 'split', height: 460, width: 1560 });
  await shot(page, 'grid-split');
  await scene(page, { sample: 'cities', layout: 'split', annotations: true, height: 300 });
  await shot(page, 'grid-annotations');
  for (const display of ['stacked', 'new', 'delta', 'inline']) {
    await scene(page, { sample: 'model', sheet: 'Revenue', layout: 'unified', display, trace: true, height: display === 'new' || display === 'inline' ? 260 : 330 });
    await shot(page, `display-${display}`);
  }
  await scene(page, { sample: 'model', sheet: 'Summary', layout: 'unified', trace: true, height: 560 });
  const ni = await cellPoint(page, 'Net income', 'value');
  await page.mouse.move(ni.x, ni.y);
  await page.waitForTimeout(700);
  await shot(page, 'trace-tooltip');
  await page.mouse.move(5, 5);
  await scene(page, { sample: 'model', sheet: 'Revenue', layout: 'unified', trace: true, height: 330 });
  const emea = await cellPoint(page, 'EMEA', 'revenue');
  await page.evaluate(({ row, col }) => window.grid.select(row, col), emea);
  await shot(page, 'trace-focus');
  // ---- HTML renderer scenes
  await scene(page, { sample: 'model', sheet: 'Revenue', renderer: 'html', layout: 'split', trace: true });
  await shot(page, 'html-split');
  await scene(page, { sample: 'model', sheet: 'Revenue', renderer: 'html', layout: 'unified', trace: true });
  await shot(page, 'html-unified');
  await page.close();
  // ---- Dark
  page = await stagePage('dark');
  await scene(page, { sample: 'model', sheet: 'Revenue', layout: 'unified', display: 'delta', trace: true, height: 330 });
  await shot(page, 'grid-dark');
  await page.close();

  // ---- The viewer app
  const app = await browser.newPage({ viewport: { width: 1440, height: 820 }, deviceScaleFactor: 2, colorScheme: 'light' });
  app.on('pageerror', (e) => errors.push(e.message));
  await app.goto(`http://localhost:${PORT}/`);
  await app.getByText('Financial model').click();
  await app.waitForSelector('.timing');
  await app.getByRole('radio', { name: 'Δ %' }).click();
  await app.waitForTimeout(600);
  await app.screenshot({ path: join(OUT, 'viewer.png') });
  console.log('  docs/images/viewer.png');
  const panel = app.locator('.trace-panel');
  await panel.screenshot({ path: join(OUT, 'viewer-outputs.png') });
  await app.getByRole('tab', { name: 'Impact' }).click();
  await app.waitForTimeout(300);
  await panel.screenshot({ path: join(OUT, 'viewer-impact.png') });
  await app.getByRole('tab', { name: 'Outputs' }).click();
  await app.locator('.tp-card').first().locator('.taddr').first().click();
  await app.waitForTimeout(400);
  await app.getByRole('tab', { name: 'Why?' }).click();
  await app.waitForTimeout(400);
  await panel.screenshot({ path: join(OUT, 'viewer-why.png') });
  console.log('  docs/images/viewer-{outputs,impact,why}.png');
  await browser.close();
} finally {
  server.kill();
}
if (errors.length) {
  console.error(`page errors:\n${errors.join('\n')}`);
  process.exit(1);
}
