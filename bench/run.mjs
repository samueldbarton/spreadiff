// Benchmark: pipeline timings + scroll smoothness, spreadiff vs Glide Data Grid.
// Usage: (cd apps/viewer && npx vite build && npx vite preview --port 4173) & node bench/run.mjs
// Env: URL (default http://localhost:4173/), SIZES (default 100000,1000000), CHROMIUM.
import { chromium } from 'playwright';
import { writeFileSync, mkdirSync } from 'node:fs';

const URL = process.env.URL ?? 'http://localhost:4173/';
const SIZES = (process.env.SIZES ?? '100000,1000000').split(',').map(Number);
const SCROLL_FRAMES = 240;

const browser = await chromium.launch({ executablePath: process.env.CHROMIUM ?? '/opt/pw-browsers/chromium', args: ['--enable-gpu-rasterization', '--ignore-gpu-blocklist'] });
const results = [];

async function scrollBench(page, renderer) {
  return page.evaluate(
    async ({ renderer, frames }) => {
      const el =
        renderer === 'spreadiff'
          ? document.querySelector('.spreadiff').shadowRoot.querySelector('.sd-scroller')
          : document.querySelector('.dvn-scroller');
      el.scrollTop = 0;
      await new Promise((r) => setTimeout(r, 300));
      const times = [];
      const slow = [];
      let last = performance.now();
      let long = 0;
      const obs = new PerformanceObserver((l) => (long += l.getEntries().length));
      try { obs.observe({ type: 'longtask', buffered: false }); } catch {}
      for (let i = 0; i < frames; i++) {
        el.scrollTop += 90; // ~3 rows per frame, plus a horizontal wobble
        if (i % 30 === 0) el.scrollLeft = (i % 60) * 20;
        await new Promise((r) => requestAnimationFrame(r));
        const now = performance.now();
        times.push(now - last);
        if (now - last > 50) slow.push([i, Math.round(now - last)]); // which frame, and how long
        last = now;
      }
      obs.disconnect();
      times.sort((a, b) => a - b);
      const q = (p) => times[Math.min(times.length - 1, Math.floor(p * times.length))];
      const mem = performance.memory ? Math.round(performance.memory.usedJSHeapSize / 1e6) : null;
      return { slow, p50: q(0.5), p95: q(0.95), max: times[times.length - 1], over25: times.filter((t) => t > 25).length, longTasks: long, heapMB: mem };
    },
    { renderer, frames: SCROLL_FRAMES },
  );
}

for (const rows of SIZES) {
  const page = await browser.newPage({ viewport: { width: 1440, height: 900 }, deviceScaleFactor: 2 });
  await page.goto(URL);
  // The viewer has two generated stress tests; other sizes aren't wired up.
  const label = { 100000: 'Stress test', 1000000: '1M rows' }[rows];
  if (!label) throw new Error(`SIZES: only 100000 and 1000000 are supported, got ${rows}`);
  await page.getByText(label).first().click();
  await page.waitForSelector('.timing', { timeout: 180_000 });
  const timing = (await page.textContent('.timing')).trim();
  // Fair comparison: spreadiff draws every row too.
  await page.getByText('Collapse unchanged').click();
  await page.waitForTimeout(300);
  const sd = await scrollBench(page, 'spreadiff');
  await page.getByRole('radio', { name: 'Glide' }).click();
  await page.waitForSelector('.dvn-scroller', { timeout: 60_000 });
  await page.waitForTimeout(500);
  const gl = await scrollBench(page, 'glide');
  results.push({ rows, timing, spreadiff: sd, glide: gl });
  console.log(JSON.stringify(results[results.length - 1]));
  await page.close();
}
await browser.close();
mkdirSync('bench/results', { recursive: true });
writeFileSync('bench/results/latest.json', JSON.stringify({ date: new Date().toISOString(), results }, null, 2));
