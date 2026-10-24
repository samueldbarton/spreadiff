#!/usr/bin/env node
// Checks downstream tracing where it will actually run: a static website.
// No LibreOffice and no server-side code; every request outside the page's
// own origin is blocked.
//
// 1. Library: packs spreadiff, installs it into a fresh Vite app, builds it,
//    serves dist/ as plain files and runs test/hosted/main.ts in Chromium.
//    Each fixture (test/hosted/fixtures: real .xlsx pairs with expectations
//    exported from LibreOffice) is parsed, diffed and traced in the worker.
//    Every changed cell's kind and causes are compared with LibreOffice's
//    but-for answers. Values recalculated from stale files, and values
//    computed for files saved without them, must equal LibreOffice's.
// 2. Viewer: serves the production build of apps/viewer, uploads fixtures
//    through its file inputs and checks the trace panel, banners and
//    contributions. Measures main-thread long tasks while it works.
//
// Usage: npm run test:hosted   (needs network access for the first npm install)
import { execSync } from 'node:child_process';
import { cpSync, createReadStream, existsSync, mkdirSync, readdirSync, readFileSync, rmSync, statSync, writeFileSync } from 'node:fs';
import { createServer } from 'node:http';
import { createRequire } from 'node:module';
import { dirname, extname, join, normalize, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { chromium } from 'playwright';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const APP = join(ROOT, 'node_modules/.hosted-test');
const FIXTURES = join(ROOT, 'test/hosted/fixtures');
const sh = (cmd, cwd) => execSync(cmd, { cwd, stdio: ['ignore', 'pipe', 'inherit'], encoding: 'utf8' });
const cases = readdirSync(FIXTURES).filter((c) => existsSync(join(FIXTURES, c, 'expected.json')));

// ---- Build both sites
rmSync(APP, { recursive: true, force: true });
mkdirSync(join(APP, 'src'), { recursive: true });
sh('npm run build', join(ROOT, 'packages/spreadiff'));
const tgz = sh(`npm pack --pack-destination ${APP}`, join(ROOT, 'packages/spreadiff')).trim().split('\n').pop();
writeFileSync(join(APP, 'package.json'), JSON.stringify({ name: 'hosted-test', private: true, type: 'module' }));
sh(`npm install --no-audit --no-fund ./${tgz} vite@8`, APP);
// No inline style attribute: the page is served with a strict CSP (main.ts sizes the grid through the CSSOM).
writeFileSync(join(APP, 'index.html'), '<!doctype html><div id="grid"></div><script type="module" src="/src/main.ts"></script>');
cpSync(join(ROOT, 'test/hosted/main.ts'), join(APP, 'src/main.ts'));
cpSync(FIXTURES, join(APP, 'public/fixtures'), { recursive: true });
writeFileSync(join(APP, 'public/fixtures/index.json'), JSON.stringify(cases));
sh('npx vite build', APP);
sh('npm run build -w apps/viewer', ROOT);

// ---- Plain static file server: what any host (S3, Pages, Netlify) does
const TYPES = { '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css', '.json': 'application/json', '.svg': 'image/svg+xml', '.xlsx': 'application/octet-stream', '.woff2': 'font/woff2', '.png': 'image/png' };
function serve(dir) {
  const server = createServer((req, res) => {
    const path = normalize(decodeURIComponent(new URL(req.url, 'http://x').pathname)).replace(/^(\.\.[/\\])+/, '');
    let file = join(dir, path);
    if (!file.startsWith(dir) || !existsSync(file)) return res.writeHead(404).end();
    if (statSync(file).isDirectory()) file = join(file, 'index.html');
    // A strict Content-Security-Policy: no inline styles or scripts, no eval, nothing from elsewhere.
    res.writeHead(200, {
      'content-type': TYPES[extname(file)] ?? 'application/octet-stream',
      'content-security-policy': "default-src 'self'; script-src 'self'; style-src 'self'; img-src 'self' data:; font-src 'self'; worker-src 'self'; connect-src 'self'",
    });
    createReadStream(file).pipe(res);
  });
  return new Promise((r) => server.listen(0, '127.0.0.1', () => r(server)));
}

const browser = await chromium.launch({ executablePath: process.env.CHROMIUM ?? '/opt/pw-browsers/chromium' });
let failed = 0;
async function newPage(origin) {
  const context = await browser.newContext({ viewport: { width: 1400, height: 900 } });
  const blocked = [];
  await context.route('**/*', (route) => {
    if (route.request().url().startsWith(origin)) return route.continue();
    blocked.push(route.request().url());
    return route.abort();
  });
  const page = await context.newPage();
  const errors = [];
  page.on('pageerror', (e) => errors.push(e.message));
  // CSP violations don't throw; collect them from the page.
  await page.addInitScript(() => document.addEventListener('securitypolicyviolation', (e) => { (window.__csp ??= []).push(`${e.violatedDirective}: ${e.blockedURI || e.sample || 'inline'}`); }));
  page.on('console', (m) => { if (/Content Security Policy|Refused to/.test(m.text())) errors.push(`console: ${m.text()}`); });
  return { page, blocked, errors };
}

// ---- 1. The library in a fresh app
{
  const server = await serve(join(APP, 'dist'));
  const origin = `http://127.0.0.1:${server.address().port}`;
  const { page, blocked, errors } = await newPage(origin);
  await page.goto(`${origin}/`);
  await page.waitForFunction(() => window.done, null, { timeout: 300_000 });
  const { report, longestGap, controlGap } = await page.evaluate(() => window.done);
  const csp = await page.evaluate(() => window.__csp ?? []);
  if (csp.length) { failed++; console.log('   ✗ CSP violations (library app):', csp.slice(0, 5)); }
  server.close();
  console.log('\n## Library in a fresh Vite app (worker, static hosting, no network)\n');
  console.log('case | changed | causes right | contributions right | stale values right | missing values right | ms');
  let cells = 0, scored = 0, scoredWant = 0, contrib = 0, contribWant = 0, stale = 0, staleWant = 0, unc = 0, uncWant = 0;
  for (const r of report) {
    const want = Object.values(JSON.parse(readFileSync(join(FIXTURES, r.name, 'expected.json'), 'utf8')).cells).filter((c) => c.kind === 'computed' && c.causes).length;
    cells += r.cells; scored += r.scored; scoredWant += want; contrib += r.contribOk; contribWant += r.contribChecked; stale += r.staleOk; staleWant += r.staleChecked; unc += r.uncachedOk; uncWant += r.uncachedChecked;
    console.log(`${r.name} | ${r.cells} | ${r.scored}/${want} | ${r.contribOk}/${r.contribChecked} | ${r.staleChecked ? `${r.staleOk}/${r.staleChecked}` : '—'} | ${r.uncachedChecked ? `${r.uncachedOk}/${r.uncachedChecked}` : '—'} | ${r.ms}`);
    for (const p of r.problems.slice(0, 10)) console.log(`   ✗ ${p}`);
    if (r.problems.length) failed++;
  }
  console.log(`TOTAL ${report.length} cases · ${cells} changed cells · causes ${scored}/${scoredWant} · contributions ${contrib}/${contribWant} · stale ${stale}/${staleWant} · missing ${unc}/${uncWant}`);
  console.log(`main thread: longest stall ${Math.round(longestGap)} ms (10 ms timer); control, inventory traced on the main thread instead: ${Math.round(controlGap)} ms · blocked requests ${blocked.length} · page errors ${errors.length}`);
  if (controlGap < 2 * longestGap) { failed++; console.log('   ✗ the control stall should dwarf the worker one: the stall measurement is not working'); }
  if (blocked.length || errors.length) { failed++; console.log(blocked, errors); }
}

// ---- 2. The viewer's production build
{
  const server = await serve(join(ROOT, 'apps/viewer/dist'));
  const origin = `http://127.0.0.1:${server.address().port}`;
  const { page, blocked, errors } = await newPage(origin);
  console.log('\n## Viewer production build\n');
  const check = (ok, msg) => { console.log(`  ${ok ? '✓' : '✗'} ${msg}`); if (!ok) failed++; };
  const upload = async (name, after) => {
    await page.goto(`${origin}/`);
    await page.evaluate(() => {
      // Longest gap between ticks of a 10 ms timer while the files load (the Long Tasks API reports nothing headless).
      let last = performance.now();
      window.__gap = 0;
      setInterval(() => { const now = performance.now(); window.__gap = Math.max(window.__gap, now - last); last = now; }, 10);
    });
    await page.locator('.drop.old input[type=file]').setInputFiles(join(FIXTURES, name, 'old.xlsx'));
    await page.locator('.drop.new input[type=file]').setInputFiles(join(FIXTURES, name, after));
    await page.waitForFunction(() => !!document.querySelector('.tp-foot'), null, { polling: 250, timeout: 120_000 });
    return page.evaluate(() => window.__gap);
  };
  const counts = (name) => {
    const c = Object.values(JSON.parse(readFileSync(join(FIXTURES, name, 'expected.json'), 'utf8')).cells);
    return { inputs: c.filter((x) => x.kind === 'input').length, formulas: c.filter((x) => x.kind === 'formula').length, computed: c.filter((x) => x.kind === 'computed').length };
  };
  for (const name of ['three_statement', 'inventory', 'loan']) {
    const longest = await upload(name, 'new.xlsx');
    const foot = await page.textContent('.tp-foot');
    const n = counts(name);
    const m = /(\d+) input edits?.*?(\d+) formula edits?.*?(\d+) downstream/.exec(foot ?? '');
    check(m && +m[1] === n.inputs && +m[2] === n.formulas && +m[3] === n.computed, `${name}: panel says "${foot?.trim()}", LibreOffice: ${n.inputs} inputs, ${n.formulas} formula edits, ${n.computed} downstream`);
    check(/causes all recalculated/.test(foot ?? ''), `${name}: every downstream cell's causes recalculated`);
    check((await page.locator('.tshare').count()) > 0, `${name}: contributions shown`);
    check(longest < 250, `${name}: longest main-thread stall ${Math.round(longest)} ms`);
  }
  // Accessibility of the whole viewer, with a model loaded (axe-core; serious and critical only).
  await upload('three_statement', 'new.xlsx');
  // Injected through the DevTools protocol: an inline <script> would (rightly) be refused by the page's CSP.
  await page.evaluate(readFileSync(createRequire(import.meta.url).resolve('axe-core/axe.min.js'), 'utf8'));
  const violations = await page.evaluate(async () =>
    (await window.axe.run(document, { resultTypes: ['violations'] })).violations
      .filter((v) => v.impact === 'serious' || v.impact === 'critical')
      .map((v) => `${v.impact} ${v.id}: ${v.help} — ${v.nodes.slice(0, 3).map((n) => n.target.join(' ')).join(', ')}`),
  );
  check(!violations.length, `viewer: no serious accessibility violations${violations.length ? `\n      ${violations.join('\n      ')}` : ''}`);
  await upload('loan', 'stale.xlsx');
  const banner = await page.textContent('.banner:not(.info)').catch(() => null);
  check(/saved without recalculating/.test(banner ?? ''), `loan, stale file: banner "${banner?.trim().slice(0, 90)}…"`);
  await upload('inventory', 'uncached.xlsx');
  const info = await page.textContent('.banner.info').catch(() => null);
  check(/computed/i.test(info ?? ''), `inventory, file without values: note "${info?.trim().slice(0, 90)}…"`);
  // Large models from `npm run bench:trace`, when they've been built: timings, reported only.
  const SCALE = join(ROOT, 'bench/trace-eval/work/scale-cases');
  for (const name of existsSync(SCALE) ? ['chain20k', 'grid30k', 'grid100k'] : []) {
    await page.goto(`${origin}/`);
    await page.evaluate(() => {
      let last = performance.now();
      window.__gap = 0;
      setInterval(() => { const now = performance.now(); window.__gap = Math.max(window.__gap, now - last); last = now; }, 10);
    });
    const t0 = Date.now();
    await page.locator('.drop.old input[type=file]').setInputFiles(join(SCALE, name, 'old.xlsx'));
    await page.locator('.drop.new input[type=file]').setInputFiles(join(SCALE, name, 'new.xlsx'));
    // Poll gently: Playwright's default visibility polling forces a layout every frame and would itself stall the page.
    await page.waitForFunction(() => !!document.querySelector('.tp-foot'), null, { polling: 500, timeout: 300_000 });
    const total = Date.now() - t0;
    const gap = await page.evaluate(() => window.__gap);
    console.log(`  · ${name}: ${(total / 1000).toFixed(1)} s from upload to trace panel (±0.5 s polling), longest main-thread stall ${Math.round(gap)} ms · ${(await page.textContent('.tp-foot'))?.trim()}`);
  }
  const csp = await page.evaluate(() => window.__csp ?? []);
  check(!csp.length, `strict CSP: no violations ${csp.slice(0, 5).join(' · ')}`);
  check(!blocked.length && !errors.length, `no outside requests (${blocked.length}) and no page errors (${errors.length}) ${errors.join(' ')}`);
  server.close();
}

await browser.close();
if (failed) {
  console.error(`\n✗ hosted test: ${failed} failure(s)`);
  process.exitCode = 1;
} else console.log('\n✓ hosted test passed');
