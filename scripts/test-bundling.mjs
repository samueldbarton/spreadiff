#!/usr/bin/env node
// Packs the library, installs it into a throwaway Vite app the way a user
// would, builds it, and runs it in Chromium. Verifies that:
//   - package `exports` resolve,
//   - both documented worker setups work (docs/workers.md),
//   - the worker entry isn't tree-shaken away (`sideEffects`).
// Usage: npm run test:bundle   (needs network access for the first npm install)
import { execSync, spawn } from 'node:child_process';
import { mkdirSync, rmSync, writeFileSync, readdirSync } from 'node:fs';
import { join, resolve, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { chromium } from 'playwright';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const APP = join(ROOT, 'node_modules/.bundle-test');
const sh = (cmd, cwd) => execSync(cmd, { cwd, stdio: ['ignore', 'pipe', 'inherit'], encoding: 'utf8' });

rmSync(APP, { recursive: true, force: true });
mkdirSync(join(APP, 'src'), { recursive: true });
sh('npm run build', join(ROOT, 'packages/spreadiff'));
const tgz = sh(`npm pack --pack-destination ${APP}`, join(ROOT, 'packages/spreadiff')).trim().split('\n').pop();
writeFileSync(join(APP, 'package.json'), JSON.stringify({ name: 'bundle-test', private: true, type: 'module' }));
sh(`npm install --no-audit --no-fund ./${tgz} vite@8`, APP);
writeFileSync(join(APP, 'index.html'), '<!doctype html><pre id="out"></pre><script type="module" src="/src/main.ts"></script>');
writeFileSync(join(APP, 'src/diff.worker.ts'), "import 'spreadiff/worker/worker.js';\n");
writeFileSync(
  join(APP, 'src/main.ts'),
  `import { createDiffWorker } from 'spreadiff/worker';
import { diffTables, parseCsv } from 'spreadiff';
import { renderText } from 'spreadiff/html';
const enc = (s: string) => new TextEncoder().encode(s).buffer as ArrayBuffer;
const run = (label: string, w: Worker) =>
  Promise.race([
    createDiffWorker(w)
      .diffBuffers({ name: 'a.csv', data: enc('id,v\\n1,a') }, { name: 'b.csv', data: enc('id,v\\n1,b') })
      .then((r) => \`\${label}: \${r.result.sheets[0].diff.stats.cellsChanged === 1 ? 'ok' : 'WRONG RESULT'}\`),
    new Promise<string>((r) => setTimeout(() => r(\`\${label}: no reply\`), 5000)),
  ]);
const html = renderText(diffTables(parseCsv('a\\n1'), parseCsv('a\\n2'))).includes('+  2') ? 'html: ok' : 'html: WRONG';
Promise.all([
  run('worker (package URL)', new Worker(new URL('spreadiff/worker/worker.js', import.meta.url), { type: 'module' })),
  run('worker (own file)', new Worker(new URL('./diff.worker.ts', import.meta.url), { type: 'module' })),
]).then((lines) => { document.getElementById('out')!.textContent = [html, ...lines].join('\\n'); });
`,
);
sh('npx vite build', APP);
const sizes = readdirSync(join(APP, 'dist/assets')).filter((f) => f.endsWith('.js'));
const server = spawn('npx', ['vite', 'preview', '--port', '4310', '--strictPort'], { cwd: APP, stdio: 'ignore' });
try {
  await new Promise((r) => setTimeout(r, 1500));
  const browser = await chromium.launch({ executablePath: process.env.CHROMIUM ?? '/opt/pw-browsers/chromium' });
  const page = await browser.newPage();
  await page.goto('http://localhost:4310/');
  await page.waitForFunction(() => document.getElementById('out')?.textContent, null, { timeout: 15000 });
  const out = await page.textContent('#out');
  await browser.close();
  console.log(out);
  if (!/^(html: ok\n)(worker \(package URL\): ok\n)(worker \(own file\): ok)$/.test(out)) {
    console.error(`✗ bundling test failed (assets: ${sizes.join(', ')})`);
    process.exitCode = 1;
  } else console.log('✓ bundling test passed');
} finally {
  server.kill();
}
