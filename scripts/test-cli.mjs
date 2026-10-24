// End-to-end test of the spreadiff-lite CLI, as a user would run it: build
// both packages, write two .xlsx files, run `spreadiff-lite` from the
// package's bin, and check its output and exit codes.
import { execFileSync, spawnSync } from 'node:child_process';
import { mkdtempSync, writeFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import * as esbuild from 'esbuild';

const ROOT = fileURLToPath(new URL('..', import.meta.url));
const run = (cmd, args) => execFileSync(cmd, args, { cwd: ROOT, stdio: ['ignore', 'pipe', 'inherit'] });

run('npm', ['run', 'build', '-w', 'spreadiff']);
run('npm', ['run', 'build', '-w', 'spreadiff-lite']);

// The fixture builder is TypeScript; bundle it to run here.
const dir = mkdtempSync(join(tmpdir(), 'spreadiff-cli-'));
// Inside the repository, so its imports (exceljs) resolve.
const builder = join(ROOT, 'node_modules/.cache/spreadiff-cli/fixtures.mjs');
await esbuild.build({ entryPoints: [join(ROOT, 'packages/spreadiff-lite/test/fixtures.ts')], bundle: true, platform: 'node', format: 'esm', outfile: builder, packages: 'external', logLevel: 'error' });
const { regionalModel } = await import(builder);
const base = { growth: 0.1, inr: 83, cny: 7.2, indiaRevenue2024: 8300, chinaRevenue2024: 720, indiaCosts2024: 4150, chinaCosts2024: 360 };
const a = join(dir, 'regions-v1.xlsx');
const b = join(dir, 'regions-v2.xlsx');
writeFileSync(a, await regionalModel(base));
writeFileSync(b, await regionalModel({ ...base, growth: 0.12, inr: 84 }));

const bin = join(ROOT, 'packages/spreadiff-lite/dist/cli.js');
const cli = (...args) => spawnSync(process.execPath, [bin, ...args], { encoding: 'utf8' });
const failures = [];
const check = (name, ok, detail = '') => {
  console.log(`  ${ok ? '✓' : '✗'} ${name}${ok ? '' : `\n${detail}`}`);
  if (!ok) failures.push(name);
};

const md = cli(a, b);
check('changed files exit 1', md.status === 1, md.stderr);
check('Markdown names the Bloomberg input', md.stdout.includes('| `Inputs!B2` | USD/INR *(Bloomberg BDP)* | 83 | 84 |'), md.stdout);
check('Markdown attributes downstream results', /\| `Model!C6` \| India › Profit \(USD\) · 2025 \|.*Growth \+0\.9881; USD\/INR −0\.6667/.test(md.stdout), md.stdout);

const same = cli(a, a);
check('identical files exit 0', same.status === 0 && same.stdout.includes('No changes.'), same.stdout + same.stderr);

const json = cli(a, b, '--json');
let parsed = null;
try {
  parsed = JSON.parse(json.stdout);
} catch {}
check('--json prints the report', parsed?.totals?.external === 1 && parsed.edits.length === 2 && !('diff' in parsed), json.stdout.slice(0, 500));

const outputs = cli(a, b, '--outputs-only', '--sheet', 'Summary', '--sheet', 'Model', '--sheet', 'Inputs');
check('--outputs-only drops intermediate results', outputs.stdout.includes('Final outputs') && !outputs.stdout.includes('Intermediate results'), outputs.stdout);

const bad = cli(a);
check('a usage error exits 2', bad.status === 2 && /needs two files/.test(bad.stderr), bad.stderr);
const missing = cli(a, join(dir, 'nope.xlsx'));
check('a missing file exits 2', missing.status === 2 && /can't read/.test(missing.stderr), missing.stderr);
const help = cli('--help');
check('--help', help.status === 0 && help.stdout.startsWith('Usage: spreadiff-lite'), help.stdout);

// As installed from npm: pack both packages, install the tarballs into an
// empty project, and run the installed bin.
const project = join(dir, 'project');
execFileSync('mkdir', ['-p', project]);
writeFileSync(join(project, 'package.json'), '{"name":"try-spreadiff-lite","private":true,"type":"module"}');
const pack = (ws) => join(dir, execFileSync('npm', ['pack', '-w', ws, '--pack-destination', dir], { cwd: ROOT, encoding: 'utf8', stdio: ['ignore', 'pipe', 'ignore'] }).trim().split('\n').pop());
const tarballs = [pack('spreadiff'), pack('spreadiff-lite')];
execFileSync('npm', ['install', '--no-audit', '--no-fund', '--silent', ...tarballs], { cwd: project, stdio: 'inherit' });
const installed = spawnSync(join(project, 'node_modules/.bin/spreadiff-lite'), [a, b], { encoding: 'utf8' });
check('installed from tarballs, the bin runs', installed.status === 1 && installed.stdout.includes('USD/INR *(Bloomberg BDP)*'), installed.stdout + installed.stderr);
writeFileSync(join(project, 'use.mjs'), "import { summarize } from 'spreadiff-lite';\nimport { readFileSync } from 'node:fs';\nconst [a, b] = process.argv.slice(2);\nprocess.stdout.write(summarize({ name: 'a.xlsx', data: readFileSync(a) }, { name: 'b.xlsx', data: readFileSync(b) }));\n");
const api = spawnSync(process.execPath, [join(project, 'use.mjs'), a, b], { encoding: 'utf8' });
check('installed from tarballs, the API imports', api.status === 0 && api.stdout.startsWith('## a.xlsx → b.xlsx'), api.stdout + api.stderr);

rmSync(dir, { recursive: true, force: true });
if (failures.length) {
  console.error(`\n✗ ${failures.length} CLI check${failures.length === 1 ? '' : 's'} failed`);
  process.exit(1);
}
console.log('\n✓ CLI test passed');
