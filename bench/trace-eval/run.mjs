#!/usr/bin/env node
// Evaluates formula tracing against LibreOffice, which really recalculates.
// Usage: npm run bench:trace            (needs LibreOffice and Python openpyxl)
//        npm run bench:trace -- --skip-build   (reuse the cases in work/)
//
// Builds, in bench/trace-eval/work/ (git-ignored):
//   cases/        five realistic models (financial, budget, inventory, loan, commission)
//   poi-cases/    real files saved by Microsoft Excel, from the Apache POI test corpus
//   scale-cases/  large models (40k–300k formulas)
// then prints the results that docs/tracing.md and docs/formats.md quote.
import { execFileSync } from 'node:child_process';
import { existsSync, mkdirSync, readdirSync, readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { build } from 'esbuild';

const HERE = dirname(fileURLToPath(import.meta.url));
const WORK = join(HERE, 'work');
mkdirSync(WORK, { recursive: true });
const sh = (cmd, args, opts = {}) => execFileSync(cmd, args, { cwd: WORK, stdio: 'inherit', ...opts });
const py = (script) => sh('python3', [join(HERE, script)], { env: { ...process.env, PYTHONPATH: HERE } });

if (!process.argv.includes('--skip-build')) {
  console.log('## Building the models with LibreOffice');
  py('build.py');
  if (!existsSync(join(WORK, 'poi'))) {
    console.log('## Fetching the Apache POI spreadsheet test files');
    sh('git', ['clone', '-q', '--depth', '1', '--filter=blob:none', '--no-checkout', 'https://github.com/apache/poi', 'poi']);
    const files = execFileSync('git', ['ls-tree', '-r', '--name-only', '-z', 'HEAD', 'test-data/spreadsheet'], { cwd: join(WORK, 'poi') })
      .toString()
      .split('\0')
      .filter((f) => /\.(xlsx|xlsb|xls)$/i.test(f));
    execFileSync('git', ['checkout', '-q', 'HEAD', '--', ...files], { cwd: join(WORK, 'poi'), stdio: 'inherit' });
  }
  py('survey.py');
  py('poi_build.py');
  py('build_scale.py');
}

for (const name of ['eval', 'engine-check', 'uncached-check', 'scale', 'reader-check', 'format-check']) {
  await build({ entryPoints: [join(HERE, `${name}.ts`)], bundle: true, platform: 'node', format: 'esm', outfile: join(WORK, `${name}.mjs`), logLevel: 'warning' });
}
const node = (script, ...args) => sh(process.execPath, ['--max-old-space-size=6000', `${script}.mjs`, ...args]);
console.log('\n## Engine: every formula recalculated from its cached inputs, vs the cached result');
node('engine-check', 'cases', 'poi-cases');
console.log('\n## Missing values computed, vs LibreOffice');
node('uncached-check');
console.log('\n## Trace vs the LibreOffice oracle: models');
node('eval', 'cases');
console.log('\n## Trace vs the LibreOffice oracle: real Excel files');
node('eval', 'poi-cases');
const poiFiles = (ext) => readdirSync(join(WORK, 'poi/test-data/spreadsheet')).filter((f) => f.toLowerCase().endsWith(ext)).map((f) => `poi/test-data/spreadsheet/${f}`);
console.log('\n## Number formats: cells as shown, vs LibreOffice');
node('format-check', 'poi-cases', 'cases');
console.log('\n## Readers vs LibreOffice: .xls');
node('reader-check', ...poiFiles('.xls'));
console.log('\n## Readers vs LibreOffice: .xlsb');
node('reader-check', ...poiFiles('.xlsb'));
console.log('\n## Readers vs LibreOffice: .ods (each .xlsx converted to .ods by LibreOffice)');
node('reader-check', '--via-ods', ...poiFiles('.xlsx'));
console.log('\n## Scale');
node('scale', 'chain20k', 'grid30k', 'grid100k');
