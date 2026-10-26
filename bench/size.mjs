// Bundle size per entry point (minified, gzipped). Usage: npm run size [-- --check]
// With --check, fails when any gzipped size is over its budget below (CI runs this).
// Also imported by scripts/docs-check.mjs, which checks the docs quote these numbers.
import { build } from 'esbuild';
import { gzipSync } from 'node:zlib';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const SRC = join(dirname(fileURLToPath(import.meta.url)), '../packages/spreadiff/src');
const entries = {
  'spreadiff/html (diff + CSV + HTML/text renderer)': 'html/index.ts',
  'spreadiff/core (diff + CSV + XLSX + trace)': 'core/index.ts',
  'spreadiff (core + canvas DiffGrid)': 'index.ts',
};
// What a bundle costs for a few common imports (bundlers drop the rest).
const imports = {
  'import { diffTables, parseCsv }': "export { diffTables, parseCsv } from './core/index.ts';",
  'import { parseXlsx }': "export { parseXlsx } from './core/xlsx.ts';",
  'import { parseFile } (every reader)': "export { parseFile } from './core/file.ts';",
  'import { DiffGrid, diffTables }': "export { DiffGrid } from './index.ts'; export { diffTables } from './core/index.ts';",
  'import { SpreadsheetDiff }': "export { SpreadsheetDiff } from './index.ts';",
};

// Gzipped ceilings in KB, about 10% above the sizes when they were set. Raising
// one is a decision: say why in the pull request.
const budgets = {
  'spreadiff/html (diff + CSV + HTML/text renderer)': 22,
  'spreadiff/core (diff + CSV + XLSX + trace)': 82,
  'spreadiff (core + canvas DiffGrid)': 112,
  'import { diffTables, parseCsv }': 9,
  'import { parseXlsx }': 33,
  'import { parseFile } (every reader)': 48,
  'import { DiffGrid, diffTables }': 40,
  'import { SpreadsheetDiff }': 108,
};

async function sizeOf(input) {
  const r = await build({ ...input, bundle: true, minify: true, format: 'esm', write: false, platform: 'browser', target: 'es2022', treeShaking: true });
  const code = r.outputFiles[0].contents;
  return { min: (code.length / 1024).toFixed(1), gz: (gzipSync(code).length / 1024).toFixed(1) };
}

/** Sizes in KB, rounded to one decimal as the docs print them. */
export async function measure() {
  const rows = [];
  for (const [name, entry] of Object.entries(entries)) {
    rows.push({ name, ...(await sizeOf({ entryPoints: [join(SRC, entry)] })) });
  }
  for (const [name, contents] of Object.entries(imports)) rows.push({ name, ...(await sizeOf({ stdin: { contents, resolveDir: SRC, loader: 'ts' } })) });
  return rows;
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  const check = process.argv.includes('--check');
  const over = [];
  for (const r of await measure()) {
    const budget = budgets[r.name];
    if (budget !== undefined && Number(r.gz) > budget) over.push(`${r.name}: ${r.gz} KB gzip, budget ${budget} KB`);
    console.log(`${r.name.padEnd(52)} ${r.min.padStart(6)} KB min  ${r.gz.padStart(5)} KB gzip${check ? `  (budget ${budget} KB)` : ''}`);
  }
  if (check && over.length) {
    console.error(`\n✗ Over budget:\n  ${over.join('\n  ')}`);
    process.exit(1);
  }
}
