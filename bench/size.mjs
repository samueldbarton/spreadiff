// Bundle size per entry point (minified, gzipped). Usage: npm run size
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

/** Sizes in KB, rounded to one decimal as the docs print them. */
export async function measure() {
  const rows = [];
  for (const [name, entry] of Object.entries(entries)) {
    const r = await build({ entryPoints: [join(SRC, entry)], bundle: true, minify: true, format: 'esm', write: false, platform: 'browser', target: 'es2022', treeShaking: true });
    const code = r.outputFiles[0].contents;
    rows.push({ name, min: (code.length / 1024).toFixed(1), gz: (gzipSync(code).length / 1024).toFixed(1) });
  }
  return rows;
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  for (const r of await measure()) console.log(`${r.name.padEnd(52)} ${r.min.padStart(6)} KB min  ${r.gz.padStart(5)} KB gzip`);
}
