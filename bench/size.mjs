// Bundle size per entry point (minified, gzipped). Usage: node bench/size.mjs
import { build } from 'esbuild';
import { gzipSync } from 'node:zlib';

const entries = {
  'spreadiff/lite (diff + CSV + HTML/text renderer)': 'packages/spreadiff/src/lite/index.ts',
  'spreadiff/core (diff + CSV + XLSX + trace)': 'packages/spreadiff/src/core/index.ts',
  'spreadiff (core + canvas DiffGrid)': 'packages/spreadiff/src/index.ts',
};
const rows = [];
for (const [name, entry] of Object.entries(entries)) {
  const r = await build({ entryPoints: [entry], bundle: true, minify: true, format: 'esm', write: false, platform: 'browser', target: 'es2022', treeShaking: true });
  const code = r.outputFiles[0].contents;
  rows.push({ name, min: code.length, gz: gzipSync(code).length });
}
for (const r of rows) console.log(`${r.name.padEnd(52)} ${(r.min / 1024).toFixed(1).padStart(6)} KB min  ${(r.gz / 1024).toFixed(1).padStart(5)} KB gzip`);
