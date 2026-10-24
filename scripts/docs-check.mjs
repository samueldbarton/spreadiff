#!/usr/bin/env node
// Documentation checks. Run with `npm run docs:check`. Fails (exit 1) when:
//   1. a ```ts / ```tsx block in the docs or in a JSDoc comment doesn't typecheck,
//   2. a block tagged <!-- example: id --> doesn't print its <!-- output: id --> block,
//      or a block tagged <!-- excerpt-of: id --> has a line that output doesn't,
//   3. a public export has no entry (### heading) in docs/api.md,
//   4. a relative link or #anchor doesn't resolve,
//   5. a public export or interface member has no JSDoc,
//   6. a bundle size ("12.3 KB") doesn't match what `npm run size` measures now.
//
// Snippets are type-checked as modules against the library source. Free
// variables they use (`before`, `container`, `diff`, …) are declared once in
// docs/snippets.d.ts so the examples stay short.
import { build } from 'esbuild';
import { execFileSync } from 'node:child_process';
import { existsSync, mkdirSync, readdirSync, readFileSync, rmSync, statSync, writeFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { dirname, join, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { listExports } from './exports.mjs';
import { measure } from '../bench/size.mjs';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const LIB = join(ROOT, 'packages/spreadiff/src');
const LITE = join(ROOT, 'packages/spreadiff-lite/src');
const OUT = join(ROOT, 'node_modules/.docs-check');
const problems = [];
const fail = (where, msg) => problems.push(`${where}: ${msg}`);

// ---- Collect documents ---------------------------------------------------
const walk = (dir) =>
  readdirSync(dir).flatMap((f) => {
    const p = join(dir, f);
    return statSync(p).isDirectory() ? walk(p) : [p];
  });
const markdown = [
  'README.md',
  'CONTRIBUTING.md',
  'CHANGELOG.md',
  'packages/spreadiff/README.md',
  'packages/spreadiff-lite/README.md',
  ...walk(join(ROOT, 'docs')).filter((f) => f.endsWith('.md')).map((f) => relative(ROOT, f)),
].filter((f) => existsSync(join(ROOT, f)));

/** Fenced code blocks with their language, line, and the HTML comment just above them. */
function blocks(file, text) {
  const out = [];
  const lines = text.split('\n');
  for (let i = 0; i < lines.length; i++) {
    const m = /^(\s*)```(\w+)?(?: ([\w-]+))?\s*$/.exec(lines[i]);
    if (!m) continue;
    const start = i;
    const body = [];
    for (i++; i < lines.length && !/^\s*```\s*$/.test(lines[i]); i++) body.push(lines[i].slice(m[1].length));
    let tag = null;
    for (let k = start - 1; k >= 0 && k >= start - 3; k--) {
      const t = /<!--\s*(example|output|excerpt-of|no-check):?\s*([\w-]*)\s*-->/.exec(lines[k]);
      if (t) {
        tag = { kind: t[1], id: t[2] };
        break;
      }
      if (lines[k].trim()) break;
    }
    // ```ts signature blocks show signatures, not runnable code: not type-checked.
    if (m[3] === 'signature') tag = { kind: 'no-check', id: '' };
    out.push({ file, line: start + 1, lang: m[2] ?? '', code: body.join('\n'), tag });
  }
  return out;
}

const docs = markdown.map((f) => ({ file: f, text: readFileSync(join(ROOT, f), 'utf8') }));
const allBlocks = docs.flatMap((d) => blocks(d.file, d.text));

// JSDoc examples in the library source.
for (const f of [...walk(LIB), ...walk(LITE)].filter((f) => /\.tsx?$/.test(f))) {
  const src = readFileSync(f, 'utf8');
  const re = /\/\*\*([\s\S]*?)\*\//g;
  let m;
  while ((m = re.exec(src))) {
    const comment = m[1].split('\n').map((l) => l.replace(/^\s*\* ?/, '')).join('\n');
    const line = src.slice(0, m.index).split('\n').length;
    for (const b of blocks(relative(ROOT, f), comment)) allBlocks.push({ ...b, line: line + b.line - 1, jsdoc: true });
  }
}

// ---- 1. Typecheck snippets ----------------------------------------------
rmSync(OUT, { recursive: true, force: true });
mkdirSync(OUT, { recursive: true });
const snippets = allBlocks.filter((b) => (b.lang === 'ts' || b.lang === 'tsx') && b.tag?.kind !== 'no-check');
const index = [];
snippets.forEach((b, i) => {
  const name = `s${i}.${b.lang}`;
  // Modules, so top-level names don't collide across snippets.
  writeFileSync(join(OUT, name), `${b.code}\nexport {};\n`);
  index.push({ name, b });
});
writeFileSync(
  join(OUT, 'tsconfig.json'),
  JSON.stringify({
    compilerOptions: {
      target: 'ES2022', module: 'ESNext', moduleResolution: 'Bundler', lib: ['ES2022', 'DOM', 'DOM.Iterable'],
      jsx: 'react-jsx', strict: true, noEmit: true, skipLibCheck: true, isolatedModules: true, noUnusedLocals: false,
      types: ['node'], typeRoots: [join(ROOT, 'node_modules/@types')],
      paths: {
        spreadiff: [join(LIB, 'index.ts')],
        'spreadiff/core': [join(LIB, 'core/index.ts')],
        'spreadiff/html': [join(LIB, 'html/index.ts')],
        'spreadiff/xlsx': [join(LIB, 'core/xlsx.ts')],
        'spreadiff/react': [join(LIB, 'react/index.tsx')],
        'spreadiff/worker': [join(LIB, 'worker/index.ts')],
        'spreadiff-lite': [join(LITE, 'index.ts')],
        react: [join(ROOT, 'node_modules/@types/react/index.d.ts')],
        'react/jsx-runtime': [join(ROOT, 'node_modules/@types/react/jsx-runtime.d.ts')],
      },
    },
    files: [join(ROOT, 'docs/snippets.d.ts'), ...index.map((x) => join(OUT, x.name))],
  }),
);
try {
  execFileSync(join(ROOT, 'node_modules/.bin/tsc'), ['-p', join(OUT, 'tsconfig.json')], { encoding: 'utf8' });
} catch (e) {
  for (const line of String(e.stdout).split('\n')) {
    const m = /(s\d+\.tsx?)\((\d+),(\d+)\): (.*)/.exec(line);
    if (!m) continue;
    const { b } = index.find((x) => x.name === m[1]);
    fail(`${b.file}:${b.line + Number(m[2])}`, m[4]);
  }
}

// ---- 2. Run examples, compare output ------------------------------------
const outputs = new Map(allBlocks.filter((b) => b.tag?.kind === 'output').map((b) => [b.tag.id, b]));
for (const b of allBlocks.filter((b) => b.tag?.kind === 'example')) {
  const expected = outputs.get(b.tag.id);
  if (!expected) {
    fail(`${b.file}:${b.line}`, `example "${b.tag.id}" has no <!-- output: ${b.tag.id} --> block`);
    continue;
  }
  const res = await build({
    stdin: { contents: b.code, loader: b.lang === 'tsx' ? 'tsx' : 'ts', resolveDir: ROOT },
    bundle: true, write: false, format: 'cjs', platform: 'node', logLevel: 'silent',
    alias: {
      spreadiff: join(LIB, 'index.ts'), 'spreadiff/core': join(LIB, 'core/index.ts'), 'spreadiff/html': join(LIB, 'html/index.ts'), 'spreadiff-lite': join(LITE, 'index.ts'),
    },
  });
  const printed = [];
  const log = (...a) => printed.push(a.map((x) => (typeof x === 'string' ? x : JSON.stringify(x))).join(' '));
  new Function('console', 'require', res.outputFiles[0].text)({ log }, createRequire(import.meta.url));
  const got = printed.join('\n').trimEnd();
  if (got !== expected.code.trimEnd()) {
    fail(`${expected.file}:${expected.line}`, `output of example "${b.tag.id}" differs.\n--- documented\n${expected.code}\n--- actual\n${got}`);
  }
  // Excerpts elsewhere (a README) must quote the real output, line for line.
  const lines = new Set(got.split('\n'));
  for (const ex of allBlocks.filter((x) => x.tag?.kind === 'excerpt-of' && x.tag.id === b.tag.id)) {
    for (const line of ex.code.split('\n')) {
      if (line.trim() && !lines.has(line)) fail(`${ex.file}:${ex.line}`, `excerpt of "${b.tag.id}" has a line its output doesn't: ${line}`);
    }
  }
}
for (const ex of allBlocks.filter((x) => x.tag?.kind === 'excerpt-of' && !allBlocks.some((b) => b.tag?.kind === 'example' && b.tag.id === x.tag.id))) {
  fail(`${ex.file}:${ex.line}`, `excerpt of "${ex.tag.id}", which isn't an example`);
}

// ---- 3. API reference covers every export ---------------------------------
const api = docs.find((d) => d.file === 'docs/api.md');
if (!api) fail('docs/api.md', 'missing');
else {
  const headings = new Set([...api.text.matchAll(/^#{2,4} `?([\w.]+)/gm)].map((m) => m[1]));
  const seen = new Set();
  for (const e of listExports()) {
    if (seen.has(e.name)) continue;
    seen.add(e.name);
    if (!headings.has(e.name)) fail('docs/api.md', `no entry for export \`${e.name}\` (${e.entry})`);
  }
}

// ---- 4. Links and anchors --------------------------------------------------
const slug = (h) =>
  h.trim().toLowerCase().replace(/<[^>]+>/g, '').replace(/[^\p{L}\p{N}\s_-]/gu, '').replace(/\s/g, '-');
const anchors = (text) => {
  const counts = new Map();
  const out = new Set();
  let inFence = false;
  for (const line of text.split('\n')) {
    if (/^\s*```/.test(line)) inFence = !inFence;
    const m = !inFence && /^#{1,6}\s+(.*)$/.exec(line);
    if (!m) continue;
    const base = slug(m[1].replace(/`/g, ''));
    const n = counts.get(base) ?? 0;
    counts.set(base, n + 1);
    out.add(n ? `${base}-${n}` : base);
  }
  return out;
};
for (const d of docs) {
  const noFences = d.text.replace(/```[\s\S]*?```/g, '');
  for (const m of noFences.matchAll(/\]\(([^)\s]+)\)/g)) {
    const href = m[1];
    if (/^(https?:|mailto:)/.test(href)) continue;
    const [path, hash] = href.split('#');
    const target = path ? resolve(dirname(join(ROOT, d.file)), path) : join(ROOT, d.file);
    if (!existsSync(target)) {
      fail(d.file, `broken link ${href}`);
      continue;
    }
    if (hash && target.endsWith('.md') && !anchors(readFileSync(target, 'utf8')).has(hash)) fail(d.file, `missing anchor ${href}`);
  }
}

// ---- 5. JSDoc coverage -------------------------------------------------------
for (const e of listExports()) {
  if (!e.documented) fail(e.file, `export \`${e.name}\` has no JSDoc`);
  for (const m of e.members) if (!m.documented) fail(e.file, `\`${e.name}.${m.name}\` has no JSDoc`);
}

// ---- 6. Bundle sizes ----------------------------------------------------------
// Any one-decimal KB figure in the docs must be a current measurement.
const sizes = new Set((await measure()).flatMap((r) => [r.min, r.gz]));
for (const d of docs) {
  if (d.file === 'CHANGELOG.md') continue; // history quotes old sizes on purpose
  for (const m of d.text.matchAll(/(\d+\.\d) KB/g)) {
    if (!sizes.has(m[1])) fail(d.file, `"${m[0]}" isn't a current bundle size (measured: ${[...sizes].join(', ')} KB); run \`npm run size\``);
  }
}

// ---- Report ----------------------------------------------------------------
const unique = [...new Set(problems)];
if (unique.length) {
  console.error(unique.join('\n\n'));
  console.error(`\n✗ ${unique.length} documentation problem${unique.length === 1 ? '' : 's'}`);
  process.exit(1);
}
console.log(
  `✓ docs: ${markdown.length} files, ${snippets.length} code blocks typechecked, ${allBlocks.filter((b) => b.tag?.kind === 'example').length} examples run, links and API coverage OK`,
);
