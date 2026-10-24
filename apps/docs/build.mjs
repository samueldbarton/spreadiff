#!/usr/bin/env node
// Builds the documentation site into apps/docs/dist from the repository's
// Markdown (the same files `npm run docs:check` verifies):
//   - one static HTML page per document, highlighted at build time;
//   - a search index, llms.txt / llms-full.txt and each page's Markdown;
//   - a small client script (search, theme, contents, copy buttons) and,
//     loaded only on pages that have one, the live demos.
// `node build.mjs --serve` rebuilds on change and serves on :4321.
import { build as esbuild } from 'esbuild';
import { cpSync, existsSync, mkdirSync, readFileSync, rmSync, watch, writeFileSync } from 'node:fs';
import { createServer } from 'node:http';
import { dirname, extname, join, normalize, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { Marked } from 'marked';
import { createHighlighter } from 'shiki';
import { BRANCH, NAV, REPO } from './nav.mjs';

const HERE = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(HERE, '../..');
const OUT = join(HERE, 'dist');
const LIB = join(ROOT, 'packages/spreadiff/src');

const LANGS = ['ts', 'tsx', 'js', 'bash', 'css', 'json', 'html', 'markdown', 'text'];
const highlighter = await createHighlighter({ themes: ['github-light', 'github-dark'], langs: LANGS.filter((l) => l !== 'text') });

/** GitHub's heading slug, as docs-check computes it. */
const slugify = (h) =>
  h.trim().toLowerCase().replace(/<[^>]+>/g, '').replace(/[^\p{L}\p{N}\s_-]/gu, '').replace(/\s/g, '-');
const escape = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const plain = (md) =>
  md
    .replace(/```[\s\S]*?```/g, ' ')
    .replace(/<!--[\s\S]*?-->/g, ' ')
    .replace(/!\[[^\]]*\]\([^)]*\)/g, ' ')
    .replace(/\[([^\]]*)\]\([^)]*\)/g, '$1')
    .replace(/[`*>#|]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

const pages = NAV.flatMap((g) => g.pages.map((p) => ({ ...p, group: g.title })));
const byFile = new Map(pages.map((p) => [p.file, p]));

/** Resolve a link in `fromFile` to a page of the site, a copied asset, or GitHub. */
function rewriteLink(href, fromFile) {
  if (/^(https?:|mailto:|#)/.test(href)) return href;
  const [path, hash] = href.split('#');
  const target = normalize(join(dirname(fromFile), path)).replace(/\\/g, '/');
  const page = byFile.get(target) ?? (target === 'README.md' ? { slug: 'index' } : null);
  if (page) return `${page.slug}.html${hash ? `#${hash}` : ''}`;
  if (target.startsWith('docs/images/')) return target.slice('docs/'.length);
  return `${REPO}/blob/${BRANCH}/${target}${hash ? `#${hash}` : ''}`;
}

function highlight(code, lang) {
  const l = LANGS.includes(lang) ? lang : lang === 'sh' || lang === 'shell' ? 'bash' : 'text';
  if (l === 'text') return `<pre class="shiki"><code>${escape(code)}</code></pre>`;
  return highlighter.codeToHtml(code, { lang: l, themes: { light: 'github-light', dark: 'github-dark' }, defaultColor: false });
}

/** Render one page's Markdown. Returns HTML and its headings (for contents and search). */
function render(md, file) {
  const headings = [];
  const counts = new Map();
  let pending = null; // an HTML comment tag that labels the next code block
  const marked = new Marked({ gfm: true });
  marked.use({
    renderer: {
      heading({ tokens, depth, text }) {
        const base = slugify(text.replace(/`/g, ''));
        const n = counts.get(base) ?? 0;
        counts.set(base, n + 1);
        const id = n ? `${base}-${n}` : base;
        const inner = this.parser.parseInline(tokens);
        if (depth >= 2 && depth <= 3) headings.push({ id, depth, text: plain(text) });
        if (depth === 1) return `<h1 id="${id}">${inner}</h1>\n`;
        return `<h${depth} id="${id}"><a class="anchor" href="#${id}" aria-label="Link to this section">#</a>${inner}</h${depth}>\n`;
      },
      code({ text, lang }) {
        const [l = '', flag] = (lang ?? '').split(/\s+/);
        const tag = pending;
        pending = null;
        let badge = '';
        if (tag?.kind === 'output') badge = '<span class="badge ok" title="npm run docs:check runs the example above and compares its output with this">Verified output</span>';
        else if (tag?.kind === 'excerpt-of') badge = '<span class="badge ok" title="Every line is checked against the real output">Verified excerpt</span>';
        else if (tag?.kind === 'example') badge = '<span class="badge ok" title="npm run docs:check runs this">Runs in CI</span>';
        else if ((l === 'ts' || l === 'tsx') && flag !== 'signature' && tag?.kind !== 'no-check') badge = '<span class="badge" title="npm run docs:check typechecks every TypeScript block against the library">Typechecked</span>';
        const label = flag === 'signature' ? 'Signature' : tag?.kind === 'output' ? 'Output' : { ts: 'TypeScript', tsx: 'TSX', bash: 'Terminal', css: 'CSS', json: 'JSON', markdown: 'Markdown', html: 'HTML', text: '', js: 'JavaScript' }[l] ?? l;
        return `<figure class="code${flag === 'signature' ? ' signature' : ''}"><figcaption><span>${escape(label)}</span>${badge}<button class="copy" type="button" aria-label="Copy code">Copy</button></figcaption>${highlight(text.replace(/\n$/, ''), l)}</figure>\n`;
      },
      html({ text }) {
        const demo = /<!--\s*demo:\s*([\w-]+)\s*-->/.exec(text);
        if (demo) return `<div class="demo" data-demo="${demo[1]}"><div class="demo-loading">Loading the live demo…</div></div>\n`;
        const tag = /<!--\s*(example|output|excerpt-of|no-check):?\s*([\w-]*)\s*-->/.exec(text);
        if (tag) pending = { kind: tag[1], id: tag[2] };
        return /^\s*<!--/.test(text) ? '' : text;
      },
      link({ href, title, tokens }) {
        const h = rewriteLink(href, file);
        const external = /^https?:/.test(h) && !h.startsWith(REPO);
        return `<a href="${escape(h)}"${title ? ` title="${escape(title)}"` : ''}${external ? ' rel="noopener"' : ''}>${this.parser.parseInline(tokens)}</a>`;
      },
      image({ href, title, text }) {
        return `<img src="${escape(rewriteLink(href, file))}" alt="${escape(text)}"${title ? ` title="${escape(title)}"` : ''} loading="lazy">`;
      },
      table(token) {
        const head = token.header.map((c, i) => `<th${token.align[i] ? ` style="text-align:${token.align[i]}"` : ''}>${this.parser.parseInline(c.tokens)}</th>`).join('');
        const body = token.rows.map((r) => `<tr>${r.map((c, i) => `<td${token.align[i] ? ` style="text-align:${token.align[i]}"` : ''}>${this.parser.parseInline(c.tokens)}</td>`).join('')}</tr>`).join('\n');
        return `<div class="table"><table><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table></div>\n`;
      },
    },
  });
  return { html: marked.parse(md), headings };
}

function layout({ title, description, slug, body, toc, prev, next, file, demos }) {
  const nav = NAV.map(
    (g) =>
      `<div class="nav-group"><p class="nav-title">${escape(g.title)}</p><ul>${g.pages
        .map((n) => pages.find((p) => p.slug === n.slug))
        .map((p) => `<li><a href="${p.slug}.html"${p.slug === slug ? ' aria-current="page"' : ''}>${escape(p.navTitle)}</a></li>`)
        .join('')}</ul></div>`,
  ).join('');
  const tocHtml = toc.length
    ? `<nav class="toc" aria-label="On this page"><p class="nav-title">On this page</p><ul>${toc
        .map((h) => `<li class="d${h.depth}"><a href="#${h.id}">${escape(h.text)}</a></li>`)
        .join('')}</ul></nav>`
    : '';
  const pager = `<nav class="pager" aria-label="Pages">${prev ? `<a class="prev" href="${prev.slug}.html"><span>Previous</span>${escape(prev.navTitle)}</a>` : '<span></span>'}${next ? `<a class="next" href="${next.slug}.html"><span>Next</span>${escape(next.navTitle)}</a>` : ''}</nav>`;
  return shell({
    title: `${title} · spreadiff`,
    description,
    demos,
    body: `<div class="layout">
<aside class="sidebar" id="sidebar" aria-label="Documentation">${nav}</aside>
<main class="content" id="main">
<div class="page-actions">
  <span class="crumb">${escape(pages.find((p) => p.slug === slug)?.group ?? '')}</span>
  <span class="actions"><button type="button" class="ghost" data-copy-page="${slug}.md">Copy page</button><a class="ghost" href="${slug}.md">Markdown</a><a class="ghost" href="${REPO}/edit/${BRANCH}/${file}">Edit on GitHub</a></span>
</div>
<article class="prose">${body}</article>
${pager}
</main>
<div class="aside">${tocHtml}</div>
</div>`,
  });
}

function shell({ title, description, body, demos, home = false }) {
  return `<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${escape(title)}</title>
<meta name="description" content="${escape(description)}">
<meta property="og:title" content="${escape(title)}">
<meta property="og:description" content="${escape(description)}">
<link rel="icon" href="favicon.svg" type="image/svg+xml">
<link rel="stylesheet" href="assets/site.css">
<script>try{var t=localStorage.getItem('sd-docs-theme');if(t==='light'||t==='dark')document.documentElement.dataset.theme=t}catch(e){}</script>
<script type="module" src="assets/client.js"></script>
${demos ? '<link rel="modulepreload" href="assets/demos.js">' : ''}
</head>
<body${home ? ' class="home"' : ''}>
<a class="skip" href="#main">Skip to content</a>
<header class="topbar">
  <button class="menu" type="button" aria-label="Menu" aria-controls="sidebar" aria-expanded="false"><svg viewBox="0 0 24 24" width="20" height="20" aria-hidden="true"><path d="M4 7h16M4 12h16M4 17h16" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg></button>
  <a class="brand" href="index.html"><svg viewBox="0 0 24 24" width="22" height="22" aria-hidden="true"><rect x="2" y="3" width="20" height="18" rx="4" fill="var(--panel)" stroke="var(--line-strong)"/><rect x="5" y="7" width="6" height="3" rx="1" fill="var(--removed)"/><rect x="13" y="7" width="6" height="3" rx="1" fill="var(--added)"/><rect x="5" y="12" width="14" height="3" rx="1" fill="var(--computed)"/></svg>spreadiff</a>
  <nav class="top-links" aria-label="Primary"><a href="getting-started.html">Docs</a><a href="spreadiff-lite.html">spreadiff-lite</a><a href="api.html">API</a></nav>
  <button class="search-button" type="button" data-search aria-label="Search the docs"><svg viewBox="0 0 24 24" width="16" height="16" aria-hidden="true"><circle cx="11" cy="11" r="7" fill="none" stroke="currentColor" stroke-width="2"/><path d="m20 20-4-4" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg><span>Search</span><kbd>⌘K</kbd></button>
  <button class="theme" type="button" data-theme-toggle aria-label="Theme: system"><svg class="i-system" viewBox="0 0 24 24" width="18" height="18" aria-hidden="true"><rect x="3" y="4" width="18" height="12" rx="2" fill="none" stroke="currentColor" stroke-width="2"/><path d="M8 20h8M12 16v4" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg><svg class="i-light" viewBox="0 0 24 24" width="18" height="18" aria-hidden="true"><circle cx="12" cy="12" r="4" fill="none" stroke="currentColor" stroke-width="2"/><path d="M12 2v2M12 20v2M4.9 4.9l1.4 1.4M17.7 17.7l1.4 1.4M2 12h2M20 12h2M4.9 19.1l1.4-1.4M17.7 6.3l1.4-1.4" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg><svg class="i-dark" viewBox="0 0 24 24" width="18" height="18" aria-hidden="true"><path d="M20 14.5A8 8 0 0 1 9.5 4a8 8 0 1 0 10.5 10.5Z" fill="none" stroke="currentColor" stroke-width="2" stroke-linejoin="round"/></svg></button>
  <a class="gh" href="${REPO}" aria-label="GitHub repository"><svg viewBox="0 0 16 16" width="20" height="20" aria-hidden="true"><path fill="currentColor" d="M8 0a8 8 0 0 0-2.53 15.59c.4.07.55-.17.55-.38v-1.33c-2.23.48-2.7-1.07-2.7-1.07-.36-.92-.89-1.17-.89-1.17-.73-.5.06-.49.06-.49.8.06 1.23.83 1.23.83.72 1.22 1.88.87 2.34.66.07-.52.28-.87.5-1.07-1.78-.2-3.65-.89-3.65-3.95 0-.87.31-1.59.82-2.15-.08-.2-.36-1.02.08-2.12 0 0 .67-.21 2.2.82a7.6 7.6 0 0 1 4 0c1.53-1.04 2.2-.82 2.2-.82.44 1.1.16 1.92.08 2.12.51.56.82 1.28.82 2.15 0 3.07-1.87 3.75-3.65 3.95.29.25.54.73.54 1.48v2.2c0 .21.15.46.55.38A8 8 0 0 0 8 0Z"/></svg></a>
</header>
${body}
<footer class="footer"><span>MIT licensed.</span><span>Every code block on this site is typechecked, and every example's output verified, by <code>npm run docs:check</code>.</span><a href="llms.txt">llms.txt</a></footer>
</body>
</html>
`;
}

function landing(snippets) {
  const card = (href, title, text, cta) => `<a class="card" href="${href}"><h3>${title}</h3><p>${text}</p><span class="cta">${cta} →</span></a>`;
  const feature = (title, text) => `<div class="feature"><h3>${title}</h3><p>${text}</p></div>`;
  return shell({
    home: true,
    demos: true,
    title: 'spreadiff: an elegant but fully featured spreadsheet diff and UI library',
    description: 'Diff CSV and Excel files by rows and columns, trace every changed formula result to the edits that caused it, and show it in a canvas grid or as Markdown.',
    body: `<main id="main">
<section class="hero">
  <p class="eyebrow">Spreadsheet diffs, explained</p>
  <h1>An elegant but fully featured<br>spreadsheet diff and UI library.</h1>
  <p class="lede">See what changed between two spreadsheets, and why: every edit, every result it moved, and how much each edit contributed, recalculated and checked against LibreOffice.</p>
  <div class="hero-actions"><a class="button primary" href="getting-started.html">Get started</a><a class="button" href="spreadiff-lite.html">spreadiff-lite for agents</a><a class="button ghost" href="${REPO}">GitHub</a></div>
</section>
<section class="hero-demo">
  <div class="demo" data-demo="hero"><div class="demo-loading">Loading the live demo…</div></div>
  <p class="caption">Live: two versions of a small financial model. Amber cells moved because of an edit upstream. Hover one to see which edits drove it, and by how much.</p>
</section>
<section class="packages">
  ${card('grid.html', 'spreadiff', 'The UI. A canvas grid that scrolls a million rows, split and unified layouts, deltas, annotations, a React component, static HTML for email and pull requests, and a Web Worker.', 'The DiffGrid')}
  ${card('spreadiff-lite.html', 'spreadiff-lite', 'No UI. A Markdown or JSON report of every edit, its before and after values, what the sheet calls the cell, and what it moved downstream. A CLI and an API for agents and CI.', 'The report')}
</section>
<section class="stack-section">
  <div class="stack-text"><h2>Built for the end of an agent's turn</h2><p>An agent edits a model and says what it did. spreadiff-lite says what that <em>moved</em>: each final output with the edits that drove it, by recalculating the workbook. Values from Bloomberg, Capital IQ or FactSet are taken as saved and reported as inputs.</p><p><a href="spreadiff-lite.html#with-an-agent">Use it with an agent →</a></p></div>
  <div>${snippets.lite}</div>
</section>
<section class="features">
  ${feature('Understands tables', 'Columns match by name; renamed columns by content. Rows match by id, or by sequence alignment. An edited row is <em>modified</em>; a reordered one, <em>moved</em>.')}
  ${feature('Explains formulas', 'Every changed cell is an input, a formula edit or a downstream result. Each result is traced to its causes by undoing each edit and recalculating.')}
  ${feature('Verified', 'Causes and contributions match LibreOffice on every scored cell of the test corpus: 3,400 of 3,400 contributions. The checks rerun in CI.')}
  ${feature('Fast', 'Only visible cells are drawn, so a million-row diff scrolls like a ten-row one. Parsing, diffing and tracing run in a worker.')}
  ${feature('Accessible', 'Screen readers get an ARIA grid of the visible cells. Colours meet WCAG AA in light and dark, checked by axe-core and a canvas contrast test.')}
  ${feature('Runs anywhere', 'No server: nothing leaves the browser. Works under a strict Content-Security-Policy. The HTML renderer and spreadiff-lite run in Node and on the edge.')}
</section>
<section class="split-section">
  <div><h2>Three lines to a diff</h2><p>Parse, diff, render. Rows match by an id column when there is one. Give the container a height and the grid fills it.</p><p><a href="getting-started.html">Getting started →</a></p></div>
  <div>${snippets.quick}</div>
</section>
</main>`,
  });
}

async function buildSite() {
  rmSync(OUT, { recursive: true, force: true });
  mkdirSync(join(OUT, 'assets'), { recursive: true });
  const search = [];
  const llms = [];
  const full = [];
  pages.forEach((p, i) => {
    const md = readFileSync(join(ROOT, p.file), 'utf8');
    const h1 = /^#\s+(.+)$/m.exec(md)?.[1] ?? p.slug;
    p.title ??= plain(h1);
    p.navTitle = p.title;
    const firstPara = md.replace(/^#.*$/m, '').split(/\n\s*\n/).map((x) => x.trim()).find((x) => x && !/^[#|`<!\-*]/.test(x)) ?? '';
    p.description = plain(firstPara).slice(0, 200);
    p.md = md;
  });
  pages.forEach((p, i) => {
    const { html, headings } = render(p.md, p.file);
    const demos = /data-demo=/.test(html);
    writeFileSync(
      join(OUT, `${p.slug}.html`),
      layout({ title: p.title, description: p.description, slug: p.slug, body: html, toc: headings, prev: pages[i - 1], next: pages[i + 1], file: p.file, demos }),
    );
    writeFileSync(join(OUT, `${p.slug}.md`), p.md);
    // Search: the page, then each section with its text.
    const sections = p.md.split(/^(?=#{2,3}\s)/m);
    search.push({ t: p.title, u: `${p.slug}.html`, g: p.group, x: plain(sections[0]).slice(0, 300) });
    sections.slice(1).forEach((sec, k) => {
      const head = /^#{2,3}\s+(.+)$/m.exec(sec)?.[1] ?? '';
      const h = headings.find((x) => x.text === plain(head)) ?? headings[k];
      search.push({ t: plain(head), p: p.title, u: `${p.slug}.html#${h?.id ?? ''}`, g: p.group, x: plain(sec.replace(/^.*$/m, '')).slice(0, 300) });
    });
    llms.push(`- [${p.title}](${p.slug}.md): ${p.description}`);
    full.push(`<!-- ${p.file} -->\n\n${p.md.trim()}\n`);
  });
  writeFileSync(join(OUT, 'search.json'), JSON.stringify(search));
  writeFileSync(
    join(OUT, 'llms.txt'),
    `# spreadiff\n\n> An elegant but fully featured spreadsheet diff and UI library. spreadiff diffs CSV and Excel files by rows and columns, separates edited inputs from formula results that moved downstream, and traces each result to its causes. spreadiff-lite reports the same as Markdown or JSON, for agents and CI.\n\n## Docs\n\n${llms.join('\n')}\n\n## Optional\n\n- [Everything in one file](llms-full.txt)\n`,
  );
  writeFileSync(join(OUT, 'llms-full.txt'), full.join('\n\n'));

  // Landing page, with snippets rendered like any page.
  const snip = (md) => render(md, 'README.md').html;
  const liteOut = /<!-- output: lite-regions -->\n```text\n([\s\S]*?)```/.exec(readFileSync(join(ROOT, 'docs/spreadiff-lite.md'), 'utf8'))[1];
  // The final outputs of the guide's verified example, rendered as the report renders.
  const lines = liteOut.split('\n');
  const at = lines.indexOf('### Final outputs');
  const excerpt = [lines.find((l) => l.startsWith('**')), '', ...lines.slice(at + 2, lines.indexOf('', at + 2))];
  writeFileSync(
    join(OUT, 'index.html'),
    landing({
      lite: `<div class="report-card"><div class="report-head"><span>spreadiff-lite regions-v1.xlsx regions-v2.xlsx</span><span class="badge ok">Verified output</span></div><div class="prose">${snip(excerpt.join('\n'))}</div></div>`,
      quick: snip("```ts\nimport { DiffGrid, diffTables, parseCsv } from 'spreadiff';\n\nconst diff = diffTables(parseCsv(before), parseCsv(after));\nnew DiffGrid({ layout: 'split', cellDisplay: 'delta' }).render({ diff, container });\n```\n"),
    }),
  );

  // Assets.
  cpSync(join(ROOT, 'docs/images'), join(OUT, 'images'), { recursive: true });
  cpSync(join(HERE, 'src/favicon.svg'), join(OUT, 'favicon.svg'));
  const fonts = join(OUT, 'assets/fonts');
  mkdirSync(fonts, { recursive: true });
  const font = (pkg, file) => cpSync(join(ROOT, 'node_modules', pkg, 'files', file), join(fonts, file));
  font('@fontsource-variable/inter', 'inter-latin-wght-normal.woff2');
  font('@fontsource-variable/jetbrains-mono', 'jetbrains-mono-latin-wght-normal.woff2');
  cpSync(join(HERE, 'src/site.css'), join(OUT, 'assets/site.css'));
  await esbuild({
    entryPoints: { client: join(HERE, 'src/client.ts'), demos: join(HERE, 'src/demos.ts') },
    outdir: join(OUT, 'assets'),
    bundle: true,
    format: 'esm',
    splitting: true,
    minify: true,
    target: 'es2022',
    logLevel: 'error',
    alias: {
      spreadiff: join(LIB, 'index.ts'),
      'spreadiff/core': join(LIB, 'core/index.ts'),
      'spreadiff/html': join(LIB, 'html/index.ts'),
      'spreadiff-lite': join(ROOT, 'packages/spreadiff-lite/src/index.ts'),
    },
  });
  checkLinks();
  return pages.length + 1;
}

/** Every link between pages, and every #anchor, must resolve in the built site. */
function checkLinks() {
  const problems = [];
  const ids = new Map();
  const htmlFiles = ['index', ...pages.map((p) => p.slug)].map((s) => `${s}.html`);
  for (const f of htmlFiles) ids.set(f, new Set([...readFileSync(join(OUT, f), 'utf8').matchAll(/\sid="([^"]+)"/g)].map((m) => m[1])));
  for (const f of htmlFiles) {
    const html = readFileSync(join(OUT, f), 'utf8');
    for (const m of html.matchAll(/\s(?:href|src)="([^"]+)"/g)) {
      const url = m[1].replace(/&amp;/g, '&');
      if (/^(https?:|mailto:|data:)/.test(url)) continue;
      const [path, hash] = url.split('#');
      const target = path || f;
      if (!existsSync(join(OUT, decodeURIComponent(target)))) problems.push(`${f}: broken link ${url}`);
      else if (hash && ids.has(target) && !ids.get(target).has(decodeURIComponent(hash))) problems.push(`${f}: missing anchor ${url}`);
    }
  }
  if (problems.length) throw new Error(`docs site has ${problems.length} broken link${problems.length === 1 ? '' : 's'}:\n${[...new Set(problems)].join('\n')}`);
}

const n = await buildSite();
console.log(`✓ docs site: ${n} pages → ${relative(ROOT, OUT)}`);

if (process.argv.includes('--serve')) {
  const types = { '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css', '.json': 'application/json', '.svg': 'image/svg+xml', '.png': 'image/png', '.woff2': 'font/woff2', '.md': 'text/markdown; charset=utf-8', '.txt': 'text/plain; charset=utf-8' };
  createServer((req, res) => {
    let p = decodeURIComponent(new URL(req.url, 'http://x').pathname);
    if (p.endsWith('/')) p += 'index.html';
    const file = join(OUT, normalize(p));
    if (!file.startsWith(OUT) || !existsSync(file)) {
      res.writeHead(404).end('Not found');
      return;
    }
    res.writeHead(200, { 'content-type': types[extname(file)] ?? 'application/octet-stream' }).end(readFileSync(file));
  }).listen(4321, () => console.log('  serving on http://localhost:4321'));
  let timer;
  for (const dir of [join(ROOT, 'docs'), join(HERE, 'src'), LIB]) {
    watch(dir, { recursive: true }, () => {
      clearTimeout(timer);
      timer = setTimeout(() => buildSite().then(() => console.log('  rebuilt')).catch((e) => console.error(e)), 150);
    });
  }
}
