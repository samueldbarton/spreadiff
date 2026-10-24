# Getting started

This page takes you from two files to an interactive diff in a few minutes.
For how the comparison works, see [How diffing works](diffing.md).

## Install

spreadiff isn't published to npm yet. Until it is, build it from this
repository and install the package directory:

```bash
git clone https://github.com/samueldbarton/spreadiff
cd spreadiff && npm install && npm run build -w spreadiff
cd /path/to/your/app && npm install /path/to/spreadiff/packages/spreadiff
```

For [spreadiff-lite](spreadiff-lite.md), build it too and install both
packages (it depends on the library):

```bash
cd spreadiff && npm run build -w spreadiff-lite
cd /path/to/your/app && npm install /path/to/spreadiff/packages/spreadiff /path/to/spreadiff/packages/spreadiff-lite
```

Requirements:
- a bundler that understands package `exports` (Vite, webpack 5, esbuild,
  Rollup) and an ES2022 target;
- for the grid, a browser from 2023 or later (Chrome / Edge 111, Firefox
  113, Safari 16.2). `spreadiff/html` and `spreadiff/core` also run in Node 18+;
- React 18+, only for `spreadiff/react`.

Details, including CSP and server-side rendering, are in
[Browser support](browser-support.md).

## Pick an entry point

| Import | Use it for | Size (min+gzip) |
| --- | --- | --- |
| `spreadiff/html` | Diff + CSV + an HTML/text renderer. No canvas or DOM APIs, so it runs on servers too | 16.0 KB |
| `spreadiff/core` | The diff engine, CSV and XLSX parsers, formula tracing. No rendering | 42.1 KB |
| `spreadiff` | Everything in core plus the canvas `DiffGrid` | 62.6 KB |
| `spreadiff/react` | The `DiffGrid` React component | + React |
| `spreadiff/worker` | Parse, diff and trace in a Web Worker | small client + worker |
| `spreadiff-lite` | A separate package with no UI: a Markdown or JSON report of what changed and what it moved downstream, and a CLI | Node / server |

Rule of thumb:
- To review a handful of edits, especially outside a browser, use `spreadiff/html`.
- To scroll through large sheets interactively, use `DiffGrid`.
- To tell a reader (or an agent) what changed and what it moved, in
  Markdown, use `spreadiff-lite`.

## 1. Diff two tables

Parsers turn files into `Table`s; `diffTables` compares two of them.

<!-- example: quickstart -->
```ts
import { diffTables, parseCsv } from 'spreadiff';

const before = parseCsv('id,name,price\n1,Lamp,10\n2,Chair,20\n3,Desk,30');
const after = parseCsv('id,name,price\n1,Lamp,12\n3,Desk,30\n4,Rug,50');

const diff = diffTables(before, after);
console.log(diff.keyColumns, diff.stats.rowsModified, diff.stats.rowsAdded, diff.stats.rowsRemoved);
```

<!-- output: quickstart -->
```text
["id"] 1 1 1
```

The diff matched rows by the `id` column automatically, so reordering rows
wouldn't change the result. Lamp's price changed (modified), Chair is gone
(removed) and Rug is new (added).

For files, `parseFile(name, bytes)` picks the parser from the name and
contents (CSV, TSV, or XLSX), and `diffWorkbooks` compares every sheet:

```ts
import { diffWorkbooks, parseFile } from 'spreadiff';

const result = diffWorkbooks(
  parseFile(oldFile.name, await oldFile.arrayBuffer()),
  parseFile(newFile.name, await newFile.arrayBuffer()),
);
for (const sheet of result.sheets) console.log(sheet.name, sheet.status, sheet.diff.stats);
```

## 2. Show it

### In the browser, interactively: `DiffGrid`

```ts
import { DiffGrid } from 'spreadiff';

const grid = new DiffGrid({ layout: 'split' });
grid.render({ diff, container }); // the grid fills `container`: give it a height
```

The grid draws on a canvas and only paints what's visible, so a
million-row diff scrolls as smoothly as a ten-row one. Click a cell to
select it, hover a changed cell for details, press `n` / `p` to jump
between changes. See [The DiffGrid](grid.md) for every option.

With React:

```tsx
import { DiffGrid } from 'spreadiff/react';

export const Review = () => <DiffGrid diff={diff} layout="split" style={{ height: 600 }} />;
```

### Anywhere, as HTML: `spreadiff/html`

```ts
import { HTML_CSS, renderHTML } from 'spreadiff/html';

container.innerHTML = `<style>${HTML_CSS}</style>` + renderHTML(diff, { layout: 'split' });
```

`renderHTML` returns a string, so you can produce it on a server, in a CI
job or in an email. See [spreadiff/html](html.md).

## 3. Big files: parse off the main thread

Parsing and diffing a large file takes seconds. Do it in a worker so the
page stays responsive:

```ts
import { createDiffWorker } from 'spreadiff/worker';

const differ = createDiffWorker(
  new Worker(new URL('spreadiff/worker/worker.js', import.meta.url), { type: 'module' }),
);
const { result, trace } = await differ.diffFiles(oldFile, newFile, {}, { trace: true });
```

`{ trace: true }` also traces formulas in the worker (next section), which
can take seconds on a large model. See [Workers](workers.md) for bundler
notes and what crosses the boundary.

## 4. Spreadsheets with formulas

For XLSX files, spreadiff also tells you **why** a cell changed. It
separates edited inputs from formula results that moved as a consequence,
and traces each result back to its causes:

```ts
import { DiffGrid, traceWorkbook } from 'spreadiff';

const trace = traceWorkbook(workbookDiff);
const grid = new DiffGrid({ trace: { trace, workbook: workbookDiff, sheet: 0 } });
```

Downstream changes show in amber, and hovering one lists the edits that
caused it, each with how much it contributed. The trace recalculates the
workbook to find them, so it also notices a file saved with stale
results. See [Tracing downstream changes](tracing.md).

## Next steps

- [Recipes](recipes.md): complete solutions to common jobs.
- [How diffing works](diffing.md): keys, moved rows, renames, options, and when results surprise you.
- [The DiffGrid](grid.md): layouts, display modes, annotations, events, keyboard.
- [Theming](theming.md): colors, fonts and dark mode through CSS variables.
- [Accessibility](accessibility.md): what screen-reader and keyboard users get.
- [Troubleshooting](troubleshooting.md): symptoms, causes and fixes.
- [API reference](api.md): every export.
