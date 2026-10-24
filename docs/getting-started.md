# Getting started

This page takes you from two files to an interactive diff in a few minutes.
For how the comparison works, see [How diffing works](diffing.md).

## Install

```bash
npm install spreadiff
```

For [spreadiff-lite](spreadiff-lite.md), the report and CLI with no UI:

```bash
npm install spreadiff-lite      # or run it once: npx spreadiff-lite old.xlsx new.xlsx
```

To try unreleased changes, build from the repository and install the
package directories:

```bash
git clone https://github.com/samueldbarton/spreadiff
cd spreadiff && npm install && npm run build -w spreadiff && npm run build -w spreadiff-lite
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
| `spreadiff/html` | Diff + CSV + an HTML/text renderer. No canvas or DOM APIs, so it runs on servers too | 19.9 KB |
| `spreadiff/core` | The diff engine, every file reader, formula tracing, review and merge. No rendering | 74.2 KB |
| `spreadiff` | Everything in core plus the canvas `DiffGrid` and `SpreadsheetDiff` | 101.3 KB |
| `spreadiff/react` | The `SpreadsheetDiff` and `DiffGrid` React components | + React |
| `spreadiff/worker` | Parse, diff and trace in a Web Worker | small client + worker |
| `spreadiff-lite` | A separate package with no UI: a Markdown or JSON report of what changed and what it moved downstream, and a CLI | Node / server |

Rule of thumb:
- To review a handful of edits, especially outside a browser, use `spreadiff/html`.
- To scroll through large sheets interactively, use `DiffGrid`.
- To tell a reader (or an agent) what changed and what it moved, in
  Markdown, use `spreadiff-lite`.

## The short path: two files in, a review out

`SpreadsheetDiff` takes two spreadsheets (`File`s from an `<input>`, URLs,
bytes) in any [supported format](formats.md) and does the rest. It
compares every sheet, traces which changes are edits and which are
formula results that moved, and shows sheet tabs, the grid, and a panel of
changed outputs, each with a waterfall of what drove it:

```ts
import { SpreadsheetDiff } from 'spreadiff';
declare const container: HTMLElement; // give it a height

const view = new SpreadsheetDiff();
await view.render({ oldFile: new URL('/v1.xlsx', location.href), newFile: new URL('/v2.xlsx', location.href), container });
```

With React:

```tsx
import { SpreadsheetDiff } from 'spreadiff/react';

export const Review = ({ before, after }: { before: File; after: File }) => (
  <SpreadsheetDiff oldFile={before} newFile={after} style={{ height: 600 }} />
);
```

URLs are fetched for you; there's no need to `fetch` them first. The
component draws inside a shadow root: style it through the
[theme variables](theming.md), not selectors.

Add `review: true` to accept and reject changes and download the result
(see [Reviewing changes](review.md)). For large files, pass `worker` so the
work happens off the main thread ([below](#3-big-files-parse-off-the-main-thread)).

To lay out the page yourself, `diffFiles` gives you the comparison, and
each of its sheets goes straight to a `DiffGrid` with its trace attached:

```ts
import { diffFiles, DiffGrid } from 'spreadiff';
declare const before: File, after: File, container: HTMLElement;

const cmp = await diffFiles(before, after);
new DiffGrid({ layout: 'split' }).render({ sheet: cmp.sheets[cmp.firstChanged], container });
console.log(cmp.report().edits.length, 'edits');
```

The rest of this page builds the same thing from its parts.

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

For files, `parseFile(name, bytes)` picks the parser from the contents
(XLSX, XLSB, XLS, ODS, CSV or TSV), and `diffWorkbooks` compares every sheet:

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
import { diffFiles } from 'spreadiff';
declare const oldFile: File, newFile: File;

const cmp = await diffFiles(oldFile, newFile, {
  worker: () => new Worker(new URL('spreadiff/worker/worker.js', import.meta.url), { type: 'module' }),
});
```

The worker is created on first use and kept. `SpreadsheetDiff` takes the
same `worker` option. Tracing (next section) runs in the worker too. See [Workers](workers.md) for bundler
notes and what crosses the boundary.

## 4. Spreadsheets with formulas

For workbooks, spreadiff also tells you **why** a cell changed. It
separates edited inputs from formula results that moved as a consequence,
and traces each result back to its causes. `diffFiles` and
`compareWorkbooks` trace whenever there are formulas, and each sheet view
carries its trace:

```ts
import { compareWorkbooks, DiffGrid, type Workbook } from 'spreadiff';
declare const before: Workbook, after: Workbook, container: HTMLElement;

const cmp = compareWorkbooks(before, after);
new DiffGrid().render({ sheet: cmp.sheets[0], container });
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
