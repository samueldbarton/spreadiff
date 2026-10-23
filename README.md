# spreadiff

Fast, good-looking spreadsheet diffs for the web. Spreadiff does for tables what
[`@pierre/diffs`](https://diffs.com) does for code: a framework-agnostic renderer
with a React wrapper, themed with CSS variables, in split or unified layouts,
with annotations, selection and keyboard navigation. It is canvas-rendered and
virtualized, so a million-row diff scrolls at 60fps.

![Unified layout with hover popover](docs/unified.png)

This repo contains:

| Path | What |
| --- | --- |
| `packages/spreadiff` | The library: diff engine, CSV/XLSX parsers, canvas renderer, React wrapper, worker helpers |
| `apps/viewer` | A drag-and-drop viewer app built on the library (plus a Glide Data Grid mode for comparison) |
| `bench/` | Playwright benchmark: spreadiff vs. Glide Data Grid on the same diffs |

```bash
npm install
npm run dev        # viewer at http://localhost:5173
npm test           # library unit tests (vitest)
npm run test:browser  # renderer tests in Chromium; needs `npx vite --port 5199` in apps/viewer
npm run fuzz:render   # random operation sequences against the renderer (same server)
npm run build      # library (tsc) + viewer (vite)
```

## What's out there, and why build this

| Project | What it does well | Why it isn't the answer here |
| --- | --- | --- |
| [daff](https://github.com/paulfitz/daff) | The reference algorithm for table diffs: row and column alignment, a highlighter format, patches | Its HTML renderer is a plain `<table>` (no virtualization). Last push was in 2023 |
| [Glide Data Grid](https://github.com/glideapps/glide-data-grid) | Canvas grid that scales to millions of rows | A general-purpose grid: React-only, peer deps pinned to React ≤ 18, no diff semantics |
| RevoGrid, AG Grid, Handsontable | Virtualized DOM grids | DOM cells cost more per frame. Diff features would have to be bolted on |
| Univer, Fortune Sheet | Full canvas spreadsheet editors | Huge editor surface area for what is a read-only view |
| [`@pierre/diffs`](https://diffs.com) | The model for the API and the look: Shadow DOM, `light-dark()` and `color-mix()` tokens, split/stacked layouts, annotations, workers | Built for text lines, not 2-D cells |
| SheetJS | Parses every spreadsheet format | ~900 KB. The npm build (0.18.5) is stale and has known CVEs. Newer builds ship only from their CDN |

None of these diffs and renders in one piece, so spreadiff does both:

- **Diff engine** (pure TS, no DOM).
  - **Columns** are matched by name. Renames are detected by comparing contents across matched rows.
  - **Rows** are matched by key when there is a unique id-like column, otherwise by sequence.
  - **Sequence alignment** uses Myers O(ND) with a work budget, and falls back to patience anchoring (unique rows plus LIS) when the budget runs out.
  - **Modified rows**: within each gap between matches, removed and added rows are paired by cell similarity and shown as modified.
- **Canvas renderer** written from scratch, not on top of Glide. The benchmark below shows why.
- **Parsers**: an RFC 4180 CSV parser, and an XLSX reader (~4 KB on [fflate](https://github.com/101arrowz/fflate)) that reads shared strings, dates, booleans and shared formulas.

## Using the library

```ts
import { diffTables, parseCsv, DiffGrid } from 'spreadiff';

const diff = diffTables(parseCsv(before), parseCsv(after), { keys: 'auto' });
const grid = new DiffGrid({ layout: 'split' });
grid.render({ diff, container: document.getElementById('diff')! }); // give it a height
grid.nextChange();
```

React:

```tsx
import { DiffGrid } from 'spreadiff/react';

<DiffGrid
  diff={diff}
  layout="unified"
  collapseUnchanged
  annotations={[{ row: 12, column: 'price', data: comment }]}
  renderAnnotation={(a) => <Comment {...a.data} />}
  onSelectionChange={setCell}
  style={{ height: 600 }}
/>
```

Off the main thread (parse and diff in a worker; results come back in a compact
transferable encoding):

```ts
import { createDiffWorker } from 'spreadiff/worker';
const w = createDiffWorker(new Worker(new URL('spreadiff/worker/worker.js', import.meta.url), { type: 'module' }));
const { result } = await w.diffFiles(oldFile, newFile); // CSV, TSV or XLSX
```

### Diff options

| Option | Default | Notes |
| --- | --- | --- |
| `keys` | `'auto'` | `string[]`, `'auto'` (a unique id-like column that matches at least 50% of rows) or `'none'` (row order) |
| `numericTolerance` | `0` | Treat numbers within ± this as equal |
| `trimWhitespace` / `ignoreCase` | `true` / `false` | |
| `compareFormulas` | `true` | A changed formula with the same cached value still counts as a change |
| `rowSimilarity` | `0.5` | Minimum share of equal cells for a removed and an added row to pair as "modified" |
| `detectRenames` | `true` | Pair a removed and an added column when ≥ 80% of their values match |

### Renderer options

`layout` (`unified` | `split`), `collapseUnchanged`, `context`, `hideUnchangedColumns`,
`frozenColumns` (number or `'keys'`), `theme` (`system` | `light` | `dark`), `minimap`,
`tooltips`, `columnWidths`, `annotations` + `renderAnnotation`, `onCellClick`,
`onCellHover`, `onSelectionChange`.
Methods: `nextChange`, `prevChange`, `scrollToRow`, `expandAll`, `collapseAll`, `select`,
`setOptions`, `setDiff`, `refreshTheme`, `cellInfo`, `cleanUp`.

Keyboard: `n`/`j` next change, `p`/`k` previous, arrows/Tab move the selection,
⌘/Ctrl+C copies the cell, Esc clears. Drag a header edge to resize a column.

### Theming

Every color and metric is a CSS custom property. Set it on any ancestor:

```css
.my-diff {
  --sd-added: #22c55e;
  --sd-removed: #ef4444;
  --sd-modified: #8b5cf6;
  --sd-font-family: 'Inter', sans-serif;
  --sd-font-size: 12px;
  --sd-row-height: 28px;
}
```

The defaults use `light-dark()` and `color-mix()`, so derived tints follow your
base colors and the page's color scheme. Canvas can't read CSS variables, so
each token is resolved through a probe element and re-resolved when the scheme
or the theme attributes change. Annotation content stays in the light DOM and
is projected through a `<slot>`, so your page CSS styles it.

| Split | Dark | Annotations |
| --- | --- | --- |
| ![](docs/split.png) | ![](docs/dark.png) | ![](docs/annotations.png) |

## Tracing downstream changes

Most changes in a model aren't edits. They're formula results that moved
*because* something upstream was edited. spreadiff classifies every changed
cell into one of these:

| Kind | Meaning | Shown as |
| --- | --- | --- |
| `input` | A typed value was edited | green/red |
| `formula` | A formula was edited | green/red + `ƒ` |
| `computed` | Same formula, new result: a downstream change | **amber** |
| `added` / `removed` | A row or column feeding a formula appeared or disappeared | row tint |

`traceWorkbook(diff)` then finds each computed change's **root causes** and
each edit's **impacts**:
- It parses formula references: A1 refs, ranges, whole rows/columns, other sheets, defined names.
- It walks the dependency graph forward from each edit, only through cells whose value actually changed. If an intermediate result didn't move, the change stopped there.

`INDIRECT`/`OFFSET`, `NOW()`/`RAND()` and external links are flagged rather than guessed.

```ts
import { diffWorkbooks, parseXlsx, traceWorkbook } from 'spreadiff';
const diff = diffWorkbooks(parseXlsx(a), parseXlsx(b));
const trace = traceWorkbook(diff);
trace.get({ sheet: 0, row: 5, col: 1 });
// { kind: 'computed', causes: [{ sheet: 2, row: 0, col: 1 }, …], causeCount: 5, … }
grid.setOptions({ trace: { trace, workbook: diff, sheet: 0 } });
```

Formulas are compared by what they **point at**, not by their text. When a
row is inserted above, `=B5*C5` becomes `=B6*C6`, and that isn't flagged.
A `SUM(D2:D5)` that grew to `D2:D6` only because a row was added is
`computed`, with "row added" as its cause. A range that was really edited is
still a formula change.

The viewer shows the trace six ways. Pick the ones that fit:

1. **Amber downstream cells** in the grid: always on, needs no trace.
2. **Tooltip**: "changed because of 6 edits", with each cause's Δ and %.
3. **Focus on select**: outlines the cell's causes (solid), computed parents (dotted) and impacts (dashed), and dims the rest.
4. **Outputs panel**: the end results that changed, each with its drivers as chips. Answers "what happened to the bottom line?"
5. **Impact panel**: each edit, and what it moved.
6. **Why? panel**: a collapsible cause tree for the selected cell.

| Outputs + amber cells | Why? tree + focus | Tooltip |
| --- | --- | --- |
| ![](docs/options/1-stacked.png) | ![](docs/options/10-why-tree.png) | ![](docs/options/11-tooltip.png) |

Limits:
- Nothing is evaluated: values come from the files' cached results. So there's no "how much of the +12% came from price vs. units" attribution; that would need a formula engine.
- Sheet-scoped names and structured table refs (`Table1[Col]`) aren't resolved; they're marked dynamic.

## Showing a change: four cell displays

`cellDisplay` sets how a changed cell reads. Any mode can use absolute and %
deltas, and they work in both layouts.

| `stacked` (old over new) | `new` (new only) | `delta` (new + Δ, %) | `inline` (old → new) |
| --- | --- | --- | --- |
| ![](docs/options/1-stacked.png) | ![](docs/options/2-new-only.png) | ![](docs/options/3-new-delta.png) | ![](docs/options/4-inline.png) |

## spreadiff/lite: code-review style, no canvas

`spreadiff/lite` renders a diff to an HTML string (or aligned plain text). It
uses no DOM APIs, so it runs on a server, in an email or in a PR comment:

- **`split`**: side by side like a code review. One row per line, line numbers on both sides, hunk separators, and character-level highlights inside changed text.
- **`unified`**: a −/+ line pair per modified row, like `git diff`.

```ts
import { diffTables, parseCsv, renderHTML, renderText, LITE_CSS } from 'spreadiff/lite';
const html = `<style>${LITE_CSS}</style>` + renderHTML(diffTables(a, b), { layout: 'split', display: 'delta' });
console.log(renderText(diffTables(a, b)));   // for terminals / markdown
```

| Lite split | Lite unified |
| --- | --- |
| ![](docs/options/6-code-split.png) | ![](docs/options/7-code-unified.png) |

Bundle sizes (`node bench/size.mjs`, esbuild, minified):

| Entry | min | gzip |
| --- | --- | --- |
| `spreadiff/lite`: diff + CSV + HTML/text renderer | 27.2 KB | **10.5 KB** |
| `spreadiff/core`: + XLSX + tracing | 36.7 KB | 15.5 KB |
| `spreadiff`: + canvas `DiffGrid` | 83.2 KB | 30.2 KB |

Lite shows only changes plus context (capped by `maxRows`). It's for reviewing
edits; for scrolling a million-row sheet, use `DiffGrid`.

## Performance

`node bench/run.mjs` (needs `vite preview` running) loads synthetic CSV pairs
(12 columns, ~0.6% of rows changed) and scrolls each renderer for 240 frames at
about 3 rows per frame, with a horizontal wobble. Collapsing is off, so both
renderers draw every row. Headless Chromium on a 4-core cloud VM:

Ranges are across 3–4 runs, because single runs vary a lot:

| Rows | Renderer | Frame p50 | Frame p95 | Worst frame | Frames > 25 ms | Long tasks |
| --- | --- | --- | --- | --- | --- | --- |
| 100k | **spreadiff** | 16.7 ms | 17.0–17.9 ms | 21–25 ms (one outlier at 71 ms) | 0 (one run: 5) | 0 |
| 100k | Glide | 16.6–20.2 ms | 19–27 ms | 33–205 ms | 3–30 | 0–4 |
| 1M | **spreadiff** | 16.7 ms | 17.1–17.3 ms | 23–30 ms | 0–4 | 0 |
| 1M | Glide | 16.7–21.7 ms | 20–27 ms | 131–147 ms | 4–29 | 1–2 |

Pipeline at 1M rows (worker):

| Stage | Time |
| --- | --- |
| Generate + parse both files | 5.7 s (of which ~3 s is generating the synthetic CSV) |
| Diff | 2.2 s |
| Transfer to the main thread | 2.2 s (11.2 s before the compact encoding below) |

Two changes moved those numbers:

1. **Lossless number parsing.** A CSV field becomes a JS number only when it
   prints back identically (`"007"`, `"1.50"` and `"1e5"` stay text). A
   char-code scanner checks this without calling `String(Number(x))`.
2. **Columnar transfer.** Diff results cross the worker boundary as typed
   arrays plus one joined string, instead of `CellValue[][]`. Structured-cloning
   24M cells was the single biggest cost, at 11 s.

### How these numbers could be wrong

- **Single runs are noisy.** Glide's worst frame ranged from 33 to 205 ms on identical input. Don't trust any one run, including the ones behind my earlier claims.
- **Medians are vsync-capped.** Both renderers hit 16.7 ms at p50, so the medians
  can't tell them apart. The signal is in the tail: worst frame and long tasks.
  A higher-refresh display, or a test that times `drawNow()` directly, would
  separate them further.
- **Headless, software raster.** On a real GPU both get faster; the gap may
  shrink or grow.
- **Not apples to apples.** spreadiff draws more per cell: old-over-new bands,
  gutters and a minimap. Glide gets diff colors from `themeOverride`, which is
  probably not how its authors would build a diff view.
- **Heap is measured in order**, after switching renderers in the same page, so
  garbage from the previous run pollutes the figures. Treat the heap numbers in
  `bench/results/latest.json` as rough.
- **Synthetic data.** Uniform numeric columns, few changes. Long text, many
  columns or dense changes will cost more. The diff's worst case (heavy
  reordering with no key) uses the patience fallback, which gives a valid
  alignment but not always a minimal one.

## Limitations and next steps

- **Accessibility.** Canvas cells aren't in the accessibility tree. The grid exposes `role="grid"`, keyboard navigation and DOM popovers, but no per-cell ARIA. A hidden DOM mirror of the visible window would fix this.
- **Main-thread memory.** Main-thread memory is still O(cells) once the result arrives. Keeping the tables in the worker and fetching visible cells on demand would remove the remaining transfer and halve memory.
- **Moved rows** in keyed mode are aligned correctly, but not flagged as "moved".
- **Annotation identity.** Annotations are tracked by object identity, so pass a memoized array. A new array every render re-creates every annotation element.
- **XLSX reader scope.** It reads values and formulas only. No styles, merged cells or `.xls`.
- **Minimap at high change density.** When changes are dense (e.g. ~4,000 changed rows in a 1M-row diff, displayed with collapsing), every pixel bucket has a change and the strip turns solid red. Density shading would read better.
