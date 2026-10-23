# Design notes

Why spreadiff is built the way it is, what each choice costs, and what
would prove it wrong. Read this before a large change.

## Prior art

Surveyed in 2026, before writing any code:

| Project | What it does well | Why it isn't the answer here |
| --- | --- | --- |
| [daff](https://github.com/paulfitz/daff) | The reference algorithm for table diffs: row and column alignment, a highlighter format, patches | Its HTML output is a plain `<table>` with no virtualization. Last pushed in 2023 |
| [Glide Data Grid](https://github.com/glideapps/glide-data-grid) | Canvas grid that scales to millions of rows | A general-purpose grid: React-only, with peer dependencies pinned to React ≤ 18, and no diff semantics |
| RevoGrid, AG Grid, Handsontable | Virtualized DOM grids | DOM cells cost more per frame, and diff features would be bolted on |
| Univer, Fortune Sheet | Full canvas spreadsheet editors | A large editor surface for what is a read-only view |
| [`@pierre/diffs`](https://diffs.com) | The model for the API and the look: Shadow DOM, `light-dark()` and `color-mix()` tokens, split and stacked layouts, annotations, workers | Built for lines of text, not a 2-D grid |
| SheetJS | Reads nearly every spreadsheet format | About 900 KB. The npm release (0.18.5) is old and has published CVEs; newer builds ship only from the vendor's CDN |

Nothing combined a table-aware diff with a renderer that stays smooth at a
million rows, so spreadiff does both. daff's ideas (match columns, then
rows, then pair leftovers by similarity) shaped the engine; the code is
new.

## The engine

### One aligned-row model for everything

A diff is a list of display rows. Each row holds an old row index and a
new row index (`-1` when absent) and a status, in typed arrays: 9 bytes
per row, so ~9 MB for a million rows on top of the tables. Modified rows
also get a byte per column marking which cells changed.

Every view reads the same structure: the unified and split grid layouts,
the lite HTML renderer, the text renderer and the CSV export. A view never
re-diffs.

*Cost:* the model can't express a row that moved. A reordered keyed row is
aligned to its counterpart, so its cells compare correctly, but it isn't
flagged as moved.

### Keys first, sequence second

There are two ways to match rows, and they differ like matching people
by passport number or by their place in a queue. A key survives any
reordering. A queue position only works if most people stay in line.

- With a unique, id-like column, rows match by key: exact, and immune to
  sorting.
- Without one, rows align like lines of text, by longest common
  subsequence. Leftover removed and added rows in each gap are then paired
  by cell similarity, so an edited row reads as *modified* and not as
  *removed + added*.

`keys: 'auto'` chooses. It picks the leftmost column whose name looks
like an id (`id`, `sku`, `Order ID`, `customerId`, …) and whose values are
unique and non-empty on both sides, and match at least half the rows.

*Why not pick any unique column?* A `name` column is often unique. Keying
on it would turn every rename into a deletion plus an insertion, which is
worse than sequence alignment. Requiring an id-like name trades missed keys
for fewer wrong keys.

*Would be proven wrong by:* real files where the auto key choice is
routinely wrong. `diff.keyColumns` reports the choice, so it can be
checked.

### Myers, with a fallback

Sequence alignment uses Myers' O((N+M)·D) algorithm over rows interned to
integers. That's fast when the edit distance D is small, which is the
normal case. It has a work budget. When the budget runs out (sorted versus
unsorted, a wholesale rewrite), it falls back to patience anchoring: rows
unique on both sides, aligned by longest increasing subsequence, with
Myers between anchors. That stays near-linear but isn't always minimal.

### Formulas are compared by what they point at

Inserting a row above `=B5*C5` turns it into `=B6*C6`. The text changed;
the meaning didn't. Formula references are mapped through the row and
column alignment before comparing, so this isn't flagged. A range that
grew only because a row was inserted inside it is a computed change, not
a formula edit.

*Cost:* a tokenizer and reference parser, plus rules for range edges
(an endpoint on a deleted row snaps inward).

### Cached values first, a formula engine to check them

Values shown are the cached results stored in the file. spreadiff
doesn't replace them with its own. It does carry a formula engine
(about 150 functions) and uses it in two places:

- **Values the file doesn't have.** Files written by tools that don't
  calculate (openpyxl, most exporters) store formulas without results.
  `parseXlsx` computes those, and only those.
- **Tracing.** Static references say which inputs *could* have moved a
  cell. Recalculating says which ones *did*: each edit is put back to
  its old value, one at a time, and the cells it reaches are evaluated.
  The same pass catches cached values the file never refreshed.

Before trusting itself, the engine evaluates a sample of each formula
shape against the old file's cached values. A shape it can't reproduce
falls back to static references for that part of the model.

*Gain:* precise causes and contributions ("+8% from price, +4% from
units"), stale values flagged rather than shown as if current, and files
without cached values still diffable.

*Cost:* `core` grew from 16 to 36 KB gzip; `lite` doesn't include the
engine. Recalculating takes 3–5× as long as static tracing
([Tracing → Cost](tracing.md#cost)). The engine is one spreadsheet's
opinion: where it disagrees with Excel (hidden rows in `SUBTOTAL`,
locale-dependent text) the calibration should catch it, but only for
the cells it samples. See [Tracing → Failure cases](tracing.md#failure-cases).

### Numbers parse only when they round-trip

A CSV field becomes a number only when it prints back identically:
`"12.5"` → `12.5`, but `"007"`, `"1.50"` and `"1e5"` stay text. Otherwise
the diff would erase a difference the user can see (`1.50` vs. `1.5`) or
claim one that isn't there. A character scanner decides this without
calling `String(Number(x))` per field.

### A small XLSX reader instead of SheetJS

The reader handles values, shared strings, dates, booleans and shared
formulas, on top of [fflate](https://github.com/101arrowz/fflate). The
whole `core` entry, tracing included, is 37.4 KB gzip.

*Cost:* no styles, merged cells, comments or legacy `.xls`. A file that
needs those needs another parser; any parser that produces a `Table` works.

## The renderer

### Canvas, written from scratch

DOM grids pay per visible cell, per frame. A canvas grid repaints only the
visible window from flat arrays. Building on Glide Data Grid was the
obvious alternative; it was rejected because:

- Diff rendering needs control over every pixel: old-over-new bands,
  split panes with synchronised scrolling, gutters, a minimap.
- Glide is React-only, and the library has to work without React.
- In the benchmark, a purpose-built renderer had a lower p95 frame time in
  5 of 6 paired runs ([Performance](performance.md#results)).

*Cost:* everything a DOM gives for free has to be rebuilt: text layout
and truncation, hit-testing, selection, copy, focus. Accessibility is the
largest gap (see below).

### Theme tokens resolved through a probe element

Canvas can't read CSS custom properties directly. Every `--sd-*` token is
applied to a hidden probe element inside the grid's Shadow DOM, read back
through `getComputedStyle`, and cached. That lets the defaults use
`light-dark()` and `color-mix()`: change one base colour and its tints
follow. The cache refreshes when the colour scheme or theme attributes
change, which takes about a millisecond.

### Shadow DOM, with slots for your content

The grid's own styles can't leak out, and page styles can't break it.
Annotations and popovers you supply stay in the light DOM and are
projected through `<slot>`s, so your CSS still styles them.

### Downstream changes are amber

A computed change (same formula, new result) is a different kind of fact
from an edit. Showing it in the edit colours makes a one-cell change
look like fifty edits. Amber separates "someone typed this" from "this
moved because of something else". The colour choice hasn't been tested
with users.

## The lite renderer

`spreadiff/lite` renders to an HTML string with no DOM APIs, so it runs on
a server, in CI or in an email. It shows changes plus context, like a code
review, and caps the output with `maxRows`.

*Why a separate entry and not a mode of `DiffGrid`?* Different jobs. A
pull-request comment needs ten rows of static HTML at 11.8 KB; a reviewer
scrolling a million rows needs the canvas. Sharing the diff model keeps
them consistent.

## Known gaps

- **Accessibility.** Canvas cells aren't in the accessibility tree. The grid has `role="grid"`, keyboard navigation and DOM popovers, but no per-cell ARIA. A hidden DOM mirror of the visible window would fix it.
- **Main-thread memory.** Tables are copied to the main thread, so memory is O(cells) there. Keeping them in the worker and fetching visible cells on demand would remove the transfer cost and roughly halve peak memory. It may also remove the occasional long frame at 1M rows ([Performance](performance.md#the-1m-row-outlier)).
- **Moved rows** aren't flagged in keyed mode.
- **Annotation identity.** Annotations are tracked by object identity. Pass a stable (memoized) array; a new array on every render re-creates every annotation element.
- **Dense minimap.** With thousands of changes in a long sheet, every minimap pixel holds a change and the strip turns solid. Density shading would read better.
- **Sheet-scoped names and structured references** (`Table1[Col]`) aren't resolved by tracing; they're marked dynamic.
