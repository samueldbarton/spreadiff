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
| SheetJS | Reads nearly every spreadsheet format | About 900 KB. The npm release (0.18.5) is old and has published CVEs; newer builds ship only from the vendor's CDN |

Nothing combined a table-aware diff with a renderer that stays smooth at a
million rows, so spreadiff does both. daff's ideas (match columns, then
rows, then pair leftovers by similarity) shaped the engine; the code is
new.

### Downstream impact (surveyed October 2026)

Before building [spreadiff-lite](spreadiff-lite.md), the question was
whether anything already reports what an edit *moved downstream*. Several
tools get part of the way:

| Tool | Downstream | Recalculates | Per-output causes and contributions | Labels |
| --- | --- | --- | --- | --- |
| [ExcelCompare](https://github.com/na-ka-na/ExcelCompare), [xlsxDiff](https://github.com/rafal-dot/xlsxDiff), [sheets-diff](https://github.com/forskscope/sheets-diff-rs), [git-xl](https://github.com/xltrail/git-xl) | No: cell-by-cell differences | No | No | No |
| Microsoft Spreadsheet Compare | Colours "calculated values" apart from "entered values" | Uses saved values | No | No |
| [sheetdelta](https://github.com/sqmyou/sheetdelta) | Lists the addresses a change reaches ("affects 3 cells downstream") | No ("it does not evaluate formulas") | No | No |
| [xlsx-review](https://github.com/Keremozdemirra/xlsx-review) | `explain_cell` lists a cell's dependents | No ("never calculates anything") | No | No |
| [agent-spreadsheet](https://github.com/PSU3D0/agent-spreadsheet) (`asp`) | `verify diff` separates "direct edits vs recalculation fallout" | Yes, a separate step (Formualizer) | Not documented | Not documented |
| spreadiff / spreadiff-lite | Every downstream result | Yes, checked against LibreOffice | Yes: each edit undone in turn | Section, row and column headings |

*How this could be wrong:* the survey covered public repositories and
vendor pages. Commercial add-ins (PerfectXL, Operis OAK, Arixcel) didn't
document their compare features well enough to judge, and
agent-spreadsheet's diff output isn't documented in detail. The claim is
narrow: the combination of recalculated per-output attribution and
label-aware Markdown wasn't found, not that no tool traces dependents.

## The engine

### One aligned-row model for everything

A diff is a list of display rows. Each row holds an old row index and a
new row index (`-1` when absent) and a status, in typed arrays: 9 bytes
per row, so ~9 MB for a million rows on top of the tables. Modified rows
also get a byte per column marking which cells changed.

Every view reads the same structure: the unified and split grid layouts,
the HTML renderer, the text renderer and the CSV export. A view never
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

*Cost:* `core` grew from 16 to 36 KB gzip; `spreadiff/html` doesn't include the
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
whole `core` entry, tracing and the other readers included, is 74.0 KB gzip;
`parseXlsx` alone, with the engine it uses to compute missing values, is
about 30 KB.

The .xlsb, .xls and .ods readers share its sheet model (`core/sheet.ts`),
and .xls and .xlsb share one formula-token decoder.

*Cost:* no styles beyond number formats, and no comments. A file that
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

## The HTML renderer

`spreadiff/html` renders to an HTML string with no DOM APIs, so it runs on
a server, in CI or in an email. It shows changes plus context, like a code
review, and caps the output with `maxRows`.

*Why a separate entry and not a mode of `DiffGrid`?* Different jobs. A
pull-request comment needs ten rows of static HTML at 19.9 KB; a reviewer
scrolling a million rows needs the canvas. Sharing the diff model keeps
them consistent.

## Known gaps

Still open, roughly by how much they'd matter:

- **Main-thread memory.** Tables are copied to the main thread, so memory there is O(cells). Keeping them in the worker and fetching visible cells on demand would:
  - remove the transfer cost;
  - roughly halve peak memory;
  - possibly remove the occasional long frame at 1M rows ([Performance](performance.md#the-1m-row-outlier)).

  It changes how every renderer gets its data, so it needs its own measurement work.
- **Formats other than .xlsx are read-only**, and .xlsb structured references aren't decoded. See [File formats](formats.md).
- **Styles.** Number formats are read and cells show as the sheet shows them, but other formatting isn't read, and a change of format alone isn't detected.
- **Vertical merges** are drawn cell by cell, since an inserted row can split them. Horizontal merges are drawn as one cell.
- **Formula-edit undo across inserted rows.** Tracing undoes an edited formula by re-evaluating its old text. When rows or columns were inserted or deleted, that text points elsewhere, so it restores the old result instead ([Tracing → Limits](tracing.md#limits)).
- **Screen readers, manually.** The grid's ARIA mirror passes automated checks but hasn't been tried with NVDA, JAWS or VoiceOver ([Accessibility](accessibility.md#not-covered-yet)).
- **Downstream changes by hue alone.** Visually, an edited input and a downstream result differ by colour (blue against amber), with the kind named in the popover and to screen readers. A pattern would serve people who can't tell those hues apart.

Closed since the first release:
- moved rows are flagged in keyed mode;
- annotations are matched by `id` or position, not object identity;
- the minimap shades by density;
- sheet-scoped names and structured references are resolved;
- merged cells are read;
- canvas cells are mirrored for screen readers, and every text colour meets WCAG AA.
