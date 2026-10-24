# Changelog

spreadiff isn't published yet; versions below are the package version in
the repository.

## Unreleased

### Added

- **spreadiff-lite**, a new package with no UI. `compare` two workbooks
  into a report of every edit, with its before and after values and what
  the sheet calls the cell. Each downstream result comes with the edits
  that moved it and their contributions, plus shape changes and notes.
  `toMarkdown` renders the report, `summarize` does both, and
  `spreadiff-lite OLD NEW` is the CLI (`--json`, `--outputs-only`, exit
  codes like `diff`). Final outputs are listed apart from intermediate
  results. See `docs/spreadiff-lite.md`.
- **Outside data.** A changed cell whose formula reads data from outside
  the workbook is now an `input`, not a recalculated result, and the
  results it feeds are traced to it. That covers Bloomberg `BDP`/`BDH`,
  LSEG, FactSet, Capital IQ, Morningstar, `RTD`, `WEBSERVICE`,
  `STOCKHISTORY`, any `_xll.` add-in function and links to other
  workbooks. New: `CellTrace.external`, `TraceStats.external`,
  `TraceOptions.liveFunctions`, `externalSource`, `LIVE_FUNCTIONS`.
- **Cell labels.** `tableLabeler` and `formatLabel` name a cell from the
  layout: section headings, row label and column headings
  (`India › Revenue · 2025`). The XLSX reader now reads indent and bold
  on text cells (`Table.labelStyles`) to nest sections.
- **A documentation site** (`apps/docs`, `npm run docs:site`) built from
  the same Markdown:
  - search, light and dark themes, contents and copy buttons;
  - live demos with Preview and Code tabs;
  - "Typechecked" and "Verified output" marks on code blocks;
  - `llms.txt` and the Markdown of each page.

  A GitHub Pages workflow publishes it.

- **Moved rows.** With rows matched by key, rows whose position changed
  are flagged: the fewest rows that explain the new order
  (`TableDiff.moved`, `DiffStats.rowsMoved`, `CellInfo.moved`). They're
  shown with `↕` in the grid and the HTML renderer, `~` in text, and `:` in
  `toUnifiedCsv`, and they count for `hasChanges` (now also exported from
  `spreadiff/html`).
- **Accessibility.** The grid mirrors its visible cells as a hidden ARIA
  grid:
  - row and column headers;
  - each change in words ("70.63, was 65.4, up 5.23 (8.0%), edited");
  - the selection as `aria-activedescendant`;
  - a summary as the accessible name.

  New options `accessibility` and `ariaLabel`. Every text colour meets
  WCAG AA in light and dark, and new `--sd-modified-fg` /
  `--sd-computed-fg` tokens carry text in those hues. Checked by axe-core
  on the grid, the HTML renderer and the viewer, and by a canvas contrast test. See
  `docs/accessibility.md`.
- **Merged cells** are read into `Table.merges`. The grid draws an
  unchanged horizontal merge as one cell.
- **Sheet-scoped defined names** are read into `Table.names` and resolved
  by formula references, the engine and tracing, including `Sheet!Name`
  from other sheets (`scopedNames`).
- **Annotation `id`.** Annotations are matched across updates by `id` or
  by row and column, no longer by object identity. A fresh array of equal
  annotations keeps its elements, and only changed `data` re-renders.
- **The minimap shades by density**, so thousands of changes no longer
  paint it solid.
- **CSP and older browsers.**
  - The grid's styles are an adopted stylesheet, so it runs under a strict
    Content-Security-Policy (tested).
  - The grid and `HTML_CSS` carry a fallback for browsers without CSS
    `light-dark()` (Safari before 17.5, Chrome before 123), derived from
    the same tokens and tested to give identical colours.
- **Docs:**
  - new guides: recipes, troubleshooting, accessibility and browser support;
  - a rewritten README;
  - a `LICENSE` file.
- **Formula tracing.** `traceWorkbook` classifies every changed cell as an
  `input` edit, a `formula` edit or a `computed` change, and links each
  computed change to its root causes and each edit to its impacts.
- **Recalculated tracing.** A formula engine (about 150 functions,
  structured and 3-D references, array formulas) undoes each edit in turn
  and recalculates what it reaches. Causes are the edits that actually
  moved a value, with each one's contribution
  (`CellTrace.contributions`). Formula shapes the engine can't reproduce
  from the old file's cached values fall back to references
  (`CellTrace.method`, `TraceStats.unmodelled`). Options `recalculate`
  and `evaluationBudget`.
- **Stale values.** Cells whose saved value doesn't match the file's own
  formulas are flagged (`CellTrace.stale`, a banner in the viewer);
  `applyRecalculation` shows the diff as if the file had been
  recalculated.
- **Missing values.** Formulas saved without a value (openpyxl and other
  tools that don't calculate) are computed on parsing
  (`computeMissingValues`, `ParseXlsxOptions.computeMissing`,
  `Workbook.computedValues`).
- **Tracing in the worker.** `createDiffWorker` takes `{ trace: true }`
  (`WorkerSettings`) and resolves with the trace too, so recalculation
  never blocks the page. `packWorkbookTrace` / `unpackWorkbookTrace` send
  it as typed arrays, parents included. The viewer traces in its worker.
- `npm run test:hosted`: the packed library in a fresh, statically served
  Vite app and the viewer's production build. Real `.xlsx` fixtures are
  checked against LibreOffice's answers, with outside network blocked.
- `npm run bench:trace`: checks tracing against LibreOffice on generated
  models and real Excel files from the Apache POI test corpus. Results
  are in [Tracing → How it was tested](docs/tracing.md#how-it-was-tested).
- `DiffOptions.precision`, `Workbook.tables` / `TableDefinition`,
  `structuredRange`, `FormulaContext`, `Table.header`, `Table.arrays`,
  `Workbook.date1904`.
- **Downstream changes in amber** in `DiffGrid`, with cause lists in the
  tooltip, and a focus mode that outlines a selected cell's causes and
  impacts.
- **Cell displays:** `cellDisplay: 'stacked' | 'new' | 'delta' | 'inline'`,
  with absolute and percent change for numbers.
- **`spreadiff/html`**: an HTML and text renderer with no DOM APIs, in a
  code-review style split or unified layout.
- **Formula comparison by meaning.** References are mapped through the row
  and column alignment, so formulas that only shifted with an insert
  aren't flagged.
- **Documentation:** guides, an API reference covering every export,
  performance and design notes, and generated screenshots. `npm run
  docs:check` keeps code blocks, example output, links, API coverage and
  bundle sizes in sync with the source.
- `npm run test:bundle`: installs the packed library into a fresh Vite app
  and runs both documented worker setups in Chromium.
- `npm run size`.

### Changed

- **`spreadiff/lite` is now `spreadiff/html`**, and `LITE_CSS`,
  `LiteOptions`, `LiteLayout` and `LiteDisplay` are now `HTML_CSS`,
  `HtmlOptions`, `HtmlLayout` and `HtmlDisplay`. The name `spreadiff-lite`
  now belongs to the new package.

- **Undoing a formula edit** while tracing re-evaluates the old formula on
  today's inputs, instead of restoring its old result. Contributions of
  formula edits are now exact. Against LibreOffice, 3,400 of 3,400
  contributions match, where the old way had 95 wrong. With rows or
  columns inserted, it still restores the old result.
- `hideUnchangedColumns` and `changedColumnsOnly` keep the row-label
  column (the first mostly-text one) when rows were matched by order, so
  every changed row still says which row it is.
- Text colours are darker in light mode (secondary text, row numbers,
  `+`/`−` signs, deltas) to meet WCAG AA.
- `spreadiff` is 57.1 KB gzip, up from 53.1, for the accessibility mirror,
  merges, moved rows and the colour fallback. `spreadiff/html` is 13.1 KB.
- The viewer has a page heading and landmarks, and doesn't inline assets
  as `data:` URLs.
- `parseFile`, the worker and the viewer default to `header: 'auto'`. The
  viewer passed `true`, which hid row-1 inputs in financial models even
  after the parser learned to detect them. The viewer's toggle is now a
  three-way "First row" setting.
- The viewer's Outputs panel lists end results largest relative change
  first, 100 at a time, with up to 8 drivers each. It had drawn every
  result: 47,854 DOM nodes and an 845 ms freeze on a 3,632-result model.
  Above 5,000 results it showed nothing.
- The viewer serves its own fonts instead of loading Google Fonts.
- `PackedTable.formulas` is packed rather than nested arrays.
- `spreadiff/core` is 38.9 KB gzip, up from 16.1, almost all of it the
  formula engine. `spreadiff/html` is unchanged in scope (13.1 KB).
- `parseXlsx` defaults to `header: 'auto'`: a first row with numbers,
  dates or formulas, or one that formulas read, is data rather than
  column names. Years next to labels still count as names, and a row
  of labels kept as data still names the columns.
- Numbers compare to 15 significant digits by default, so Excel's
  17-digit and LibreOffice's 15-digit writes of the same result are equal.
- The grid no longer draws a corner triangle on changed cells; the
  formula mark (ƒ) remains.
- Computed cells use a clearer amber.
- Downstream (computed) cells are quieter: a lighter amber tint and no
  left bar, so they read as secondary to edits.
  The dark-mode amber no longer turns olive. Modified rows show a dot in
  the gutter instead of `~`, and column change counts are subtler.
- The viewer's toolbar is regrouped: layout and cell display as segmented
  controls, matching and comparison settings in one Options popover, and
  the change summary beside the sheet tabs. The landing page has sample
  cards and clearer drop zones.

- Unkeyed row alignment hashes rows instead of building a string per row:
  3–5× faster (0.85–0.93 s → 0.22–0.36 s at 100k rows; 10–14.5 s → 2.9–5.0 s
  at 1M, in Node).
- Keyed alignment keeps numeric keys as numbers and only tracks repeated
  keys separately: diff time at 1M rows went from 3.3–3.6 s to 1.1–1.3 s.
- `renderText` marks removed and added lines with ASCII `-` and `+`, so
  Markdown ` ```diff ` blocks colour them.
- `createDiffWorker` returns plain functions; destructuring them works.

### Removed

- The `SPREADIFF_CSS` export, which was unused.
- The `--sd-modified-bg` and `--sd-computed-fg` theme tokens, which had no
  effect.

### Fixed

- The text renderer wrote "1 unchanged rows".
- The closed tooltip was exposed to assistive technology without a name.
- Regressions found by comparing with 0.1.0, fixed before release:
  - diffs of formula-heavy XLSX sheets were 4–7× slower (row identity
    built and hashed each formula's relative text; now an
    allocation-free hash, checked against the old form on 2 million
    formulas);
  - `header: 'auto'` named columns `A`, `B`, … on sheets with years
    across the top, or with a parameter among the labels, losing key
    matching (on the 90,000-formula test model the diff fell back to
    sequence alignment and took 4× longer).
- `WorkbookTrace.parents()` built the old workbook's whole formula graph
  on first use, about 1.7 s at 300,000 formulas, which froze the page on
  the first click in focus mode. It now reads just that cell's references.
- A tooltip too tall to fit above or below its cell was placed partly
  outside the grid, cutting off its title. It now takes the roomier side
  and stays inside.
- XLSX: structured references (`Sales[Amount]`), 3-D references
  (`Jan:Mar!B2`), shared formulas over whole columns or rows, CRLF and
  `_x000D_` escapes in text, and LibreOffice's `=#VALUE!` and `TRUE()`
  forms were misread, so traces through them were lost or formulas
  showed as edited.
- An edited row in a model sheet could split into removed + added when
  the edit changed most of its values; equal-size gaps now pair rows in
  place. Row identity includes formula shapes, so a row with an edited
  formula isn't matched as unchanged.
- The worker entry was tree-shaken out of production builds, because the
  package declared `sideEffects: false`. Workers built that way did
  nothing.

## 0.1.0

Initial version: the diff engine (column matching with rename detection,
keyed and sequence row alignment, similarity pairing), CSV and XLSX
parsers, the canvas `DiffGrid` with unified and split layouts, the React
wrapper, worker helpers with packed transfer, and the viewer app. Includes
the fixes from an adversarial review and from fuzzing.
