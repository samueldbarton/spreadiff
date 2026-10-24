# Troubleshooting

Symptoms first, then the likely cause and the fix. If yours isn't here,
[How diffing works → When results surprise you](diffing.md#when-results-surprise-you)
covers more about matching.

## The grid

**Nothing shows.** The grid fills its container and scrolls inside it,
so a container with no height shows nothing. Give it one: a fixed
`height`, `flex: 1` in a flex column, or `position: absolute; inset: 0`.

**Colours are wrong, or everything is black.**
- **An old browser.** The grid needs CSS `color-mix()`: Chrome / Edge
  111, Firefox 113, Safari 16.2. Older browsers can't resolve its colours.
  See [Browser support](browser-support.md).
- **An invalid custom value.** A `--sd-*` value the browser can't parse
  (a typo, an unsupported colour function) makes that token invalid, and
  everything derived from it.
- **Variables set from script.** If you changed them from JavaScript, call
  `grid.refreshTheme()`. Changes to `class`, `style` or `data-theme` on
  `<html>` or the container are picked up by themselves.

**The scroll position resets.** `setDiff` with a new diff object resets
scroll, selection and folding. In React, memoize the diff (`useMemo`):
a new object on every render looks like a new diff.

**No popover.** Popovers appear on changed cells and on rows that moved,
after a short hover. Check that `tooltips` isn't `false`.

**Text is cut off with "…".** Columns size to a sample of their
content. Drag a header's right edge, or set `columnWidths: { name: px }`.

**Screen readers don't read the cells.** The accessibility mirror is on
by default. Check that `accessibility` isn't `false`, and that the grid
has focus (click it or Tab to it). See [Accessibility](accessibility.md).

## Matching

**Columns are named `A`, `B`, `C`, or the header shows up as a row.** For
XLSX, `header: 'auto'` decides per sheet whether row 1 holds names or
data. A row with numbers, dates or formulas in it, or one that formulas
read, is data. Two exceptions keep the names:
- years next to labels;
- a row that is mostly labels.

Force it with `header: true` or `false` (in `parseXlsx`, `parseFile`, or
the worker's `{ header }`). See [How diffing works → Step 1](diffing.md#step-1-columns).

**Every row is marked changed.** The rows were probably reordered and no
key column was found, so they were matched by order. Check
`diff.keyColumns`; pass `keys: ['id']` if it's empty. With a key,
reordered rows show as moved.

**An edited row shows as removed + added.** More than half of its cells
changed, so by order it no longer looks like the same row. Match by key,
or lower `rowSimilarity`.

**Numbers that look the same are flagged, or tiny differences aren't.**
By default numbers are equal when they agree to 15 significant digits
(`precision`), which hides float noise such as `0.1 + 0.2`. For
tolerances in your own units (rounding to cents), use
`numericTolerance`.

## Formulas and tracing

**A downstream cell's causes say "references".** Its causes came from
following formula references, not from recalculating. The likely
reasons:
- the engine doesn't support a function on the path;
- it couldn't reproduce the old file's result for that formula;
- the cell uses `INDIRECT` / `OFFSET`;
- two edits were each enough on their own.

References never miss a cause but can list one that didn't matter.
`trace.stats.unmodelled` counts the formulas involved. See
[Tracing → When it can't recalculate](tracing.md#when-it-cant-recalculate).

**Changes with no cause (`unexplained: 'unknown'`).** The two files
probably came from different spreadsheet apps (Excel and LibreOffice,
say). The same formula can give slightly different results there, in
`TEXT` formats, date rounding or `CELL`, with nothing edited. Compare
files saved by the same app.

**"Saved without recalculating".** The new file's stored results don't
match its own formulas, for example because it was saved in manual
calculation mode or written by a script. The viewer shows the
recalculated values. In code, use `applyRecalculation(diff, trace)`
([recipe](recipes.md#warn-when-a-file-was-saved-without-recalculating)).

**A file shows "computed" values.** It had formulas without stored
results (openpyxl and similar tools save them that way), so spreadiff
computed them when parsing. `Workbook.computedValues` has the counts.
Turn it off with `computeMissing: false`.

**The trace is slow on a big model.** Recalculating 300,000 formulas
takes seconds. Run it in the worker (`{ trace: true }`,
[Workers](workers.md#formula-traces)). For references only, which is 3–5×
faster, pass `{ recalculate: false }`.

## Files

**"legacy .xls (BIFF) files aren't supported".** Save the file as
`.xlsx` or `.csv`. `.xlsb` and `.ods` aren't read either.

**A CSV column of numbers is text.** Numbers are only converted when
they round-trip exactly: `007`, `1.50` and `1e5` stay text, so a
formatting change is never hidden. Compare them with `numericTolerance`,
which reads both sides as numbers.

**Odd characters in a CSV.** Plain-text CSV is read as UTF-8. Excel's
"Unicode Text" export (UTF-16) is detected by its byte-order mark when
you use `parseFile`. A different delimiter is detected; force one with
`delimiter`.

## Workers and bundlers

**The worker never replies, or fails to load.**
- **The bundler.** Some bundlers don't resolve package names inside
  `new URL('spreadiff/worker/worker.js', import.meta.url)`. Use a one-line
  worker file of your own instead ([Workers →
  setup](workers.md#if-your-bundler-doesnt-resolve-the-worker-url)).
- **Module workers.** It must be one (`{ type: 'module' }`). Firefox
  supports module workers from 114.
- **A CSP** needs `worker-src 'self'` (or `script-src`) to allow it.

**"Diff worker crashed".** Usually out of memory on a huge file. The tab
holds both tables. Pending requests reject rather than hang. Create a new
worker to continue.
