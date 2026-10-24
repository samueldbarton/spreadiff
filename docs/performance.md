# Performance

What was measured, how, and what the numbers don't tell you. Every figure
on this page is a range across repeated runs. Single runs varied by up to
7× in the tail, so a single run would mislead.

## Summary

| Question | Answer (1M rows × 12 columns, ~0.8% of rows changed) |
| --- | --- |
| Does scrolling keep up? | Yes. Frame p95 was 17.1–21.5 ms (60 Hz is 16.7 ms) |
| Are there hitches? | Sometimes. 1 run in 3 had one frame of ~220 ms. Cause not yet confirmed, see [below](#the-1m-row-outlier) |
| How long does the diff take? | 1.1–1.3 s keyed in the browser's worker; 2.9–5.0 s in Node without a key |
| How long until it's on screen? | 7.4–8.4 s in the benchmark: generate + parse 4.8–4.9 s, diff 1.1–1.3 s, transfer 1.5–2.2 s |
| How long does a recalculated trace take? | 0.24–0.36 s for a 10,000-formula model; 4.2–5.9 s for 300,000 formulas, 3–5× the references-only trace. See [Tracing → Cost](tracing.md#cost) |
| What does it cost to ship? | 19.9 KB gzip for `spreadiff/html`, 101.1 KB for everything; less for what you import (below) |

The rest of the page shows where each number comes from.

## Scrolling

### Method

`npm run bench` (see [Running it](#running-it)) opens the viewer in
headless Chromium 141 and loads a generated pair of CSVs:

- 12 columns: an integer `id` and 11 random two-decimal numbers.
- Per row: 0.2% deleted, 0.4% have one cell edited, 0.2% get a new row
  inserted after them. At 1M rows that is ~2,000 removed, ~2,000 added and
  ~4,000 modified rows.

It then turns *Collapse unchanged* off, so every row is drawn, and scrolls
the grid 90 px (about 3 rows) per animation frame for 240 frames. Every 30
frames the horizontal scroll position also changes. It records the time
between frames, long tasks (>50 ms, from `PerformanceObserver`), and the JS
heap. Then it switches the same page to [Glide Data
Grid](https://github.com/glideapps/glide-data-grid) and repeats.

Machine: a 4-core Intel Xeon at 2.1 GHz cloud VM, 15 GB RAM, headless
Chromium with no GPU, `deviceScaleFactor: 2`, a 1440×900
viewport.

### Results

Three runs at each size:

| Rows | Renderer | Frame p95 | Worst frame | Frames > 25 ms (of 240) | Long tasks |
| --- | --- | --- | --- | --- | --- |
| 100k | **spreadiff** | 16.9–17.5 ms | 20–65 ms | 0–3 | 0 |
| 100k | Glide | 19.1–21.8 ms | 29–44 ms | 2–5 | 0 |
| 1M | **spreadiff** | 17.1–21.5 ms | 31–221 ms | 2–6 | 0–1 |
| 1M | Glide | 20.1–23.1 ms | 34–139 ms | 2–6 | 0–1 |

The median frame was 16.6–16.8 ms for both renderers in every run, so it
is left out: it only shows that both hit vsync most of the time.

Reading it:

- **p95 is the most consistent signal, and it isn't unanimous.**
  spreadiff's p95 was lower than Glide's in 5 of 6 paired runs, by
  1.6–5.6 ms. In the sixth (1M rows) it was 1.1 ms higher: 21.5 vs.
  20.4 ms.
- **Worst frames don't separate them.** Both renderers have single frames
  above 100 ms at 1M rows, in different runs. An earlier series of three
  runs had spreadiff's 1M worst frame at 34, 228 and 233 ms. An outlier
  in roughly half the runs is not noise to wave away.
- **At 100k, spreadiff's slow frames came early** (frames 8 and 18 of 240,
  59–65 ms). That pattern fits warm-up (JIT compilation, the glyph cache
  filling) better than steady-state cost, but it hasn't been isolated.

### The 1M-row outlier

The ~220 ms frame appeared at different points in the scroll (frame 172 in
the latest series), never at the start, with the JS heap at ~590 MB. The
working hypothesis is a major garbage collection: the diff holds about 24M
cell values on the main thread, and a full mark of that heap could take
this long. That fits, but it is unconfirmed. Two things would test it:

1. Run with `--js-flags=--trace-gc` and check whether a mark-compact lines
   up with the slow frame.
2. Rerun after the planned change that keeps tables in the worker and
   fetches visible cells on demand. If the outlier disappears with the
   heap, that is strong evidence.

Until then, expect an occasional hitch at a million rows.

## The pipeline

What happens between two files and a rendered grid, at 1M rows, in the
viewer's worker (three runs):

| Stage | Time | Notes |
| --- | --- | --- |
| Generate + parse both CSVs | 4.8–4.9 s | Includes generating the synthetic CSV text. Parsing alone wasn't timed separately |
| Diff | 1.07–1.33 s | Keyed by `id` (`keys: 'auto'`) |
| Transfer to the main thread | 1.45–2.17 s | Packed encoding, then unpacked into `CellValue[][]` |

At 100k rows: diff 82–148 ms, transfer 145–191 ms.

### Keyed vs. unkeyed

Rows are matched either by key or, without one, by sequence alignment (see
[How diffing works](diffing.md#step-2-rows)). The browser benchmark only
exercises the keyed path, because its data has an `id` column. Both paths
in Node 22 on the same machine, same data, three runs each:

| Rows | `keys: 'auto'` (by `id`) | `keys: 'none'` (by sequence) |
| --- | --- | --- |
| 100k | 114–157 ms | 220–357 ms |
| 1M | 1.2–3.0 s | 2.9–5.0 s |

Both report identical stats on this data. Run-to-run variance at 1M is
large (a 2.5× spread between the fastest and slowest keyed run), probably
from garbage collection, though that wasn't checked.

### What changed the numbers

| Change | Before | After |
| --- | --- | --- |
| Packed worker transfer: typed arrays plus one joined string, instead of structured-cloning `CellValue[][]` | 11.2 s transfer at 1M | 1.5–2.2 s |
| Keyed alignment: keys stay numbers, a `Map` only for repeated keys | 3.3–3.6 s diff at 1M | 1.1–1.3 s |
| Unkeyed alignment: rows hashed to numbers instead of joined into one string each | 10–14.5 s at 1M, 0.85–0.93 s at 100k | 2.9–5.0 s, 0.22–0.36 s |
| Number parsing: a char-code scanner decides whether a field round-trips as a number, instead of `String(Number(x)) === x` | not separately measured | |

The unkeyed change has a lesson in it. The first rerun after the fix showed
no improvement at all, which looked like the hypothesis (string building
and its garbage) was wrong. The benchmark had run a stale bundle. Rebuilt,
it showed the 3–5× improvement above. When a result contradicts a profile,
check the measurement before the theory.

## Bundle size

`npm run size` bundles each entry point with esbuild (minified, ES2022,
tree-shaken):

| Entry | min | gzip | Contents |
| --- | --- | --- | --- |
| `spreadiff/html` | 51.1 KB | **19.9 KB** | Diff engine, CSV parser, number formats, HTML/text renderer |
| `spreadiff/core` | 189.9 KB | 74.0 KB | + every file reader, formula engine, tracing, review and merge |
| `spreadiff` | 275.9 KB | 101.1 KB | + canvas `DiffGrid` and `SpreadsheetDiff` |

The XLSX reader uses [fflate](https://github.com/101arrowz/fflate), which is
counted in `core`. Most of `core` is now the formula engine used for
recalculating traces and computing missing values: before it, `core` was
about 16 KB gzip. `spreadiff/html` doesn't import it.

Those are whole entries. Bundlers drop what you don't import, so the
cost depends on what you call (gzip, measured the same way):

| You import | gzip |
| --- | --- |
| `diffTables`, `parseCsv` | 8.0 KB |
| `parseXlsx` (with the engine, for missing values) | 29.5 KB |
| `parseFile` (every reader) | 42.8 KB |
| `DiffGrid`, `diffTables` | 35.7 KB |
| `SpreadsheetDiff` (everything it shows) | 97.6 KB |

`parseFile` costs about 13 KB more than `parseXlsx` because it can read
any format; import `parseXlsx` if .xlsx is all you need. `spreadiff/react` adds a thin wrapper on top of
`spreadiff`. React itself isn't counted.

### The accessibility mirror

The grid keeps a hidden ARIA copy of the visible cells
([Accessibility](accessibility.md)). To check that it costs nothing
while scrolling, the scroll benchmark was run with it on and off, in the
same session, on the same machine. That session was slower overall: Glide
Data Grid was about 1.6× slower than in the table above.

| Rows | Frame p95, mirror off | Frame p95, mirror on |
| --- | --- | --- |
| 100k | 22.3–23.5 ms (3 runs) | 21.5–23.4 ms (4 runs) |
| 1M | 22.3–28.7 ms (3 runs) | 25.2–26.6 ms (4 runs) |

The ranges overlap. The first single pair of runs at 1M (22.3 against
26.3 ms) suggested a cost; more runs didn't bear it out. The mirror
compares only scroll position and selection each frame, and rebuilds
once scrolling settles.

## How these numbers could be wrong

- **Headless, no GPU.** Real GPUs rasterise canvases faster. The
  absolute numbers would drop; whether the gap between renderers grows or
  shrinks is unknown.
- **Vsync caps what's visible.** At 60 Hz anything under 16.7 ms looks the
  same. A 120 Hz display, or timing `drawNow()` directly instead of frame
  intervals, would show differences that are hidden here.
- **The comparison isn't symmetric.** spreadiff draws more per cell
  (old-over-new bands, gutters, a minimap). Glide gets its diff colours
  from a per-cell `themeOverride`, which is probably not how its authors
  would build a diff view. The comparison says that a purpose-built
  renderer keeps up. It doesn't say Glide is slow.
- **Heap numbers are order-dependent.** Both renderers run in the same page,
  so garbage from the first one is counted against the second. The heap
  figures in `bench/results/latest.json` are rough.
- **The data is easy.** Uniform numeric columns, short values, sparse
  changes, and an id column. Long text, wide sheets, dense changes or
  heavy reordering without a key will all cost more. With no key and very
  different tables, the alignment falls back to patience anchoring. That is
  near-linear, but not always minimal.
- **Three runs is a small sample.** The ranges are min–max of three, not
  confidence intervals. They show the spread, not its distribution.

## Running it

```bash
npm run build
(cd apps/viewer && npx vite preview --port 4173) &
npm run bench                      # 100k and 1M; writes bench/results/latest.json
SIZES=100000 npm run bench         # one size
npm run size                       # bundle sizes
```

At 1M rows the page's JS heap reaches ~600 MB, plus the worker's.
