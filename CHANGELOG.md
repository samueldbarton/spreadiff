# Changelog

spreadiff isn't published yet; versions below are the package version in
the repository.

## Unreleased

### Added

- **Formula tracing.** `traceWorkbook` classifies every changed cell as an
  `input` edit, a `formula` edit or a `computed` change, and links each
  computed change to its root causes and each edit to its impacts.
- **Downstream changes in amber** in `DiffGrid`, with cause lists in the
  tooltip, and a focus mode that outlines a selected cell's causes and
  impacts.
- **Cell displays:** `cellDisplay: 'stacked' | 'new' | 'delta' | 'inline'`,
  with absolute and percent change for numbers.
- **`spreadiff/lite`**: an HTML and text renderer with no DOM APIs, in a
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

- The worker entry was tree-shaken out of production builds, because the
  package declared `sideEffects: false`. Workers built that way did
  nothing.

## 0.1.0

Initial version: the diff engine (column matching with rename detection,
keyed and sequence row alignment, similarity pairing), CSV and XLSX
parsers, the canvas `DiffGrid` with unified and split layouts, the React
wrapper, worker helpers with packed transfer, and the viewer app. Includes
the fixes from an adversarial review and from fuzzing.
