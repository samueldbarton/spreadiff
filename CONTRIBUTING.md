# Contributing

## Setup

```bash
git clone https://github.com/samueldbarton/spreadiff
cd spreadiff
npm install
npm run dev          # the viewer at http://localhost:5173, with the library from source
```

Node 20.19+ or 22.12+ (Vite 8's minimum). The browser-based checks use Playwright with Chromium.
If Chromium isn't at `/opt/pw-browsers/chromium`, set `CHROMIUM` to its
path, or run `npx playwright install chromium`.

## Checks

| Command | What it checks | Needs | When to run it |
| --- | --- | --- | --- |
| `npm test` | Unit tests for the engine, parsers, tracing, labels, the HTML renderer, the worker client, and spreadiff-lite's report and Markdown (vitest) | | Every change |
| `npm run typecheck` | All workspaces | | Every change |
| `npm run docs:check` | Doc code blocks typecheck, examples print their shown output, every export is documented, links resolve, bundle sizes are current | | Every change to the API, the docs or the bundle |
| `npm run test:browser` | The canvas renderer in Chromium: regressions; merges and moved rows; annotation identity and minimap density; the ARIA mirror, with axe-core on the grid and the HTML renderer; canvas colour contrast (WCAG AA, light and dark); the adopted stylesheet and the `light-dark()` fallback | Dev server on port 5199 (below); restart it if results look stale | Renderer, theme or HTML-renderer changes |
| `npm run fuzz:render` | A seeded sequence of 3,000 random operations against `DiffGrid`, checking invariants after each | Dev server on port 5199 | Renderer changes |
| `npm run test:bundle` | Packs the library, installs it into a fresh Vite app, builds it and runs it in Chromium: `exports`, both worker setups, `sideEffects` | Network access the first time (it runs `npm install`) | Changes to `package.json`, entry points or the worker |
| `npm run test:cli` | Builds both packages and runs the `spreadiff-lite` CLI on two real `.xlsx` files: output, `--json`, exit codes. Then packs both packages, installs the tarballs into an empty project, and runs the installed bin and API | Network access the first time (it runs `npm install`) | Changes to spreadiff-lite, `package.json` or entry points |
| `npm run docs:site` | Builds the documentation site into `apps/docs/dist` and fails on a broken link or `#anchor` between its pages. `npm run docs:dev` serves it on port 4321 and rebuilds on change | | Changes to the docs or the site |
| `npm run build` | Library and spreadiff-lite (tsc), viewer (Vite) | | Before a PR |
| `npm run size` | Bundle size per entry point | | Changes that add code to an entry |
| `npm run bench` | Scroll smoothness and pipeline timings vs. Glide Data Grid | `vite preview` on port 4173 (below) | Performance work |
| `npm run test:hosted` | Tracing in a browser, hosted: the packed library in a fresh Vite app and the viewer's production build, both served as static files under a strict CSP with outside requests blocked. Every changed cell of 19 real `.xlsx` fixtures is checked against LibreOffice's answers (causes, contributions, stale and missing values), axe-core runs over the viewer, and main-thread stalls are measured. If `bench:trace` has built the large models, those are timed too | Network access the first time (it runs `npm install`) | Changes to tracing, the engine, the XLSX reader, the worker or the viewer's trace UI |
| `npm run bench:trace` | Tracing and the formula engine against LibreOffice: generated models, real Excel files (Apache POI test corpus) and 40k–300k-formula models. Prints the tables in [docs/tracing.md](docs/tracing.md#how-it-was-tested). About 30 minutes from scratch; `-- --skip-build` reuses the cases | LibreOffice (`soffice`), Python with openpyxl, network access the first time | Changes to tracing, the engine, the XLSX reader or the diff |
| `npm run bench:engine` | The formula engine alone against LibreOffice: about 350 curated edge cases plus random expressions (`-- 3000 11` for 3,000 with seed 11). Every difference is printed; [bench/engine-fuzz/README.md](bench/engine-fuzz/README.md) lists the ones where LibreOffice, not Excel, is the odd one out. About a minute | LibreOffice (`soffice`), Python with openpyxl | Changes to `core/engine` |

Starting the servers:

```bash
(cd apps/viewer && npx vite --port 5199) &          # test:browser, fuzz:render
npm run build && (cd apps/viewer && npx vite preview --port 4173) &   # bench
```

`fuzz:render` takes an operation count and a seed (default 7):
`node bench/fuzz-render.mjs 500 42`. Each failure prints as JSON with the operation index, the error and the
eight operations before it. The sequence
is deterministic, so rerunning with the same count and seed replays it.

## Documentation

The docs live in `docs/`, with the entry point at
[docs/README.md](docs/README.md). Every public export needs JSDoc and an
entry in [docs/api.md](docs/api.md); `docs:check` fails otherwise.

**Code blocks.** Every ` ```ts ` or ` ```tsx ` block is typechecked
against the library source. Common free variables (`before`, `after`,
`container`, `diff`, `oldFile`, …) are declared in
[docs/snippets.d.ts](docs/snippets.d.ts). To opt a block out, put
`<!-- no-check -->` on the line before it. A ` ```ts signature ` block is
also skipped.

**Examples with output.** Tag a block with `<!-- example: some-id -->` and
put the expected output in a ` ```text ` block tagged
`<!-- output: some-id -->`. `docs:check` runs the example and compares.

**Screenshots.** Images in `docs/images/` are generated, not hand-made:

```bash
npm run docs:screenshots
```

This starts the viewer on port 5299, renders each scene from
`apps/viewer/shots.html` (the library alone) and the app itself, and
overwrites the images. To add one, add a scene in
`scripts/screenshots.mjs`. Check the result by eye: the script fails on
page errors, but not on an ugly picture.

**Numbers.** A performance claim needs a measurement behind it, stated as
a range across runs, with the method nearby. Bundle sizes are checked
automatically; timings aren't, so update [docs/performance.md](docs/performance.md)
when you change something on the hot path.

**Style.** Short sentences, concrete examples, and the reason behind a
default. When something can go wrong, say so on the same page.

## Pull requests

- Run `npm test`, `npm run typecheck` and `npm run docs:check`, plus
  whichever browser checks the table above lists for your change.
- Add a test for every bug fix. Fail it first, then fix.
- Add a line to [CHANGELOG.md](CHANGELOG.md) under *Unreleased* for
  anything a user would notice.
