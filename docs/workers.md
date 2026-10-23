# Workers

Parsing and diffing run in milliseconds for typical files, and in seconds
for files with millions of cells. To keep the page responsive with big
files, run them in a Web Worker.

## Setup

`spreadiff/worker` gives you a small client; the worker itself is
`spreadiff/worker/worker.js`. Create it as a module worker and hand it to
`createDiffWorker`:

```ts
import { createDiffWorker } from 'spreadiff/worker';

const differ = createDiffWorker(
  new Worker(new URL('spreadiff/worker/worker.js', import.meta.url), { type: 'module' }),
);

const { result, timings } = await differ.diffFiles(oldFile, newFile, { keys: 'auto' });
console.log(result.sheets.length, 'sheets', timings);
```

`diffFiles` reads two `File`s (from an `<input type="file">` or a drop
event), detects CSV / TSV / XLSX by name and contents, and resolves to a
[`WorkbookDiff`](api.md#workbookdiff), ready to render.

### If your bundler doesn't resolve the worker URL

The `new URL('spreadiff/worker/worker.js', import.meta.url)` form is
verified with Vite 8. If your bundler doesn't resolve package names
inside `new URL()`, add a one-line worker file of your own and point at
that instead:

```ts
// diff.worker.ts
import 'spreadiff/worker/worker.js';
```

```ts
import { createDiffWorker } from 'spreadiff/worker';

const differ = createDiffWorker(new Worker(new URL('./diff.worker.ts', import.meta.url), { type: 'module' }));
```

Both forms are covered by a test that builds them with Vite and runs them
in Chromium. Webpack 5 supports the same `new Worker(new URL(…))` syntax,
but it isn't tested here.

## API

| Method | |
| --- | --- |
| `diffFiles(oldFile, newFile, options?, settings?)` | Read, parse and diff two `File`s |
| `diffBuffers({ name, data }, { name, data }, options?, settings?)` | Same, from `ArrayBuffer`s. The buffers are **transferred**: unusable afterwards in the calling thread |
| `terminate()` | Stop the worker; pending requests reject |

- `options` are [diff options](diffing.md#options-at-a-glance).
- `settings` is [`WorkerSettings`](api.md#workersettings):
  - `header`: `'auto'` (the default) detects whether row 1 holds column names or data. `true` or `false` forces it.
  - `trace`: also trace downstream changes in the worker (below).
- Methods are plain functions, so destructuring them works:
  `const { diffFiles } = createDiffWorker(w)`.

**Errors.** A parse or diff error rejects the promise with the worker's
message, e.g. `old.xls: legacy .xls (BIFF) files aren't supported`. If the
worker itself dies (most often out of memory on a huge file), every pending
request rejects instead of waiting forever. Create a new worker to continue.

## What crosses the boundary

Copying a diff from a worker to the page by structured clone allocates one
object per cell. At 1M rows × 12 columns that alone took 11 seconds. The
worker therefore sends a packed form:
- cell types and numbers as typed arrays, which are *transferred* (zero-copy);
- all strings joined into a single string.

The client unpacks it. The same measurement drops to about 2 seconds.

If you write your own worker, use the same pieces from `spreadiff/core`:

```ts
import { diffWorkbooks, packWorkbookDiff, parseFile, unpackWorkbookDiff } from 'spreadiff/core';

// In the worker:
const result = diffWorkbooks(parseFile('a.xlsx', oldBytes), parseFile('b.xlsx', newBytes));
const { packed, transfer } = packWorkbookDiff(result);
// postMessage({ packed }, transfer);

// On the page:
const diff = unpackWorkbookDiff(packed);
```

Lossless number parsing (`"12.5"` → `12.5`, while `"007"` stays text) also
matters here: numbers travel in the typed array, strings don't.

## Formula traces

Tracing recalculates the formulas around each change. That takes 0.2–0.4 s
for a 10,000-formula model and several seconds for 300,000 formulas
([Tracing → Cost](tracing.md#cost)), which would freeze the page. Ask the
worker to do it:

```ts
import { createDiffWorker } from 'spreadiff/worker';
import { DiffGrid } from 'spreadiff';
declare const oldFile: File, newFile: File, container: HTMLElement;

const differ = createDiffWorker(
  new Worker(new URL('spreadiff/worker/worker.js', import.meta.url), { type: 'module' }),
);
const { result, trace } = await differ.diffFiles(oldFile, newFile, {}, { trace: true });

new DiffGrid({ trace: trace && { trace, workbook: result, sheet: 0 } }).render({ diff: result.sheets[0].diff, container });
```

`trace` is `undefined` when neither file has formulas. Pass
[`TraceOptions`](api.md#traceoptions) instead of `true` to set budgets.

A trace is an object with methods, which can't be posted as is. The worker
sends it packed (`packWorkbookTrace`): every cell's trace as typed arrays,
plus the parent list of every downstream and formula-edit cell, computed
in the worker. `parents()` of any other cell returns `[]`. The client
unpacks it with `unpackWorkbookTrace`. With your own worker, do the same:

```ts
import { packWorkbookDiff, packWorkbookTrace, traceWorkbook, unpackWorkbookDiff, unpackWorkbookTrace, type WorkbookDiff } from 'spreadiff/core';
declare const result: WorkbookDiff;

// In the worker:
const traced = packWorkbookTrace(traceWorkbook(result), result);
const diff = packWorkbookDiff(result);
// postMessage({ diff: diff.packed, trace: traced.packed }, [...diff.transfer, ...traced.transfer]);

// On the page:
const page = unpackWorkbookDiff(diff.packed);
const trace = unpackWorkbookTrace(traced.packed, page);
```

Measured in Chromium on the production viewer (`npm run test:hosted`):

| Workbook | Longest main-thread stall, tracing in the worker | Tracing on the page instead |
| --- | --- | --- |
| 10,000 formulas, 83 downstream changes | 35–60 ms | 222–238 ms |
| 90,000 formulas, 3,632 downstream | 130–175 ms | |
| 40,000-formula chain, every cell downstream | 240–480 ms | |
| 300,000 formulas, 11,903 downstream | 350–680 ms | |

What's left of the stall on big files is the page's share of the work:
decoding the message, building the grid and the first render. It happens
once, while the files load.

Everything runs in the browser. Nothing calls a server, and LibreOffice
was only used to produce the expected answers the tests check against.
