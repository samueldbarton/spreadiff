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
| `diffFiles(oldFile, newFile, options?, parse?)` | Read, parse and diff two `File`s |
| `diffBuffers({ name, data }, { name, data }, options?, parse?)` | Same, from `ArrayBuffer`s. The buffers are **transferred**: unusable afterwards in the calling thread |
| `terminate()` | Stop the worker; pending requests reject |

- `options` are [diff options](diffing.md#options-at-a-glance).
- `parse` is `{ header?: boolean }` (default: the first row is a header).
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

`traceWorkbook` returns an object with methods, which can't be posted
between threads. Run it on the page after the diff arrives. It doesn't
evaluate formulas, so it's fast: see [Tracing → Cost](tracing.md#cost).
