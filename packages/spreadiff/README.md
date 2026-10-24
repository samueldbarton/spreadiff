# spreadiff

**See what changed between two spreadsheets, and why.**

Diff two CSV or XLSX files, then show the result as a virtualized canvas
grid or as static HTML. Columns match by name and rows by key or by
sequence. An edited row reads as modified and a reordered row as moved.
For workbooks with formulas, spreadiff tells edited inputs apart from
results that moved as a consequence. It traces each result back to its
causes, with each cause's contribution, by recalculating with its own
formula engine, checked against LibreOffice.

![A spreadsheet diff with downstream changes in amber](https://raw.githubusercontent.com/samueldbarton/spreadiff/main/docs/images/viewer.png)

```ts
import { DiffGrid, diffTables, parseCsv } from 'spreadiff';

const diff = diffTables(parseCsv(before), parseCsv(after));
new DiffGrid({ layout: 'split', cellDisplay: 'delta' }).render({ diff, container });
```

| Import | Contents | min + gzip |
| --- | --- | --- |
| `spreadiff/html` | Diff engine, CSV parser, HTML and text renderers. No DOM APIs | 16.0 KB |
| `spreadiff/core` | + XLSX reader, formula engine and tracing | 43.4 KB |
| `spreadiff` | + canvas `DiffGrid` | 63.9 KB |
| `spreadiff/react` | React component for `DiffGrid` (React 18+) | + React |
| `spreadiff/worker` | Parse, diff and trace in a Web Worker | |

- **Runs anywhere.** `spreadiff/html` and `spreadiff/core` run in Node and edge runtimes. The
  grid needs Chrome / Edge 111, Firefox 113 or Safari 16.2. Nothing is sent
  anywhere, and it works under a strict CSP.
- **Accessible.** Screen readers get an ARIA grid of the visible cells,
  and the colours meet WCAG AA in light and dark.
- **Measured.** Causes and contributions match LibreOffice's on every
  scored cell of the test corpus (3,400 contributions). The method and
  failure cases are in [Tracing](https://github.com/samueldbarton/spreadiff/blob/main/docs/tracing.md#how-it-was-tested).

**Documentation:**
- [Getting started](https://github.com/samueldbarton/spreadiff/blob/main/docs/getting-started.md)
- [Recipes](https://github.com/samueldbarton/spreadiff/blob/main/docs/recipes.md)
- [How diffing works](https://github.com/samueldbarton/spreadiff/blob/main/docs/diffing.md)
- [Tracing](https://github.com/samueldbarton/spreadiff/blob/main/docs/tracing.md)
- [The DiffGrid](https://github.com/samueldbarton/spreadiff/blob/main/docs/grid.md)
- [Accessibility](https://github.com/samueldbarton/spreadiff/blob/main/docs/accessibility.md)
- [Troubleshooting](https://github.com/samueldbarton/spreadiff/blob/main/docs/troubleshooting.md)
- [API reference](https://github.com/samueldbarton/spreadiff/blob/main/docs/api.md)

MIT licensed.
