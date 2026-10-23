# spreadiff

Spreadsheet diffs for the web. Diff two CSV or XLSX files, then render the
result as a virtualized canvas grid or as static HTML. For spreadsheets
with formulas, spreadiff tells edited inputs apart from results that moved
as a consequence, and traces each one back to its causes.

![A spreadsheet diff with downstream changes in amber](https://raw.githubusercontent.com/samueldbarton/spreadiff/main/docs/images/viewer.png)

```ts
import { DiffGrid, diffTables, parseCsv } from 'spreadiff';

const diff = diffTables(parseCsv(before), parseCsv(after));
new DiffGrid({ layout: 'split', cellDisplay: 'delta' }).render({ diff, container });
```

| Import | Contents | min + gzip |
| --- | --- | --- |
| `spreadiff/lite` | Diff engine, CSV parser, HTML and text renderers. No DOM APIs | 11.8 KB |
| `spreadiff/core` | + XLSX reader, formula engine and tracing | 36.3 KB |
| `spreadiff` | + canvas `DiffGrid` | 51.5 KB |
| `spreadiff/react` | React component for `DiffGrid` (React 18+) | + React |
| `spreadiff/worker` | Parse and diff in a Web Worker | |

**Documentation:**
[Getting started](https://github.com/samueldbarton/spreadiff/blob/main/docs/getting-started.md) ·
[How diffing works](https://github.com/samueldbarton/spreadiff/blob/main/docs/diffing.md) ·
[Tracing](https://github.com/samueldbarton/spreadiff/blob/main/docs/tracing.md) ·
[The DiffGrid](https://github.com/samueldbarton/spreadiff/blob/main/docs/grid.md) ·
[API reference](https://github.com/samueldbarton/spreadiff/blob/main/docs/api.md) ·
[Performance](https://github.com/samueldbarton/spreadiff/blob/main/docs/performance.md)
