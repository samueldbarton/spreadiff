---
name: spreadiff
description: Compare two versions of a spreadsheet (.xlsx, .xlsm, .xlsb, .xls, .ods, .csv, .tsv) and report what was edited, which formula results moved because of it, and by how much. Use after editing a workbook to check the edit did what was intended, to explain why a number changed between two versions, to review someone else's spreadsheet changes, to add a spreadsheet diff UI to a web app, or to set up spreadsheet diffs in git or CI.
---

# spreadiff

Two tools, pick by what the output is for:

- **A report** (an agent's reply, a CI log, a PR comment): the
  `spreadiff-lite` CLI or its `compare` / `toMarkdown` API. No UI.
- **A UI in a web app**: the `spreadiff` library. `SpreadsheetDiff` for
  the whole review in one component, `DiffGrid` for just the grid.

## Check your own edit to a workbook

Keep a copy of the file before editing it, then:

```bash
spreadiff-lite before.xlsx after.xlsx --outputs-only
```

Exit code: 0 no changes, 1 changes, 2 error. Read the report before
replying:

- **Edits** must be exactly the cells you meant to change. An extra edit
  is a mistake (often a shifted range or a value typed over a formula).
- **Final outputs** are the totals and results nothing else reads. Check
  the ones you meant to move moved, in the direction and by roughly the
  size you expect, and that nothing else did.
- **Driven by** names the edits behind each result, with their
  contributions. A result driven by an edit you didn't expect is a
  dependency you missed.
- **Notes** flags stale files (saved without recalculating) and formulas
  spreadiff computed because the file had no value for them, as openpyxl
  writes. Say so in your reply when it appears.

Then include the Markdown in your reply, trimmed with `--max-rows 20` if
it is long.

Useful options: `--json` (a stable structure to check in code),
`--sheet NAME`, `--keys id,region` (match rows on these columns),
`--min-change 0.1%` (ignore immaterial moves), `--all` (no row caps).

## Explain why a number changed

```bash
spreadiff-lite v1.xlsx v2.xlsx --json > report.json
```

Find the cell in `downstream` by `id` (`Summary!C12`). Its `causes` list
the edits that moved it, largest first, each with `contribution`: the
result now minus the result with that edit undone, in the cell's units.
Contributions add up to the change (`delta`) unless edits interact (a
price and a volume both moved); the difference is that interaction. They
come from recalculating, so they're measured, not estimated. A cell with
`method: 'references'` couldn't be recalculated: its causes are the edits
its formulas read, which may include some that didn't move it, and its
contributions are null.

## In code (Node)

```ts
import { readFileSync } from 'node:fs';
import { compare, toMarkdown } from 'spreadiff-lite';

const report = compare(
  { name: 'model.xlsx', data: readFileSync('backup/model.xlsx') },
  { name: 'model.xlsx', data: readFileSync('model.xlsx') },
);
const outputs = report.downstream.filter((c) => c.output);
console.log(toMarkdown(report, { maxRows: 30 }));
```

## A diff UI in a web app

```ts
import { SpreadsheetDiff } from 'spreadiff';

// oldFile/newFile: File, Blob, URL, bytes, or { name, data }. Give the container a height.
await new SpreadsheetDiff({ review: true }).render({ oldFile, newFile, container });
```

React: `import { SpreadsheetDiff } from 'spreadiff/react'` and
`<SpreadsheetDiff oldFile={a} newFile={b} style={{ height: 600 }} />`.
For large files pass `worker: () => new Worker(new URL('spreadiff/worker/worker.js', import.meta.url), { type: 'module' })`.
To build your own layout: `const cmp = await diffFiles(a, b)`, then
`new DiffGrid().render({ sheet: cmp.sheets[cmp.firstChanged], container })`.
Don't pair a grid with a trace by sheet index by hand; the sheet view
carries the right one.

## Git and CI

- Readable `git diff`: `*.xlsx diff=spreadsheet` in `.gitattributes`, and
  `git config diff.spreadsheet.textconv "spreadiff-lite --textconv"`.
- The full report in `git diff`: `git config diff.spreadsheet.command "spreadiff-lite --git"`.
- Merging CSVs by row key: `git config merge.spreadsheet.driver "spreadiff-lite merge %O %A %B -o %A --name %P"`.
- Pull-request comments: the GitHub Action `samueldbarton/spreadiff@main`
  (needs `fetch-depth: 0` on checkout and `pull-requests: write`).

Details: `docs/git.md`.

## Limits to tell the user about

- Recalculation uses spreadiff's own engine, checked against
  LibreOffice, not Excel. A formula it can't compute is traced by its
  references (`method: 'references'`), never guessed.
- Data-feed formulas (Bloomberg `BDP`, `RTD`, links to other workbooks)
  can't be recalculated; their changed values are reported as inputs.
- Only .xlsx can be written back (`applyDecisionsToXlsx`); other formats
  export CSV. .xlsb structured references aren't decoded.
- A change of formatting alone (colour, font) isn't a change.
