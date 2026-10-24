# Recipes

Complete solutions to common jobs. Every block here is typechecked
against the library by `npm run docs:check`, and the one marked as an
example is run and its output compared.

- [Compare two files a user picks, in the browser](#compare-two-files-a-user-picks)
- [Explain why a result changed](#explain-why-a-result-changed)
- [A diff in a pull-request comment](#a-diff-in-a-pull-request-comment)
- [A command-line diff of two workbooks](#a-command-line-diff-of-two-workbooks)
- [Show only what changed](#show-only-what-changed)
- [Match rows on two columns, ignore rounding](#match-rows-on-two-columns-ignore-rounding)
- [A list of changes that jumps to each one](#a-list-of-changes-that-jumps-to-each-one)
- [Warn when a file was saved without recalculating](#warn-when-a-file-was-saved-without-recalculating)
- [Let people download the diff](#let-people-download-the-diff)
- [Review comments on rows](#review-comments-on-rows)
- [Match your design system](#match-your-design-system)

## Compare two files a user picks

Parse, diff and trace in a worker, so even a large workbook never freezes
the page; then show the first changed sheet with its trace.

```ts
import { DiffGrid } from 'spreadiff';
import { createDiffWorker } from 'spreadiff/worker';

const differ = createDiffWorker(new Worker(new URL('spreadiff/worker/worker.js', import.meta.url), { type: 'module' }));
const grid = new DiffGrid({ layout: 'split', cellDisplay: 'delta' });

async function compare(oldFile: File, newFile: File, container: HTMLElement) {
  const { result, trace } = await differ.diffFiles(oldFile, newFile, {}, { trace: true });
  const sheet = Math.max(0, result.sheets.findIndex((s) => s.status !== 'equal'));
  grid.setOptions({ trace: trace && { trace, workbook: result, sheet } });
  grid.render({ diff: result.sheets[sheet].diff, container });
}

const inputs = document.querySelectorAll<HTMLInputElement>('input[type=file]');
const run = () => {
  const [a, b] = [inputs[0].files?.[0], inputs[1].files?.[0]];
  if (a && b) void compare(a, b, container);
};
inputs.forEach((i) => i.addEventListener('change', run));
```

Nothing leaves the browser: the files are read and processed in the
worker. For bundler setup, see [Workers](workers.md#setup).

## Explain why a result changed

A trace knows each downstream cell's causes, and how much each one
contributed, measured by undoing it and recalculating.

<!-- example: explain -->
```ts
import { columnLabel, diffWorkbooks, formatCell, traceWorkbook, type CellRef, type Workbook, type WorkbookDiff } from 'spreadiff';

// Lines: price × qty = total. Summary: the grand total.
const book = (applePrice: number, pearQty: number): Workbook => ({
  name: 'model.xlsx',
  sheets: [
    {
      name: 'Lines',
      columns: ['item', 'price', 'qty', 'total'],
      rows: [['apple', applePrice, 10, applePrice * 10], ['pear', 3, pearQty, 3 * pearQty]],
      formulas: [[null, null, null, '=B2*C2'], [null, null, null, '=B3*C3']],
    },
    { name: 'Summary', columns: ['line', 'value'], rows: [['Total', applePrice * 10 + 3 * pearQty]], formulas: [[null, '=SUM(Lines!D2:D3)']] },
  ],
});

const address = (d: WorkbookDiff, r: CellRef) => {
  const t = d.sheets[r.sheet].diff;
  const o = t.newTable.origin ?? { row: 1, col: 0 };
  return `${t.newTable.name}!${columnLabel(t.columns[r.col].newIndex + o.col)}${t.newRows[r.row] + o.row + 1}`;
};
const value = (d: WorkbookDiff, r: CellRef, side: 'old' | 'new') => {
  const t = d.sheets[r.sheet].diff;
  const row = side === 'old' ? t.oldRows[r.row] : t.newRows[r.row];
  const col = side === 'old' ? t.columns[r.col].oldIndex : t.columns[r.col].newIndex;
  return formatCell((side === 'old' ? t.oldTable : t.newTable).rows[row][col]);
};

const d = diffWorkbooks(book(2, 5), book(2.5, 8));
const trace = traceWorkbook(d);
const total = { sheet: 1, row: 0, col: 1 };
const t = trace.get(total)!;
console.log(`${address(d, total)} ${value(d, total, 'old')} → ${value(d, total, 'new')}`);
t.causes
  .map((c, i) => ({ c, share: t.contributions?.[i] ?? 0 }))
  .sort((a, b) => Math.abs(b.share) - Math.abs(a.share))
  .forEach(({ c, share }) => console.log(`  ${address(d, c)} ${value(d, c, 'old')} → ${value(d, c, 'new')}: ${share > 0 ? '+' : ''}${share}`));
```

<!-- output: explain -->
```text
Summary!B2 35 → 49
  Lines!C3 5 → 8: +9
  Lines!B2 2 → 2.5: +5
```

Each contribution is "the result now, minus the result without that
edit". Edits can interact (a rate times a balance), so they needn't add
up to the total change. Here they do: +9 and +5 make +14. See
[Tracing](tracing.md#causes-and-contributions).

## A diff in a pull-request comment

`spreadiff/html` has no DOM dependency, so this runs in a CI job. Plain
text in a `diff` fence colours itself on GitHub:

```ts
import { readFileSync, writeFileSync } from 'node:fs';
import { diffTables, hasChanges, parseCsv, renderText } from 'spreadiff/html';

const d = diffTables(parseCsv(readFileSync('base/prices.csv', 'utf8')), parseCsv(readFileSync('prices.csv', 'utf8')));
const s = d.stats;
writeFileSync(
  'comment.md',
  hasChanges(d)
    ? `**prices.csv**: ${s.rowsModified} edited, ${s.rowsAdded} added, ${s.rowsRemoved} removed, ${s.rowsMoved} moved\n\n` +
        '```diff\n' + renderText(d, { changedColumnsOnly: true, maxRows: 200 }) + '```\n'
    : 'No changes to **prices.csv**.\n',
);
```

Post `comment.md` with your CI's usual step (for example
`gh pr comment --body-file comment.md`). For rich HTML instead, use
`renderHTML` with `HTML_CSS` ([spreadiff/html](html.md)).

## A command-line diff of two workbooks

Every changed sheet as text, plus what moved downstream, from Node:

```ts
import { readFileSync } from 'node:fs';
import { diffWorkbooks, parseFile, traceWorkbook } from 'spreadiff/core';
import { renderText } from 'spreadiff/html';

const [a, b] = process.argv.slice(2);
const d = diffWorkbooks(parseFile(a, readFileSync(a)), parseFile(b, readFileSync(b)));
for (const s of d.sheets) {
  if (s.status === 'equal') continue;
  console.log(`\n## ${s.name} (${s.status})\n`);
  console.log(renderText(s.diff, { changedColumnsOnly: true }));
}
const { inputs, formulas, computed, stale } = traceWorkbook(d).stats;
console.log(`${inputs} input edits, ${formulas} formula edits → ${computed} downstream changes${stale ? `, ${stale} stale in the new file` : ''}`);
```

Run it with `npx tsx diff.ts before.xlsx after.xlsx`.

## Show only what changed

Fold unchanged rows, hide unchanged columns and show just the new value,
tinted:

```ts
import { DiffGrid } from 'spreadiff';

const grid = new DiffGrid({ collapseUnchanged: true, context: 0, hideUnchangedColumns: true, cellDisplay: 'new' });
grid.render({ diff, container });
```

Key columns stay visible and pinned, so every row still says which row
it is.

## Match rows on two columns, ignore rounding

```ts
import { diffTables, parseCsv } from 'spreadiff';

const d = diffTables(parseCsv(before), parseCsv(after), {
  keys: ['region', 'month'], // a composite key
  numericTolerance: 0.005, // half a cent
  ignoreCase: true,
});
console.log(d.keyColumns, d.stats);
```

Keys must exist in both tables. With `keys: 'auto'` (the default),
spreadiff picks an id-like column by itself; `d.keyColumns` says which.
See [How diffing works](diffing.md#step-2-rows).

## A list of changes that jumps to each one

```ts
import { DiffGrid, ROW_STATUS_NAMES } from 'spreadiff';

const grid = new DiffGrid();
grid.render({ diff, container });

const list = document.createElement('ol');
for (const [row, mask] of diff.changedCells) {
  const col = mask.indexOf(1);
  const li = document.createElement('li');
  const r = diff.newRows[row];
  li.textContent = `Row ${r + 2}, ${diff.columns[col].name} (${ROW_STATUS_NAMES[diff.status[row]]})`;
  li.onclick = () => grid.select(row, col); // scrolls to it, unfolding if needed
  list.append(li);
}
document.body.append(list);
```

Keyboard users get the same with `n` and `p` (or `grid.nextChange()`).

## Warn when a file was saved without recalculating

A file stores each formula's last result. If someone saved it without
recalculating (manual calculation mode, a script), those results are
stale. The trace catches it, and `applyRecalculation` shows the diff as
it would be:

```ts
import { applyRecalculation, DiffGrid, traceWorkbook } from 'spreadiff';

const trace = traceWorkbook(workbookDiff);
let shown = workbookDiff;
if (trace.stats.stale > 0) {
  console.warn(`${trace.stats.stale} formula results in the new file are out of date; showing them recalculated.`);
  shown = applyRecalculation(workbookDiff, trace);
}
new DiffGrid({ trace: { trace, workbook: shown, sheet: 0 } }).render({ diff: shown.sheets[0].diff, container });
```

## Let people download the diff

`toUnifiedCsv` writes daff's "highlighter" format, readable in any
spreadsheet app:

```ts
import { toUnifiedCsv } from 'spreadiff';

const blob = new Blob([toUnifiedCsv(diff)], { type: 'text/csv' });
const a = Object.assign(document.createElement('a'), { href: URL.createObjectURL(blob), download: 'diff.csv' });
a.click();
URL.revokeObjectURL(a.href);
```

## Review comments on rows

Annotations render your own DOM below a row and scroll with it. Give them
an `id`, so a comment keeps its element (and anything typed into it) when
the list updates:

```ts
import { DiffGrid, type CellAnnotation } from 'spreadiff';

type Comment = { id: string; author: string; body: string };
declare const comments: { id: string; row: number; column: string; author: string; body: string }[];

const grid = new DiffGrid<Comment>({
  annotations: comments.map((c): CellAnnotation<Comment> => ({ id: c.id, row: c.row, column: c.column, data: c })),
  renderAnnotation: (a) => {
    const el = document.createElement('div');
    el.className = 'review-comment'; // your page's CSS applies
    el.textContent = `${a.data.author}: ${a.data.body}`;
    return el;
  },
});
grid.render({ diff, container });
```

## Match your design system

Set the `--sd-*` variables wherever the grid sits; they inherit like any
CSS custom property:

```css
.review-panel {
  --sd-font-family: var(--brand-font);
  --sd-added: var(--brand-success);
  --sd-removed: var(--brand-danger);
  --sd-computed: var(--brand-warning);
  --sd-row-height: 28px;
}
```

Text in the status colours uses darkened `-fg` variants, so contrast
holds. Check a custom palette anyway: [Theming](theming.md#tokens).
