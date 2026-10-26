# Reviewing and merging changes

A diff says what changed. Reviewing decides what to keep: accept some
changes, reject others, and get the file back with the rejected ones
undone. Merging combines two people's changes to the same starting file.

## In the component

`SpreadsheetDiff` with `review: true` adds Accept, Reject and Undo buttons
over the grid, and a download button. Select a cell or drag over a range,
then press a button, or `a`, `r` or `u` in the grid:

```ts
import { SpreadsheetDiff } from 'spreadiff';
declare const before: File, after: File, container: HTMLElement;

const view = new SpreadsheetDiff({
  review: true,
  onDecisionsChange: (decisions) => console.log(decisions.length, 'sheets'),
});
await view.render({ oldFile: before, newFile: after, container });
```

A rejected change is dimmed and marked ↺, an accepted one ✓, and the
header counts what's decided. The download is the new .xlsx with rejected
changes undone (for an .xlsx or .xlsm new file), or the shown sheet as CSV
for other formats. `view.exportXlsx()`, `view.exportCsv()` and
`view.mergedTable()` give you the same without the button.

### Keeping decisions

To restore a review (saved, or shared with a colleague), pass `decisions`.
They're applied when the files are compared, and replaced whenever you
pass a different array. In React, keep them in state: store what
`onDecisionsChange` gives you and pass it back. Passing back the array the
component just gave you changes nothing, so this doesn't loop.

```tsx
import { SpreadsheetDiff } from 'spreadiff/react';
import type { SheetDecisions } from 'spreadiff';
import { useState } from 'react';
declare const before: File, after: File;

export function Review() {
  const [decisions, setDecisions] = useState<SheetDecisions[]>();
  return <SpreadsheetDiff oldFile={before} newFile={after} review decisions={decisions} onDecisionsChange={setDecisions} style={{ height: 600 }} />;
}
```

The component still updates its own decisions when someone presses
Accept or Reject; `decisions` doesn't have to be passed back for that.

## What can be decided

- **Edits**: a value or formula someone changed, and a value from outside
  data (a refresh) are decisions to make.
- **Rows and columns** that were added or removed, and renamed columns.
- **Not downstream results.** A formula result that moved because its
  inputs did isn't a decision: it follows them. To undo it, reject the edits
  that moved it (the tooltip and the outputs panel name them). Selecting a
  range that includes results decides only its edits. `isDecidable` tells
  them apart.

Anything not decided is accepted, so rejecting is the only step you need.

## With the functions

Decisions are plain data, `SheetDecisions`, one per sheet: `Map`s from a
cell key (`row * diff.columns.length + col`), an aligned row or a column
index to `'accept'` or `'reject'`. `decideRange` fills them from a range,
such as the grid's `getRange()`; `applyDecisions` gives back the table.

<!-- example: review -->
```ts
import { applyDecisions, decideRange, diffTables, noDecisions, parseCsv } from 'spreadiff';

const before = parseCsv('id,item,price\n1,Lamp,10\n2,Chair,20');
const after = parseCsv('id,item,price\n1,Lamp,12\n2,Chair,20\n3,Rug,50');
const diff = diffTables(before, after);

// Keep the old price of the Lamp (aligned row 0, column 2), and drop the Rug (aligned row 2).
let decisions = decideRange(diff, noDecisions(), [0], [2], 'reject');
decisions = decideRange(diff, decisions, [2], [0, 1, 2], 'reject');
console.log(JSON.stringify(applyDecisions(diff, decisions).rows));
```

<!-- output: review -->
```text
[[1,"Lamp",10],[2,"Chair",20]]
```

`tableToCsv` writes the result as CSV.

A rejected removed row comes back where it was; a rejected removed column
comes back after the column it followed.

## Writing the .xlsx back

`applyDecisionsToXlsx` writes rejected cell changes into the new .xlsx in
place, so styles, charts and everything the reader doesn't read survive:

```ts
import { applyDecisionsToXlsx, diffWorkbooks, parseXlsx, type SheetDecisions } from 'spreadiff';
declare const oldBytes: Uint8Array, newBytes: Uint8Array, decisions: SheetDecisions[];

const diff = diffWorkbooks(parseXlsx(oldBytes), parseXlsx(newBytes));
const { file, applied, recalculated, skipped } = applyDecisionsToXlsx(newBytes, diff, decisions);
```

- An old formula goes back with its references moved to the new file's
  rows (`=B4*2` becomes `=B5*2` if a row went in above). One that points at
  a cell the new file doesn't have is skipped.
- Formulas that depend on a reverted cell get fresh cached values
  (`recalculated` counts them), computed by spreadiff's engine, and the
  file is marked to recalculate fully when opened. Excel does; LibreOffice
  and some viewers show the cached values, which is why they're refreshed.
- Not written, and listed in `skipped` with a reason: row and column
  decisions (inserting or deleting rows would mean rewriting every
  reference in the workbook), cells in array formulas, and table headers.
  For those, apply the decisions to the tables with `applyDecisions` and
  export CSV.

The other formats (.xlsb, .xls, .ods) are read-only: review them, then
export each sheet as CSV.

## Three-way merge

When two people change copies of the same file, `mergeTables` combines
them against the original. A change made on one side is taken; the same
change made on both is taken once; different changes to the same cell are
conflicts. Rows are matched by key, as `diffTables` matches them, so edits
to different rows or columns never conflict, wherever the rows moved.

<!-- example: merge -->
```ts
import { mergeTables, parseCsv } from 'spreadiff';

const base = parseCsv('id,price,qty\n1,10,5\n2,20,1');
const ours = parseCsv('id,price,qty\n1,12,5\n2,20,1');
const theirs = parseCsv('id,price,qty\n1,10,6\n2,25,1\n3,30,2');
const { table, conflicts, stats } = mergeTables(base, ours, theirs);
console.log(JSON.stringify(table.rows), conflicts.length, JSON.stringify(stats));
```

<!-- output: merge -->
```text
[[1,12,6],[2,25,1],[3,30,2]] 0 {"ours":1,"theirs":3,"both":0}
```

A conflict takes `prefer` (ours by default) and is listed with the base,
our and their values. To show the result, give `result.diff` to a
`DiffGrid` with `conflictAnnotations(result)` as its annotations: each
conflict appears under its row. Kinds of conflict:

| `kind` | When |
| --- | --- |
| `cell` | Both changed a cell, to different values |
| `delete-edit` | One side deleted a row the other edited |
| `both-added` | Both added a row with the same key, with different values |

For merging CSV files in git, see [the merge driver](git.md#merging-csv-files).

## How it can be wrong

- **Undoing an edit doesn't undo what someone did because of it.** If a
  reviewer rejects a price change, a discount someone set to compensate
  stays. Reject related edits together.
- **The recomputed values are spreadiff's.** Its engine reproduces the
  cached result of 99.8% of the formulas in the test corpus (46,162 of
  46,249; see [Tracing](tracing.md)). A formula it can't compute keeps its
  cached value until the file is recalculated.
- **The merge is cell-wise.** Two edits that are each fine but wrong
  together (one side raises prices, the other raises the discount assuming
  the old prices) merge without a conflict.
