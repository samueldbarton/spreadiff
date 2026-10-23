# Tracing downstream changes

In a spreadsheet model, most changed cells weren't edited. They're formula
results that moved *because* something upstream was edited. Change one
price, and every revenue, margin and total that depends on it changes too.
A plain diff shows all of them in the same color, and the one edit that
caused them disappears in the noise.

spreadiff separates the two, and traces each result back to its causes.

![A financial model: edited inputs in green and red, downstream results in amber](images/display-stacked.png)

## Four kinds of change

Every changed cell gets a [`ChangeKind`](api.md#changekind):

| Kind | What happened | Grid color |
| --- | --- | --- |
| `input` | A typed-in value was edited | green / red |
| `formula` | A formula was edited (compared by meaning; see [Formulas](diffing.md#formulas)) | green / red, with `ƒ` |
| `computed` | **Same formula, new result**: a downstream change | amber |
| `added` / `removed` | A cell of an added or removed row/column that feeds a formula | row tint |

Telling `input` from `computed` needs only the two files, not a full
trace. So the grid colors downstream changes amber by default
(`highlightComputed`), even with no trace attached.

## Root causes and impacts

`traceWorkbook(diff)` builds the formula dependency graph of both
workbooks. It then walks forward from every root change (inputs, formula
edits, added and removed cells), recording:

- on each `computed` cell, the **causes**: which root changes it depends on;
- on each root change, the **impacts**: which computed cells it moved.

<!-- example: trace -->
```ts
import { diffWorkbooks, traceWorkbook, type Table, type Workbook } from 'spreadiff';

// price × qty = total, and a grand total: =SUM(D2:D3)
const book = (applePrice: number): Workbook => {
  const lines: Table = {
    name: 'Lines',
    columns: ['item', 'price', 'qty', 'total'],
    rows: [['apple', applePrice, 10, applePrice * 10], ['pear', 3, 5, 15]],
    formulas: [[null, null, null, '=B2*C2'], [null, null, null, '=B3*C3']],
  };
  const summary: Table = {
    name: 'Summary',
    columns: ['line', 'value'],
    rows: [['Total', applePrice * 10 + 15]],
    formulas: [[null, '=SUM(Lines!D2:D3)']],
  };
  return { name: 'model.xlsx', sheets: [lines, summary] };
};

const diff = diffWorkbooks(book(2), book(2.5));
const trace = traceWorkbook(diff);

const price = trace.get({ sheet: 0, row: 0, col: 1 })!; // Lines!B2
const total = trace.get({ sheet: 1, row: 0, col: 1 })!; // Summary!B2
console.log(price.kind, price.impactCount);
console.log(total.kind, JSON.stringify(total.causes));
```

<!-- output: trace -->
```text
input 2
computed [{"sheet":0,"row":0,"col":1}]
```

The price edit moved two cells (the apple's total and the grand total),
and the grand total traces back to exactly one edit: the price.

The walk **only passes through cells whose value changed**. If an
intermediate result came out the same, say price went up and quantity down
by the same factor, the change stopped there, and nothing below it lists
those edits as causes.

For a cause *tree* (one level at a time), `trace.parents(ref)` returns the
changed cells a formula reads directly.

## What the trace understands

| In a formula | Handled |
| --- | --- |
| `B2`, `$B$2`, `B2:D9` | ✓ |
| `C:C`, `3:5` (whole columns / rows) | ✓, clipped to the table |
| `Lines!D2`, `'Q&A 2026'!B3` | ✓, other sheets of the same workbook |
| `TaxRate` (workbook-scoped defined names) | ✓, read from the XLSX |
| Ranges that grew because rows were inserted | ✓, treated as the same formula (cause: "row added") |
| Rows removed from a range | ✓, the removed cells become causes |
| `NOW()`, `TODAY()`, `RAND()` | flagged `unexplained: 'volatile'`, and treated as their own cause |
| `INDIRECT`, `OFFSET`, unknown names, external `[1]Book!A1` links | flagged `unexplained: 'dynamic'` |
| Sheet-scoped names, structured refs (`Table1[Col]`) | not resolved; flagged `dynamic` |

A computed change with no cause the trace can see gets
`unexplained: 'unknown'`. Typically it reads a cell outside the parsed
range, or a value that changed without its formula being stored.

## Showing it

Attach the trace to a grid:

```ts
import { DiffGrid, traceWorkbook } from 'spreadiff';

const trace = traceWorkbook(workbookDiff);
const grid = new DiffGrid({
  trace: { trace, workbook: workbookDiff, sheet: 0 }, // `sheet`: which sheet this grid shows
  cellDisplay: 'delta', // new value with Δ and % underneath
});
grid.render({ diff: workbookDiff.sheets[0].diff, container });
```

Hovering a downstream cell then lists its causes with their changes, and
selecting it outlines them and dims everything else:

| Tooltip | Focus on select |
| --- | --- |
| ![Tooltip listing the six edits that changed Net income](images/trace-tooltip.png) | ![EMEA revenue selected: its causes outlined, the rest dimmed](images/trace-focus.png) |

In the focus view:
- **solid** outlines are root causes;
- **dotted** outlines are computed cells the formula reads directly;
- **dashed amber** outlines are the cells an edit flowed into.

Causes on other sheets appear in the tooltip and in
`CellInfo.trace`; the outlines only cover the visible sheet.

### Panels in the viewer app

The viewer in [`apps/viewer`](../apps/viewer/src/TracePanel.tsx) builds
three panels on top of the trace API. They're example code to copy or adapt,
not library components:

| Outputs | Impact | Why? |
| --- | --- | --- |
| End results that changed (cells nothing else reads), each with the edits that drove it | Each edit, ranked by how much it moved | The cause tree of the selected cell |
| ![](images/viewer-outputs.png) | ![](images/viewer-impact.png) | ![](images/viewer-why.png) |

"Outputs" answers the question reviewers usually start with: *what
happened to the bottom line, and why?* It's computed as the changed
computed cells that no other changed cell reads.

## Cost

There's no formula evaluation: values come from the cached results stored
in the files. Building the graph is linear in the number of formulas, and
the walk visits only changed cells. A 20,000-cell running total
(`C[i] = C[i-1] + B[i]`) where every cell changes traces in about 0.6 s
(measured in Node on a 4-core cloud VM). A `budget` option caps the work, after which `trace.truncated` is set.

## Limits

- **Causes, not contributions.** The trace can say *Net income +12% was
  driven by price, EMEA units and the tax rate*, but not how much each
  contributed. That needs re-evaluating the model with one change at a time,
  i.e. a formula engine.
- **Only stored formulas.** A formula replaced by a pasted value reads as a
  `formula` change on that cell (the formula is gone), and the trace can't
  see what the old formula depended on beyond the old file's references.
- **One workbook.** References to other workbooks are flagged, not followed.
