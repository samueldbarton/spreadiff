# spreadiff-lite

**What changed between two spreadsheets, and what it moved downstream, as
Markdown or JSON.** No UI.

spreadiff-lite is for the cases where nobody is looking at a grid:
- an agent that edits a workbook and has to say what its edits did;
- a CI job that comments on a pull request;
- a script that checks a model before it goes out.

It reads two `.xlsx`, `.csv` or `.tsv` files and reports four things:
- **each edit:** the cell, what the sheet calls it, and its value before
  and after;
- **each result that moved downstream:** with the edits that moved it and
  how much each one contributed, found by recalculating the workbook;
- **values from data feeds:** Bloomberg, Capital IQ, FactSet, LSEG, RTD,
  links to other workbooks. They are reported as inputs, never
  recalculated;
- **shape changes:** rows, columns and sheets added, removed or moved.

It uses the same engine as `spreadiff`, so the
[numbers on correctness](tracing.md#how-it-was-tested) apply to it too.

- [Quick start](#quick-start)
- [The report](#the-report)
- [How cells are named](#how-cells-are-named)
- [Outside data](#outside-data)
- [With an agent](#with-an-agent)
- [Command line](#command-line)
- [How it can be wrong](#how-it-can-be-wrong)

## Quick start

Install it with `npm install spreadiff-lite`, or run it without
installing:

```bash
npx spreadiff-lite model-v1.xlsx model-v2.xlsx            # Markdown to stdout
npx spreadiff-lite model-v1.xlsx model-v2.xlsx --json     # the report as JSON
```

It exits with 0 when the files are the same, 1 when they differ and 2 on
an error, like `diff`.

From code:

```ts
import { readFileSync } from 'node:fs';
import { compare, toMarkdown } from 'spreadiff-lite';

const report = compare(
  { name: 'model-v1.xlsx', data: readFileSync('model-v1.xlsx') },
  { name: 'model-v2.xlsx', data: readFileSync('model-v2.xlsx') },
);
console.log(toMarkdown(report)); // or JSON.stringify(report)
```

`summarize(oldFile, newFile)` does both in one call. Already have parsed
`Workbook`s, or a diff and trace from `spreadiff`? Pass the workbooks to
`compare`, or the diff and trace to `reportOf`.

## The report

A model in two versions. Growth was raised from 10% to 12%, and the
Bloomberg rupee rate refreshed from 83 to 84:

<!-- example: lite-regions -->
```ts
import { computeMissingValues, type CellValue, type Table, type Workbook } from 'spreadiff/core';
import { compare, toMarkdown } from 'spreadiff-lite';

// A sheet from rows of values and =formulas. Labels indented with spaces, as in many models.
const sheet = (name: string, rows: (CellValue | string)[][]): Table => ({
  name,
  columns: rows[0].map((_, i) => String.fromCharCode(65 + i)),
  origin: { row: 0, col: 0 },
  rows: rows.map((r) => r.map((v) => (typeof v === 'string' && v.startsWith('=') ? null : v))),
  formulas: rows.map((r) => r.map((v) => (typeof v === 'string' && v.startsWith('=') ? v : null))),
});

const model = (name: string, growth: number, inr: number): Workbook => {
  const wb: Workbook = {
    name,
    sheets: [
      sheet('Inputs', [['Growth', growth], ['USD/INR', '=BDP("USDINR Curncy","PX_LAST")']]),
      sheet('Model', [
        [null, 2024, '=B1+1'],
        ['India', null, null],
        ['  Revenue', 8300, '=B3*(1+Inputs!$B$1)'],
        ['  Costs', 4150, '=B4*(1+Inputs!$B$1)'],
        ['  Profit (USD)', '=(B3-B4)/Inputs!$B$2', '=(C3-C4)/Inputs!$B$2'],
      ]),
    ],
  };
  wb.sheets[0].rows[1][1] = inr; // the feed's last value, as the file saved it
  computeMissingValues(wb); // the other formulas' results, as Excel would have saved them
  return wb;
};

console.log(toMarkdown(compare(model('regions-v1.xlsx', 0.1, 83), model('regions-v2.xlsx', 0.12, 84))));
```

<!-- output: lite-regions -->
```text
## regions-v1.xlsx → regions-v2.xlsx

**2 edits** (1 input, 1 outside-data value) moved **4 results downstream** (2 final outputs). Every cause was found by recalculating.

### Edits

| Cell | What | Before | After | Change | Moves |
| --- | --- | --- | --- | --- | --- |
| `Inputs!B1` | Growth | 0.1 | 0.12 | +0.02 (+20.0%) | 3 |
| `Inputs!B2` | USD/INR *(Bloomberg BDP)* | 83 | 84 | +1 (+1.20%) | 2 |

### Final outputs

| Cell | What | Before | After | Change | Driven by |
| --- | --- | --- | --- | --- | --- |
| `Model!B5` | India › Profit (USD) · 2024 | 50 | 49.4047619047619 | −0.5952380952381 (−1.19%) | USD/INR −0.5952 |
| `Model!C5` | India › Profit (USD) · 2025 | 55 | 55.3333333333333 | +0.3333333333333 (+0.61%) | Growth +0.9881; USD/INR −0.6667 |

### Intermediate results

| Cell | What | Before | After | Change | Driven by |
| --- | --- | --- | --- | --- | --- |
| `Model!C3` | India › Revenue · 2025 | 9130 | 9296 | +166 (+1.82%) | Growth +166 |
| `Model!C4` | India › Costs · 2025 | 4565 | 4648 | +83 (+1.82%) | Growth +83 |

### Notes

- 1 value came from outside data (a data add-in, a web or RTD function, or another workbook). It is taken as saved, not recalculated.
```

Reading it:
- **Edits** are the cells someone changed:
  - `input`: a value was typed over;
  - `formula`: the formula was edited, shown as formulas with the result
    each gave;
  - outside data: marked with the provider and function.

  *Moves* counts the downstream results each edit moved.
- **Final outputs** are results that no other changed formula reads: the
  ends of the chains, usually what a reader cares about.
  **Intermediate results** are the steps in between. When every result is
  an output, there is one *Downstream* table.
- **Driven by** lists the edits that moved a result, largest first. Each
  one's contribution is *the result now, minus the result without that
  edit*, found by undoing the edit and recalculating. Edits can interact (a
  rate times a balance), so contributions needn't add up to the change.
  *(by references)* marks a result whose causes come from formula
  references instead, because a formula on the path couldn't be
  evaluated. Those causes are complete but may include one that didn't
  matter, and have no contributions.
- **Notes** say what to be careful of, such as values from outside data,
  stale files and unevaluated formulas.

Values are printed as the file stores them (15 significant digits) and
contributions to 4. `toMarkdown(report, { maxRows, maxCauses, digits,
intermediate: false, headingLevel })` changes that.

<!-- demo: lite -->

### As JSON

`JSON.stringify(report)` gives the same content as data:

```js
{
  old: 'regions-v1.xlsx', new: 'regions-v2.xlsx',
  totals: { inputs: 1, formulas: 0, external: 1, downstream: 4, outputs: 2, recalculated: 4, unexplained: 0, stale: 0 },
  edits: [
    { id: 'Inputs!B1', sheet: 'Inputs', cell: 'B1', label: 'Growth', kind: 'input', before: 0.1, after: 0.12, delta: 0.02, ratio: 0.2, impact: 3, … },
    { id: 'Inputs!B2', label: 'USD/INR', kind: 'external', external: { provider: 'Bloomberg', function: 'BDP' },
      formula: { before: '=BDP("USDINR Curncy","PX_LAST")', after: '=BDP("USDINR Curncy","PX_LAST")' }, before: 83, after: 84, impact: 2, … },
  ],
  downstream: [
    { id: 'Model!C5', label: 'India › Profit (USD) · 2025', labels: { sections: ['India'], row: ['Profit (USD)'], column: ['2025'] },
      kind: 'computed', before: 55, after: 55.333…, method: 'recalculated', output: true, causeCount: 2,
      causes: [{ id: 'Inputs!B1', label: 'Growth', contribution: 0.988… }, { id: 'Inputs!B2', label: 'USD/INR', contribution: -0.666… }], … },
    …
  ],
  structure: [], // e.g. { sheet: 'Revenue', kind: 'rows-added', rows: [6], labels: ['LATAM'], impact: 10 }
  notes: ['1 value came from outside data …'],
}
```

Every field is described in the [API reference](api.md#spreadiff-lite).
`report.diff` and `report.trace` hold the underlying `spreadiff` diff and
trace. They aren't serialised.

## How cells are named

`Model!C5` means nothing to a reader; `India › Profit (USD) · 2025` does.
spreadiff-lite names each cell from the layout around it, as a person
would read the sheet:

```text
       A                B         C
 1                      2024      2025       ← column headings
 2  India                                    ← a section heading
 3    Revenue           8300      9296       ← C3: India › Revenue · 2025
 4    Costs             4150      4648
 5  China
 6    Revenue            800       896       ← C6: China › Revenue · 2025
```

The rules:
1. **Label columns** are the columns left of the first one that holds
   mostly numbers. A row's label is its first one or two text cells there.
   In a table matched by an id column, the id leads.
2. **Section headings** are rows with a label and nothing to its right.
   Their nesting comes from, in order of trust:
   1. the cell's indent (Excel's *Increase indent*) or leading spaces;
   2. bold text;
   3. without either, layout: a heading directly above another heading is
      its parent (`India` over `Revenue`), and the next heading after data
      starts a sibling.
   
   A `Total …` row closes its section ("Total India" closes India).
3. **Column headings** are the nearest rows above the cell that hold only
   text, dates or years to the right of the labels, or the header row. A
   year or date row must run on (2024, 2025, or month-ends), so a row of
   numbers that happen to look like years isn't taken for one. In a stack
   of headings, an upper one spans to its right until the next one, as a
   merged or centred title does (`Actual › Q2`). At most the three nearest
   heading rows are used.

The same labels are available on their own: `tableLabeler(table)(row,
col)` and `formatLabel` in `spreadiff/core`.

**It can guess wrong.** These are inferences from the layout. They are
not definitions in the file. Known failure modes:
- *Two levels with no cue.* Headings at the same depth with no indent,
  bold or Total rows. `Asia › India`, then later `Europe`, can come out as
  `Asia › Europe`. Indent the file's labels, or add total rows.
- *A blank data row.* A labelled row with no values yet reads as a
  heading, so the rows under it inherit its label.
- *Free text to the right.* A notes column in the middle of a model can
  stop a row being read as a heading.

When the label is wrong, the address (`Model!C5`) is still exact. Every
row of the report has both.

## Outside data

Some formulas fetch their value from outside the workbook. They can't be
recalculated without the connection, and their value can change when
nothing in the file did. spreadiff-lite treats such a cell as an **input**:
- its saved value is taken as given;
- when it changed, it's an edit of kind `external`, with its provider;
- the results it feeds are traced to it like any other input.

Recognised, with or without Excel's `_xll.` prefix:

| Provider | Functions |
| --- | --- |
| Bloomberg | `BDP`, `BDH`, `BDS`, `BQL`, `BEQS`, `BSRCH`, `BCURVE`, `BLP`, `BLPH`, `BLPW` |
| LSEG (Refinitiv) | `TR`, `RDP.DATA`, `RDP.HISTORY`, `RDP.SEARCH`, `RHISTORY`, `RDATA`, `RSEARCH`, `DSGRID` |
| FactSet | `FDS`, `FDSB`, `FDSC`, `FQL` |
| S&P Capital IQ | `CIQ…` (any function starting with CIQ), `SPG` |
| Morningstar | `MSDP`, `MSTS` |
| Excel | `RTD`, `WEBSERVICE`, `STOCKHISTORY`, the `CUBE…` functions, `PY` (Python in Excel), `COPILOT`, `IMAGE`, `TRANSLATE`, `DETECTLANGUAGE` |
| any add-in | any other function written with Excel's `_xll.` prefix |
| link | a reference to another workbook (`[1]Sheet1!A1`) |

The `_xll.` rule catches XLL add-ins this table doesn't name. Visible
Alpha's functions, for one, weren't verified, and are only caught if the
file writes them that way.

Add your own with `liveFunctions` (or `--live` on the command line):

```ts
import { compare } from 'spreadiff-lite';
declare const a: { name: string; data: Uint8Array }, b: typeof a;

compare(a, b, { trace: { liveFunctions: { 'Visible Alpha': ['VA*'], 'In-house': ['PRICE_OF'] } } });
```

A name ending in `*` matches every function starting with it.
`externalSource(formula)` tells you what a formula would be classified as.

Not covered:
- Excel's linked data types (`=A1.Price` on a Stocks cell);
- VBA user-defined functions. The engine can't evaluate them, so causes
  through them come from references, and they are never treated as
  outside data.

## With an agent

The report is built for the end of an agent's turn. The agent edits the
workbook, saves it, runs spreadiff-lite against the version it started
from, and includes the Markdown in its reply. A reader then sees:
- what the agent changed;
- what that moved;
- anything it didn't expect, such as a final output that moved and
  shouldn't have, or a stale file.

```ts
import { readFileSync } from 'node:fs';
import { compare, toMarkdown } from 'spreadiff-lite';

const before = { name: 'model.xlsx', data: readFileSync('backup/model.xlsx') };
const after = { name: 'model.xlsx', data: readFileSync('model.xlsx') };
const report = compare(before, after);

// Check the edit did what was intended before reporting it.
const moved = new Set(report.downstream.filter((c) => c.output).map((c) => c.id));
if (!moved.has('Summary!C2')) throw new Error('The edit did not reach the total it was meant to change');
const unexpected = report.downstream.filter((c) => c.output && c.sheet === 'Budget');
console.log(toMarkdown(report, { maxRows: 30 }) + (unexpected.length ? `\nAlso moved ${unexpected.length} budget lines.` : ''));
```

For coding agents that load skills, [`skills/spreadiff/SKILL.md`](../skills/spreadiff/SKILL.md)
packages this workflow: when to run it, what to check in the report, and
the limits to mention. Copy it into the agent's skills directory.

Tips:
- **Keep it short.** `--outputs-only` (or `intermediate: false`) lists
  only final outputs; `maxRows` caps each table and counts the rest.
  Either keeps a large change within a reply.
- **Recalculate the file too.** If the tool that wrote the file doesn't
  calculate (openpyxl, most scripts), spreadiff computes the missing
  results and says so in *Notes*.
- **Use the exit code.** 0, 1 and 2 let a CI step fail on any change, or
  on none.

## Command line

```text
spreadiff-lite OLD NEW [options]
```

| Option | |
| --- | --- |
| `--json` | Print the report as JSON instead of Markdown |
| `--all` | Don't cap the tables (default: 50 rows each) |
| `--max-rows N` | Rows per table |
| `--outputs-only` | List final outputs, not intermediate results |
| `--sheet NAME` | Only this sheet (repeat for more) |
| `--keys COL[,COL]` | Match rows on these columns |
| `--header auto\\|yes\\|no` | Whether row 1 holds column names (default `auto`) |
| `--no-recalc` | Follow formula references instead of recalculating: faster, may over-attribute, no contributions |
| `--live NAME[,NAME]` | More data-feed functions to treat as inputs, as `Provider:FN,FN` or `FN` |
| `--tolerance X` | Numbers within X of each other are equal |
| `--min-change P` | Ignore moves smaller than this share of the number (`0.1%` or `0.001`): a materiality threshold |
| `--textconv FILE` | Print a workbook as line-per-row text, for git ([Spreadsheets in git](git.md)) |
| `--git …` | Act as git's diff driver; `spreadiff-lite merge BASE OURS THEIRS` is the merge driver ([Spreadsheets in git](git.md)) |

Exit status: 0 no changes, 1 changes, 2 error.

## How it can be wrong

- **Labels are inferred** from the layout, not read from the file. See
  [How cells are named](#how-cells-are-named). The address next to each
  label is exact.
- **Contributions are one edit at a time.** Each is "the result without
  that edit", with every other edit in place. They describe what each edit
  did given the others. They are not shares of a total, and they needn't
  sum to the change.
- **One recalculation engine.** It's spreadiff's own, checked against
  LibreOffice, not Excel. Where it can't reproduce a formula's saved value,
  that part of the trace falls back to references, and the report says so.
- **Outside data is trusted as saved.** A Bloomberg value that refreshed
  is an input. If the file was saved before the feed updated, the report
  can't know.
- **Final outputs are relative to what changed.** "Final" means no other
  *changed* cell reads it. A cell that feeds an unchanged cell elsewhere
  still counts as final here.
