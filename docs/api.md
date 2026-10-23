# API reference

Every public export, grouped by task. The same symbol is often available
from several entry points: import from the smallest one that has it.

| Entry point | Contents |
| --- | --- |
| `spreadiff` | Everything below except React, lite and the worker client |
| `spreadiff/core` | Diffing, parsing (CSV, XLSX), tracing, formulas, transfer, utilities. No DOM |
| `spreadiff/lite` | `diffTables`, `parseCsv`, text utilities, and the HTML/text renderers |
| `spreadiff/xlsx` | `parseXlsx` only |
| `spreadiff/react` | The `DiffGrid` React component and its types |
| `spreadiff/worker` | `createDiffWorker` and its types |
| `spreadiff/worker/worker.js` | The worker script itself (load it with `new Worker`) |

Coordinates used throughout:
- **aligned row**: an index into a `TableDiff`'s rows (`0 … rowCount - 1`);
- **old / new row**: an index into `oldTable.rows` / `newTable.rows`;
- **aligned column**: an index into `TableDiff.columns`.

All indexes are 0-based. Display row numbers (2, 3, …) are derived from
`Table.origin`.

---

## Diffing

### diffTables

```ts signature
diffTables(oldTable: Table, newTable: Table, options?: DiffOptions): TableDiff
```

Diff two tables: match columns, match rows, compare cells. See
[How diffing works](diffing.md). Formulas referencing other sheets are
compared textually; use [`diffWorkbooks`](#diffworkbooks) for
cross-sheet-aware comparison.

From `spreadiff`, `spreadiff/core`, `spreadiff/lite`.

### diffWorkbooks

```ts signature
diffWorkbooks(oldBook: Workbook, newBook: Workbook, options?: DiffOptions): WorkbookDiff
```

Diff every sheet. Sheets pair by name (two single-sheet workbooks pair
regardless of name). Unpaired sheets become `added` / `removed` entries
diffed against an empty table. All sheets are aligned before cells are
compared, so formulas referencing other sheets are compared through those
sheets' alignments.

From `spreadiff`, `spreadiff/core`.

### hasChanges

```ts signature
hasChanges(diff: TableDiff): boolean
```

True when any row was added, removed, modified or moved, or any column
added, removed or renamed.
From `spreadiff`, `spreadiff/core`, `spreadiff/lite`.

### cellsEqual

```ts signature
cellsEqual(a: CellValue, b: CellValue, options?: DiffOptions): boolean
```

The cell comparison the diff uses, with the same `numericTolerance`,
`trimWhitespace` and `ignoreCase` options. Formulas aren't considered.

### DiffOptions

Options for `diffTables` and `diffWorkbooks`.

| Field | Type | Default | |
| --- | --- | --- | --- |
| `keys` | `string[] \| 'auto' \| 'none'` | `'auto'` | How rows are matched; see [Step 2: rows](diffing.md#step-2-rows) |
| `numericTolerance` | `number` | `0` | Numbers within this absolute difference are equal |
| `precision` | `number` | `15` | Significant digits two numbers must share (one unit in the last place allowed). 15 hides Excel-vs-LibreOffice float noise such as `0.9382716049382716` vs `0.938271604938272`; 17 compares doubles exactly |
| `trimWhitespace` | `boolean` | `true` | Ignore leading/trailing whitespace in text |
| `ignoreCase` | `boolean` | `false` | Compare text case-insensitively |
| `compareFormulas` | `boolean` | `true` | Count formula-meaning changes (when both sides have formulas) |
| `rowSimilarity` | `number` | `0.5` | Share of equal non-formula cells needed to pair a removed and an added row as modified |
| `detectRenames` | `boolean` | `true` | Pair look-alike removed/added columns (≥ 80% matching values) as renames |

### TableDiff

The result of diffing two tables.

| Field | Type | |
| --- | --- | --- |
| `oldTable`, `newTable` | `Table` | The inputs |
| `columns` | `AlignedColumn[]` | Aligned columns, in display order |
| `rowCount` | `number` | Number of aligned rows |
| `oldRows` | `Int32Array` | Old row per aligned row, `-1` if added |
| `newRows` | `Int32Array` | New row per aligned row, `-1` if removed |
| `status` | `Uint8Array` | Row status per aligned row, as [`ROW_STATUS`](#row_status) numbers |
| `moved` | `Uint8Array` | `1` where a row matched by key moved: the fewest rows whose moving explains the new order. Independent of `status` (a moved row can also be modified). Always `0` when rows were matched by order. See [Moved rows](diffing.md#moved-rows) |
| `changedCells` | `Map<number, Uint8Array>` | Aligned row → mask over `columns`, `1` where the cell changed. Only modified rows |
| `formulaChanged` | `Map<number, Uint8Array>` | Same shape: cells whose formula changed in meaning |
| `keyColumns` | `string[]` | Key columns used; empty when matched by order |
| `stats` | `DiffStats` | Counts |
| `options` | `Required<DiffOptions>` | Options used, defaults filled in |

### AlignedColumn

| Field | Type | |
| --- | --- | --- |
| `oldIndex` | `number` | Column in the old table, `-1` if added |
| `newIndex` | `number` | Column in the new table, `-1` if removed |
| `name` | `string` | Current name (old name for removed columns) |
| `oldName` | `string?` | Previous name, when `renamed` |
| `status` | `ColumnStatus` | |
| `changes` | `number` | Changed cells in the column |

### DiffStats

`rowsAdded`, `rowsRemoved`, `rowsModified`, `rowsEqual`, `rowsMoved`,
`cellsChanged`, `columnsAdded`, `columnsRemoved`, `columnsRenamed`, all
numbers. `cellsChanged` counts changed cells in modified rows only; cells
of added/removed rows and columns aren't included. `rowsMoved` overlaps
`rowsEqual` and `rowsModified`: a moved row is also one of those.

### RowStatus

`'equal' | 'added' | 'removed' | 'modified'`

### ColumnStatus

`'equal' | 'added' | 'removed' | 'renamed'`

### ROW_STATUS

`{ equal: 0, added: 1, removed: 2, modified: 3 }`: the numeric encoding in
`TableDiff.status`. Compare against it in hot loops:
`diff.status[r] === ROW_STATUS.modified`.

### ROW_STATUS_NAMES

`['equal', 'added', 'removed', 'modified']`: decode a status number,
`ROW_STATUS_NAMES[diff.status[r]]`.

### SheetDiff

| Field | Type | |
| --- | --- | --- |
| `name` | `string` | Sheet name |
| `status` | `'equal' \| 'modified' \| 'added' \| 'removed'` | |
| `diff` | `TableDiff` | The sheet's diff |

### WorkbookDiff

| Field | Type | |
| --- | --- | --- |
| `sheets` | `SheetDiff[]` | New-workbook order, removed sheets where they were |
| `names` | `{ old?, new? }?` | Defined names of each workbook, for tracing |
| `tables` | `{ old?, new? }?` | `TableDefinition`s of each workbook, for structured references |
| `date1904` | `{ old?, new? }?` | Each workbook's date system, for recalculation |
| `computedValues` | `{ old?, new? }?` | Each workbook's `computedValues` |

---

## Data

### CellValue

`string | number | boolean | null`. Parsers produce numbers only when the
conversion is lossless (`"007"` stays text). Dates from XLSX are ISO
strings: `"2026-01-31"`, `"2026-01-31 14:05:00"` or `"14:05:00"`.

### Table

| Field | Type | |
| --- | --- | --- |
| `name` | `string` | Sheet name; pairs sheets and resolves `Sheet!A1` refs |
| `columns` | `string[]` | Column names |
| `rows` | `CellValue[][]` | Data rows, without the header |
| `formulas` | `(string \| null)[][]?` | Formula text per cell (`"=B2*C2"`), parallel to `rows`; rows may be missing. Set (even empty) for any source that can hold formulas |
| `origin` | `{ row, col }?` | Sheet position of `rows[0][0]`, 0-based. Default `{ row: 1, col: 0 }` (A2) |
| `header` | `CellValue[]?` | The header row's original cells, when the first row became column names. Formula evaluation reads them on the row above `origin.row` |
| `arrays` | `{ r0, c0, r1, c1 }[]?` | Array-formula and spill blocks, 0-based inclusive sheet coordinates. Every cell of a block holds the block's formula |
| `merges` | `{ r0, c0, r1, c1 }[]?` | Merged ranges, 0-based inclusive sheet coordinates. The value is in the top-left cell. `DiffGrid` draws an unchanged horizontal merge as one cell |
| `names` | `Record<string, string>?` | Defined names scoped to this sheet. Its formulas see them over `Workbook.names`; other sheets reach them as `Sheet!Name`. See `scopedNames` |

Build tables by hand when your data isn't a file:

```ts
import { diffTables, type Table } from 'spreadiff';

const toTable = (rows: { sku: string; qty: number }[]): Table => ({
  name: 'stock',
  columns: ['sku', 'qty'],
  rows: rows.map((r) => [r.sku, r.qty]),
});
const d = diffTables(toTable([{ sku: 'A1', qty: 3 }]), toTable([{ sku: 'A1', qty: 4 }]), { keys: ['sku'] });
```

### Workbook

| Field | Type | |
| --- | --- | --- |
| `name` | `string` | Usually the file name |
| `sheets` | `Table[]` | In workbook order |
| `names` | `Record<string, string>?` | Defined names: `TaxRate` → `Inputs!$B$2` |
| `tables` | `TableDefinition[]?` | Excel tables, to resolve structured references (`Sales[Amount]`) |
| `date1904` | `boolean?` | Dates count from 1904-01-01 (old Mac Excel) |
| `computedValues` | `MissingValues?` | Set when formulas without cached values were computed on parsing |

### TableDefinition

An Excel table (*Insert → Table*), as structured references see it.

| Field | Type | |
| --- | --- | --- |
| `name` | `string` | Name in formulas: `Sales` |
| `sheet` | `string` | Sheet it's on |
| `range` | `{ r0, c0, r1, c1 }` | Whole table, header and totals rows included; 0-based inclusive sheet coordinates |
| `columns` | `string[]` | Column names, left to right |
| `headerRows` | `number` | 0 or 1 |
| `totalsRows` | `number` | 0 or 1 |

---

## Parsing

### parseFile

```ts signature
parseFile(name: string, data: ArrayBuffer | Uint8Array, options?: ParseFileOptions): Workbook
```

Parse a file by content and name:
- zip content → XLSX;
- `.tsv` → tab-delimited;
- otherwise → CSV, decoded as UTF-8, or UTF-16 when it has a byte-order mark.

Throws for legacy `.xls` files and for files named `.xlsx` that aren't zips.

### ParseFileOptions

| Field | Type | Default | |
| --- | --- | --- | --- |
| `header` | `boolean \| 'auto'` | `'auto'` | First (non-empty) row holds column names. XLSX: see `ParseXlsxOptions.header`; CSV treats `'auto'` as `true` |
| `delimiter` | `string` | detected | CSV delimiter |

### parseCsv

```ts signature
parseCsv(text: string, options?: ParseCsvOptions): Table
```

RFC 4180 CSV:
- **quoting**: quoted fields, `""` escapes and line breaks inside quotes; text
  after a closing quote is kept (`"ab"cd` → `abcd`);
- **line endings**: CRLF, LF and CR;
- **delimiter**: honours Excel's `sep=;` first line;
- **cleanup**: a leading BOM is stripped, trailing empty rows are dropped,
  empty fields become `null`.

### ParseCsvOptions

| Field | Type | Default | |
| --- | --- | --- | --- |
| `delimiter` | `string` | detected | See [`detectDelimiter`](#detectdelimiter) |
| `header` | `boolean` | `true` | First row holds column names (made unique; blanks become letters) |
| `numbers` | `boolean \| 'lossless'` | `'lossless'` | `'lossless'` converts only fields that print back identically (`"12.5"` yes, `"007"` / `"1.50"` / `"1e5"` no); `true` converts anything numeric; `false` keeps strings |
| `name` | `string` | `'Sheet1'` | Table name |

### detectDelimiter

```ts signature
detectDelimiter(text: string): string
```

Guess `,`, tab, `;` or `|` from the first lines: prefers the one that
appears the same number of times on each line. Defaults to `,`.

### parseXlsx

```ts signature
parseXlsx(data: ArrayBuffer | Uint8Array, options?: ParseXlsxOptions): Workbook
```

Read `.xlsx` / `.xlsm`. It reads:
- cached values and formula text, with shared and array formulas expanded
  per cell. Formulas with no cached value are computed (see
  `computeMissing`); cached values are never replaced;
- dates (by number format, including the 1904 date system and Excel's
  1900 leap-year quirk), booleans, errors (`"#DIV/0!"`) and rich text;
- workbook-scoped defined names, Excel tables and the date system.

Each sheet is trimmed to its used range, and `origin` records where it
starts. Styles, merged cells, charts and comments are ignored. Scanning is
linear in the file size, even for malformed files.

From `spreadiff`, `spreadiff/core`, `spreadiff/xlsx`.

### ParseXlsxOptions

| Field | Type | Default | |
| --- | --- | --- | --- |
| `header` | `boolean \| 'auto'` | `'auto'` | Whether each sheet's first non-empty row holds column names. `'auto'`: yes, unless it has numbers (years next to labels excepted), dates or formulas, or a formula reads one of its cells. Then it's data, as suits a financial model, and columns are named from its labels if it's mostly labels, else lettered. See [Step 1: columns](diffing.md#step-1-columns) |
| `sheets` | `string[]` | all | Only parse these sheets |
| `name` | `string` | `'workbook.xlsx'` | Workbook name |
| `computeMissing` | `boolean` | `true` | Compute formulas the file has no value for (files from openpyxl and other tools that don't calculate) |

### computeMissingValues

```ts signature
computeMissingValues(workbook: Workbook): MissingValues
```

Compute, in place, every formula cell whose value is `null`. Values the
file has are never changed. A cell that depends on one that couldn't be
computed fails too, rather than reading it as blank. `parseXlsx` calls
this unless `computeMissing` is false. See [Stale and missing
values](tracing.md#stale-and-missing-values).

### MissingValues

`{ computed, failed }`: formula cells without a cached value that were
computed, and that couldn't be (unsupported function, circular
reference) and are left empty.

### uniqueNames

```ts signature
uniqueNames(names: string[]): string[]
```

Make column names unique and non-empty, as the parsers do: blanks become
column letters, and repeats become `x (2)`, `x (3)`, … without colliding with
existing names.

### columnLabel

```ts signature
columnLabel(index: number): string
```

Spreadsheet column letters for a 0-based index: `0` → `A`, `25` → `Z`, `26` → `AA`.

---

## Tracing

### traceWorkbook

```ts signature
traceWorkbook(diff: WorkbookDiff, options?: TraceOptions): WorkbookTrace
```

Classify every changed cell as `input`, `formula` or `computed`, and find
each computed change's root causes and each root change's impacts. See
[Tracing downstream changes](tracing.md).

### traceTable

```ts signature
traceTable(diff: TableDiff, options?: TraceOptions): WorkbookTrace
```

`traceWorkbook` for a single table diff (it becomes sheet 0).

### TraceOptions

| Field | Default | |
| --- | --- | --- |
| `maxCauses` | `25` | Causes kept per cell (`causeCount` stays exact) |
| `maxImpacts` | `500` | Impacts kept per root change (`impactCount` stays exact) |
| `budget` | `5_000_000` | Graph steps before stopping and setting `truncated` |
| `recalculate` | `true` | Recalculate formulas for exact causes, contributions and stale values. `false`: causes from references only |
| `evaluationBudget` | `2_000_000` | Formula evaluations for recalculation, across the trace. Past it, the rest falls back to references |

### WorkbookTrace

| Member | |
| --- | --- |
| `get(ref): CellTrace \| undefined` | The trace of a changed cell; `undefined` if it didn't change |
| `parents(ref): CellRef[]` | Changed cells the cell's formula reads directly (one level up), including cells of added/removed rows in its ranges. At most 200 |
| `sheet(index): ReadonlyMap<number, CellTrace>` | Every traced cell of a sheet, keyed by `row * columns.length + col` |
| `stats: TraceStats` | Totals |
| `truncated: boolean` | The budget ran out; some causes may be missing |

### CellTrace

| Field | Type | |
| --- | --- | --- |
| `kind` | `ChangeKind` | |
| `causes` | `CellRef[]` | Computed cells: root changes they depend on (capped) |
| `causeCount` | `number` | Exact number of causes |
| `impacts` | `CellRef[]` | Root changes: computed cells they moved (capped) |
| `impactCount` | `number` | Exact number of impacts |
| `unexplained` | `'volatile' \| 'dynamic' \| 'unknown'?` | Computed cells without a traced cause: why |
| `contributions` | `(number \| null)[]?` | Parallel to `causes`, when `method` is `'recalculated'`: new value − value without that cause. Parts needn't add up to the total |
| `method` | `'recalculated' \| 'references'?` | How `causes` was found: by undoing each edit and recalculating (exact), or by following references (may over-include) |
| `stale` | `{ cached, recalculated }?` | The new file's saved value is out of date: `cached` is what the file says, `recalculated` what its formulas give |

### CellRef

`{ sheet, row, col }`: sheet index into `WorkbookDiff.sheets`, aligned row,
aligned column index.

### ChangeKind

`'input' | 'formula' | 'computed' | 'added' | 'removed'`. See [Four kinds
of change](tracing.md#four-kinds-of-change).

### TraceStats

`{ inputs, formulas, computed, unexplained, stale, recalculated, unmodelled }`:
counts of changed cells by kind; `unexplained` counts computed cells with
no traced cause, `stale` those whose saved value was out of date,
`recalculated` those whose causes came from recalculation, and
`unmodelled` formulas near a change the engine couldn't reproduce.

### applyRecalculation

```ts signature
applyRecalculation(diff: WorkbookDiff, trace: WorkbookTrace): WorkbookDiff
```

The diff as it would be had the new file been recalculated: stale cells
take their recalculated values and count as changed. Alignment is
unchanged, so `trace` still applies. Returns `diff` itself when nothing
was stale; never modifies it.

### changeKind

```ts signature
changeKind(diff: TableDiff, row: number, col: number): ChangeKind
```

Classify one changed cell without a full trace: `formula` if its formula
changed, `computed` if it has the same formula on both sides, else
`input`. This is what the renderers use when no trace is attached.

---

## Formulas

### formulaRefs

```ts signature
formulaRefs(formula: string, names?: Record<string, string>, context?: FormulaContext): FormulaRefs
```

What a formula reads. `names` resolves defined names (pass
`Workbook.names`); `context` resolves structured references and 3-D
references.

<!-- example: formula-refs -->
```ts
import { formulaRefs } from 'spreadiff';

const r = formulaRefs("=SUM(Lines!D2:D9)*TaxRate+'Q&A'!B3", { TaxRate: 'Inputs!$B$4' });
console.log(r.ranges.map((x) => `${x.sheet}:${x.r0},${x.c0}-${x.r1},${x.c1}`).join(' '), r.dynamic, r.volatile);
```

<!-- output: formula-refs -->
```text
Lines:1,3-8,3 Q&A:2,1-2,1 Inputs:3,1-3,1 false false
```

### FormulaRefs

| Field | Type | |
| --- | --- | --- |
| `ranges` | `SheetRange[]` | Every range read, including through names |
| `dynamic` | `boolean` | Precedents can't be fully known (INDIRECT, OFFSET, unknown names, external links, `#REF!`) |
| `volatile` | `boolean` | Uses NOW, TODAY, RAND, RANDBETWEEN or RANDARRAY |

### FormulaContext

| Field | Type | |
| --- | --- | --- |
| `tables` | `TableDefinition[]?` | The workbook's Excel tables, for `Sales[Amount]` |
| `row` | `number?` | 0-based sheet row of the formula's cell, for `[@Qty]` / `[#This Row]` |
| `sheets` | `string[]?` | Sheet names in workbook order, to expand `Jan:Mar!B2` |

Without it, structured and 3-D references make the result `dynamic`.

### structuredRange

```ts signature
structuredRange(table: string, spec: string, context?: FormulaContext): SheetRange | null
```

Resolve a structured reference to a sheet range: `structuredRange('Sales',
'[[#Totals],[Amount]]', ctx)`. Handles `#All`, `#Data`, `#Headers`,
`#Totals`, `#This Row` / `@` and column spans. `null` when the table or a
column isn't known.

### scopedNames

```ts signature
scopedNames(workbookNames: Record<string, string> | undefined, sheets: readonly { name: string; names?: Record<string, string> }[], sheet: string): Record<string, string> | undefined
```

The defined names a formula on `sheet` sees: the workbook's, then that
sheet's own (`Table.names`, which win), plus every sheet's own as
`Sheet!Name`. Pass the result as `names` to `formulaRefs`:

```ts
import { formulaRefs, scopedNames, type Workbook } from 'spreadiff';
declare const wb: Workbook;

const names = scopedNames(wb.names, wb.sheets, 'Inputs');
const refs = formulaRefs('=A1*Rate', names); // Inputs' own Rate, if it has one
```

### SheetRange

`{ sheet, r0, c0, r1, c1 }`: 0-based, inclusive sheet coordinates. `sheet`
is `null` for the formula's own sheet. Whole-column ranges (`C:C`) have
`r0 = r1 = -1`; whole-row ranges (`3:3`) have `c0 = c1 = -1`.

### tokenizeFormula

```ts signature
tokenizeFormula(formula: string): { tokens: FormulaToken[]; dynamic: boolean; volatile: boolean; names: string[] }
```

Split a formula into reference tokens and everything else (string literals
kept intact, whitespace dropped). The building block for `formulaRefs` and
`canonicalFormula`.

### FormulaToken

`{ type: 'text'; text } | { type: 'ref'; text; range: SheetRange }`

### canonicalFormula

```ts signature
canonicalFormula(formula: string, map: (range: SheetRange) => string): string
```

The formula with every reference replaced by `map(range)`. Two formulas
with equal canonical forms read the same cells. The diff uses this, with a
`map` that translates sheet coordinates into aligned rows, to compare
formulas by meaning.

---

## Canvas rendering

### DiffGrid

```ts signature
new DiffGrid<A = unknown>(options?: DiffGridOptions<A>)
```

The interactive canvas grid; see [The DiffGrid](grid.md) for a guide.

| Member | |
| --- | --- |
| `render({ diff, container })` | Mount into `container` and show `diff` |
| `setDiff(diff)` | Show another diff |
| `setOptions(partial)` | Change options (cheap when nothing layout-related changed) |
| `getOptions()` | Current options |
| `nextChange()` / `prevChange()` | Select the next / previous block of changes; `false` at the end |
| `scrollToRow(row, align?)` | Scroll an aligned row into view (`'start' \| 'center' \| 'nearest'`) |
| `select(row, col?)` / `select(null)` | Select a cell; `false` if out of range |
| `expandAll()` / `collapseAll()` | Unfold / refold unchanged regions |
| `cellInfo(row, col)` | A cell's [`CellInfo`](#cellinfo); throws `RangeError` if it doesn't exist |
| `refreshTheme()` | Re-read CSS variables |
| `drawNow()` | Paint synchronously (tests, benchmarks) |
| `cleanUp()` | Unmount and release observers |
| `host` | The element inserted into your container |

### DiffGridOptions

All optional; see the [options table](grid.md#options) for defaults and
behavior: `layout`, `cellDisplay`, `collapseUnchanged`, `context`,
`hideUnchangedColumns`, `frozenColumns`, `columnWidths`, `rowNumberOffset`,
`theme`, `minimap`, `tooltips`, `highlightComputed`, `trace`, `traceFocus`,
`annotations`, `renderAnnotation`, `onCellClick`, `onCellHover`,
`onSelectionChange`, `accessibility`, `ariaLabel`.

### DiffGridRenderProps

`{ diff: TableDiff; container: HTMLElement }`: the argument to `render`.

### DiffLayout

`'unified' | 'split'`

### CellDisplay

`'stacked' | 'new' | 'delta' | 'inline'`. See [Displaying a
change](grid.md#displaying-a-change).

### CellInfo

| Field | Type | |
| --- | --- | --- |
| `row` | `number` | Aligned row |
| `column` | `AlignedColumn` | |
| `columnIndex` | `number` | Aligned column index |
| `status` | `RowStatus` | Row status |
| `changed` | `boolean` | The cell changed |
| `moved` | `boolean` | The row moved (see `TableDiff.moved`) |
| `oldRow`, `newRow` | `number` | Row indexes, `-1` when missing |
| `oldValue`, `newValue` | `CellValue \| undefined` | `undefined` when the cell doesn't exist on that side |
| `oldFormula`, `newFormula` | `string \| null \| undefined` | Formula text, `null` if none |
| `kind` | `ChangeKind?` | For changed cells |
| `trace` | `CellTrace?` | With a trace attached |

### CellAnnotation

`{ row: number; column?: string; data: A; id?: string | number }`: a note
rendered below aligned row `row`; `column` also marks that cell.
Annotations are matched across `setOptions` calls by `id`, or else by row
and column, so a fresh array of equal annotations keeps its elements; one
is re-rendered only when its `data` changes (`===`). See
[Annotations](grid.md#annotations).

### TraceContext

`{ trace: WorkbookTrace; workbook: WorkbookDiff; sheet: number }`: the
`trace` option; `sheet` is the grid's sheet within `workbook`.

---

## React

### DiffGrid (React)

```ts signature
<DiffGrid diff={…} {...DiffGridOptions} className? style? ref? />
```

From `spreadiff/react`. Renders a container `<div>` with the canvas grid
inside; option props update the grid in place. `renderAnnotation` may
return JSX (portaled into place).

### DiffGridProps

`DiffGridOptions` plus `diff`, `className`, `style`, and a React-flavored
`renderAnnotation: (a) => ReactNode`.

### DiffGridHandle

The `ref` value: `nextChange`, `prevChange`, `scrollToRow`, `expandAll`,
`collapseAll`, `select`, `refreshTheme`, `cellInfo`, and `instance` (the
underlying `DiffGrid`, or `null` before mount).

---

## Lite rendering

From `spreadiff/lite`. See [spreadiff/lite](lite.md) for a guide.

### renderHTML

```ts signature
renderHTML(diff: TableDiff, options?: LiteOptions): string
```

The diff as an HTML string (changes plus context). Cell content is escaped.
Style it with [`LITE_CSS`](#lite_css).

### renderText

```ts signature
renderText(diff: TableDiff, options?: { context?; changedColumnsOnly?; maxRows?; maxWidth? }): string
```

An aligned plain-text diff with ASCII `+` / `-` line markers, for terminals,
logs and ```` ```diff ```` blocks. `maxWidth` (default 24) caps column width.

### LiteOptions

| Field | Default | |
| --- | --- | --- |
| `layout` | `'split'` | `'split'` or `'unified'` |
| `display` | `'stacked'` | How changed cells read |
| `context` | `2` | Unchanged rows around each change |
| `changedColumnsOnly` | `false` | Only changed columns, plus keys (or the row-label column when rows were matched by order) |
| `maxRows` | `2000` | Rows of output before truncating |
| `kindOf` | — | `(row, col) => ChangeKind \| undefined` override |
| `rowNumberOffset` | from `origin` | Displayed number of row index 0 |

### LiteLayout

`'split' | 'unified'`

### LiteDisplay

`'stacked' | 'new' | 'delta' | 'inline'`

### LITE_CSS

A CSS string for `renderHTML` output, scoped under `.sdl-wrap` and themed
by the `--sd-*` variables. Include it once per page.

---

## Values and text

### formatCell

```ts signature
formatCell(v: CellValue): string
```

Display text: `''` for `null`, `TRUE` / `FALSE`, and numbers at 15
significant digits (`0.1 + 0.2` → `"0.3"`).

### numericDelta

```ts signature
numericDelta(before: CellValue, after: CellValue): { delta: number; ratio: number | null } | null
```

`after − before` and its ratio to `|before|` (`null` when `before` is 0),
or `null` unless both read as finite numbers. `numericDelta(49, 52)` →
`{ delta: 3, ratio: 0.0612… }`.

### diffText

```ts signature
diffText(before: string, after: string): TextSpan[]
```

A character-level diff (word-level for strings over 40 characters):
`equal`, `removed` and `added` runs that rebuild both strings.

### TextSpan

`{ text: string; kind: 'equal' | 'added' | 'removed' }`

---

## Export

### toUnifiedCsv

```ts signature
toUnifiedCsv(diff: TableDiff): string
```

The diff as a CSV readable in any spreadsheet app, in daff's "highlighter"
style:
- a first column of `+++` (added), `---` (removed), `->` (modified) or
  `:` (moved, otherwise unchanged);
- changed cells as `old->new`.

Cells that would run as formulas when opened (`=`, `+`, `-`, `@`, except
plain numbers) are prefixed with `'`.

<!-- example: unified-csv -->
```ts
import { diffTables, parseCsv, toUnifiedCsv } from 'spreadiff';

console.log(toUnifiedCsv(diffTables(parseCsv('id,v\n1,a\n2,b'), parseCsv('id,v\n1,z\n3,c'))));
```

<!-- output: unified-csv -->
```text
@@,id,v
->,1,a->z
---,2,b
+++,3,c
```

---

## Workers

From `spreadiff/worker` unless noted. See [Workers](workers.md).

### createDiffWorker

```ts signature
createDiffWorker(worker: Worker): DiffWorker
```

Wrap a module `Worker` running `spreadiff/worker/worker.js`.

### DiffWorker

| Member | |
| --- | --- |
| `diffFiles(oldFile, newFile, options?, settings?)` | Read, parse and diff two `File`s |
| `diffBuffers(old: FileData, new: FileData, options?, settings?)` | Same from bytes; buffers are transferred |
| `terminate()` | Stop the worker; pending requests reject |

### WorkerSettings

| Field | Type | Default | |
| --- | --- | --- | --- |
| `header` | `boolean \| 'auto'` | `'auto'` | Whether the first row holds column names; see `ParseFileOptions.header` |
| `trace` | `boolean \| TraceOptions` | `false` | Also trace downstream changes in the worker. See [Workers → Formula traces](workers.md#formula-traces) |

### FileData

`{ name: string; data: ArrayBuffer }`: the name picks the parser.

### WorkerDiffResult

`{ result: WorkbookDiff; trace?: WorkbookTrace; timings: { parse, diff, trace } }`:
`trace` when `settings.trace` asked for it and the files have formulas;
timings in milliseconds, measured inside the worker (`trace` is 0 when not
asked).

### WorkerRequest

The message format of `spreadiff/worker/worker.js`:
`{ id, old: FileData, new: FileData, options?, parse?, trace? }`. Replies are
`{ id, packed: PackedWorkbookDiff, trace?: PackedWorkbookTrace, timings }` or `{ id, error }`. Only
needed if you drive the worker yourself.

### packWorkbookDiff

```ts signature
packWorkbookDiff(d: WorkbookDiff): { packed: PackedWorkbookDiff; transfer: ArrayBuffer[] }
```

Encode a workbook diff for `postMessage(packed, transfer)`. From
`spreadiff`, `spreadiff/core`. The diff's typed arrays are transferred:
don't use `d` afterwards.

### unpackWorkbookDiff

```ts signature
unpackWorkbookDiff(p: PackedWorkbookDiff): WorkbookDiff
```

Decode it on the other side.

### packTable

```ts signature
packTable(t: Table): PackedTable
```

Encode one table (used by `packWorkbookDiff`).

### unpackTable

```ts signature
unpackTable(p: PackedTable): Table
```

Decode one table.

### PackedTable

`{ name, columns, width, height, types: Uint8Array, numbers: Float64Array, strings: string, formulas?, origin?, header?, arrays? }`:
cell types (0 null, 1 string, 2 number, 3 true, 4 false) and numbers
row-major, strings joined by U+0000. `formulas` is packed the same way:
`{ lengths: Int32Array, width, cells: Float64Array, text }`, the position
of each formula and their text joined. Nested arrays of 300,000 formula
strings took 150–330 ms to structured-clone; packed, 10–23 ms.

### PackedWorkbookDiff

A `WorkbookDiff` whose tables are `PackedTable`s.

### packWorkbookTrace

```ts signature
packWorkbookTrace(trace: WorkbookTrace, diff: WorkbookDiff): { packed: PackedWorkbookTrace; transfer: ArrayBuffer[] }
```

Encode a trace of `diff` for `postMessage(packed, transfer)`. The parent
lists of downstream and formula-edit cells are computed here, so
`parents()` of other cells is `[]` after unpacking.

### unpackWorkbookTrace

```ts signature
unpackWorkbookTrace(p: PackedWorkbookTrace, diff: WorkbookDiff): WorkbookTrace
```

Rebuild the trace on the other side. `diff` is the unpacked diff it was
made for.

### PackedWorkbookTrace

`{ sheets: PackedSheetTrace[], parents, stats, truncated }`.

### PackedSheetTrace

One sheet's traces as typed arrays: `keys`, `codes` (kind, method,
unexplained, has contributions), `counts` (cause and impact counts),
offsets `causeAt` / `impactAt` / `contribAt` into `refs` (sheet, row, col
triples) and `contributions` (NaN for null), plus `stale` values by
index. Structured-cloning one object per traced cell cost about 200 ms per
40,000 cells.

---

## Algorithms

### alignSequences

```ts signature
alignSequences(a: Int32Array, b: Int32Array): number[]
```

Longest-common-subsequence alignment of two integer sequences. Returns
matched index pairs flattened as `[a0, b0, a1, b1, …]`, increasing in both.
Optimal (Myers) for small edit distances; when the work budget runs out it
falls back to patience anchoring, which stays near-linear but may not be
minimal. The diff uses it on interned row contents.
