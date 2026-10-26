# API reference

Every public export, grouped by task. The same symbol is often available
from several entry points: import from the smallest one that has it.

| Entry point | Contents |
| --- | --- |
| `spreadiff` | Everything below except React, the HTML renderer and the worker client |
| `spreadiff/core` | Comparing, diffing, parsing (XLSX, XLSB, XLS, ODS, CSV), tracing, formulas, number formats, review and merge, transfer, utilities. No DOM |
| `spreadiff/html` | `diffTables`, `parseCsv`, text utilities, and the HTML/text renderers |
| `spreadiff/xlsx` | `parseXlsx`, and `applyDecisionsToXlsx` for writing a reviewed file |
| `spreadiff/react` | The `SpreadsheetDiff` and `DiffGrid` React components and their types |
| `spreadiff/worker` | `createDiffWorker` and its types |
| `spreadiff/worker/worker.js` | The worker script itself (load it with `new Worker`) |
| `spreadiff-lite` | A separate package: `compare`, `toMarkdown`, `summarize` and the report types ([below](#spreadiff-lite)) |

Coordinates used throughout:
- **aligned row**: an index into a `TableDiff`'s rows (`0 … rowCount - 1`);
- **old / new row**: an index into `oldTable.rows` / `newTable.rows`;
- **aligned column**: an index into `TableDiff.columns`.

All indexes are 0-based. Display row numbers (2, 3, …) are derived from
`Table.origin`.

---

## Comparing files

The short path: two files in, every sheet diffed and traced, each sheet
ready for a grid. See [Getting started](getting-started.md).

### diffFiles

```ts signature
diffFiles(oldInput: SpreadsheetInput, newInput: SpreadsheetInput, options?: DiffFilesOptions): Promise<Comparison>
```

Read both inputs (fetching URLs, reading Blobs), parse them by content,
diff every sheet and, when either has formulas, trace downstream changes.
With `worker`, the parsing, diffing and tracing run in a Web Worker.

```ts
import { diffFiles, DiffGrid } from 'spreadiff';
declare const container: HTMLElement;

const cmp = await diffFiles(new URL('/v1.xlsx', location.href), new URL('/v2.xlsx', location.href));
new DiffGrid().render({ sheet: cmp.sheets[cmp.firstChanged], container });
```

### compareWorkbooks

```ts signature
compareWorkbooks(a: Workbook | Table, b: Workbook | Table, options?: CompareOptions): Comparison
```

The same, synchronously, for workbooks or tables you already have.

### SpreadsheetInput

`Blob | URL | ArrayBuffer | Uint8Array | { name, data } | Workbook | Table`.
`File` is a `Blob`. `{ name, data }` takes bytes, or text for CSV. Bytes are
parsed by content, so a missing or wrong extension still works.

### DiffFilesOptions

| Field | Type | Default | |
| --- | --- | --- | --- |
| `diff` | `DiffOptions` | | How to match rows and compare cells |
| `parse` | `ParseFileOptions` | | `header`, `delimiter`; ignored for parsed inputs |
| `trace` | `boolean \| 'auto' \| TraceOptions` | `'auto'` | Trace downstream changes; `'auto'` traces when either side has formulas |
| `worker` | `DiffWorker \| (() => Worker)` | none | Run in a worker: a client from `createDiffWorker`, or a function that creates the `Worker` on first use |
| `signal` | `AbortSignal` | | Abort fetching URL inputs |

`compareWorkbooks` takes the first three (its `CompareOptions`).

### Comparison

| Field | Type | |
| --- | --- | --- |
| `names` | `{ old, new }` | File names |
| `workbook` | `WorkbookDiff` | The workbook diff |
| `trace` | `WorkbookTrace \| null` | Null when not traced |
| `sheets` | `SheetView[]` | One per sheet, in `workbook.sheets` order |
| `firstChanged` | `number` | The first sheet with a change (0 if none) |
| `report()` | `ChangeReport` | Every change with its causes and contributions, built on first call (the same report as `spreadiff-lite`'s [`compare`](#compare)) |
| `timings` | `{ parse, diff, trace }` | Milliseconds |

### SheetView

A `SheetDiff` plus `index` (its position in `sheets`) and `trace` (a
`TraceContext` bound to this sheet, or undefined). Pass it to
`DiffGrid.render({ sheet, container })`: the grid finds the sheet's trace
by the diff itself, so a sheet can't be shown with another sheet's trace.

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

From `spreadiff`, `spreadiff/core`, `spreadiff/html`.

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
From `spreadiff`, `spreadiff/core`, `spreadiff/html`.

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
| `relativeTolerance` | `number` | `0` | Numbers within this fraction of the larger are equal: `0.001` ignores moves under 0.1%. A materiality threshold: such changes aren't changes anywhere (stats, grid, tracing) |

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
| `objects` | `ObjectChange[]?` | Changes to defined names, charts and pivot tables; absent when none (see [`diffWorkbookObjects`](#diffworkbookobjects)) |

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
| `labelStyles` | `{ r, c, indent?, bold? }[]?` | Indented or bold text cells, 0-based sheet coordinates. The XLSX reader fills it; row labels use it to nest sections |
| `merges` | `{ r0, c0, r1, c1 }[]?` | Merged ranges, 0-based inclusive sheet coordinates. The value is in the top-left cell. `DiffGrid` draws an unchanged horizontal merge as one cell |
| `numberFormats` | `{ codes, rows, date1904? }?` | Excel number formats: `rows[r][c]` (a `Uint16Array` per row, parallel to `rows`) indexes `codes`; 0 is General. See [`displayValue`](#displayvalue) |
| `hiddenRows` | `{ rows, filtered }?` | Hidden sheet rows (0-based, ascending); `filtered` are those an active filter hid. `SUBTOTAL` skips them as Excel does |
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
| `charts` | `ChartInfo[]?` | Charts (XLSX reader only) |
| `pivots` | `PivotInfo[]?` | Pivot tables (XLSX reader only) |

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

Parse a file by its content:
- a zip: `.xlsx`/`.xlsm`, `.xlsb` (it holds `xl/workbook.bin`) or `.ods` (its `mimetype` says so);
- a compound file: `.xls` (Excel 97-2003). A password-protected workbook of
  any kind is one too, and throws a "password-protected" error;
- otherwise text: CSV (`.tsv` names are tab-delimited), decoded as UTF-8, or
  UTF-16 when it has a byte-order mark.

Throws for files named `.xlsx`, `.xlsb`, `.xls` or `.ods` that are none of these.

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

### parseXlsb

```ts signature
parseXlsb(data: ArrayBuffer | Uint8Array, options?: ParseXlsbOptions): Workbook
```

Read an Excel binary workbook (`.xlsb`): what `parseXlsx` reads, from the
binary parts. Formulas are decoded from their tokens, shared and array
formulas included. Not read: structured references to tables (those cells
keep their values, without a formula), charts and pivot tables.

### ParseXlsbOptions

`{ name?: string; header?: boolean | 'auto' }`, as for `parseXlsx`.

### parseXls

```ts signature
parseXls(data: ArrayBuffer | Uint8Array, options?: ParseXlsOptions): Workbook
```

Read an Excel 97-2003 workbook (`.xls`, BIFF8): values, formulas (decoded
from their tokens: shared, array and data-table formulas, 3-D and
external references), dates, number formats, merged cells, hidden rows and
defined names. Excel 5 and older files throw.

### ParseXlsOptions

`{ name?: string; header?: boolean | 'auto' }`, as for `parseXlsx`.

### parseOds

```ts signature
parseOds(data: ArrayBuffer | Uint8Array, options?: ParseOdsOptions): Workbook
```

Read an OpenDocument spreadsheet (`.ods`): values, formulas (converted from
OpenFormula to Excel syntax), merged cells, hidden and filtered rows and
named ranges. Number formats aren't read.

### ParseOdsOptions

`{ name?: string; header?: boolean | 'auto' }`, as for `parseXlsx`.

How closely the readers agree with LibreOffice reading the same files is
measured in [Reading other formats](formats.md).

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
| `contributions` | `'but-for'` | How a change is split among its causes: `'but-for'` (one at a time; needn't add up) or `'shapley'` (adds up; one recalculation per subset of the edits reaching a result, up to 10 edits, within another `evaluationBudget`). See [Contributions that add up](tracing.md#contributions-that-add-up) |
| `liveFunctions` | none | More data-feed functions, by provider, on top of [`LIVE_FUNCTIONS`](#live_functions): `{ 'In-house': ['PRICE_OF'] }`. A changed cell calling one is an outside-data input |

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
| `contributions` | `(number \| null)[]?` | Parallel to `causes`, when `method` is `'recalculated'`: how much of the change each accounts for, as `attribution` says |
| `attribution` | `'but-for' \| 'shapley'?` | With `contributions`: `'but-for'`, new value − value without that cause (parts needn't add up); `'shapley'`, its average effect over every order of the edits (parts add up). See [Contributions that add up](tracing.md#contributions-that-add-up) |
| `method` | `'recalculated' \| 'references'?` | How `causes` was found: by undoing each edit and recalculating (exact), or by following references (may over-include) |
| `stale` | `{ cached, recalculated }?` | The new file's saved value is out of date: `cached` is what the file says, `recalculated` what its formulas give |
| `external` | `ExternalSource?` | On `input` cells whose formula reads outside data (Bloomberg `BDP`, `RTD`, another workbook): where from. Their saved value is taken as given, never recalculated |

### CellRef

`{ sheet, row, col }`: sheet index into `WorkbookDiff.sheets`, aligned row,
aligned column index.

### ChangeKind

`'input' | 'formula' | 'computed' | 'added' | 'removed'`. See [Four kinds
of change](tracing.md#four-kinds-of-change).

### TraceStats

`{ inputs, formulas, computed, unexplained, stale, recalculated, unmodelled, external }`:
counts of changed cells by kind; `unexplained` counts computed cells with
no traced cause, `stale` those whose saved value was out of date,
`recalculated` those whose causes came from recalculation, `unmodelled`
formulas near a change the engine couldn't reproduce, and `external`
inputs whose value came from outside data (also counted in `inputs`).

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

### externalSource

```ts signature
externalSource(formula: string, extra?: Record<string, string[]>): ExternalSource | null
```

Where a formula's value comes from, if from outside the workbook: a
data-feed function from [`LIVE_FUNCTIONS`](#live_functions) or `extra`
(with or without Excel's `_xll.` prefix), any other `_xll.` add-in
function, or a link to another workbook (`[1]Sheet1!A1`). Text inside
strings is ignored. `null` for a formula that reads only the workbook.
See [spreadiff-lite → Outside data](spreadiff-lite.md#outside-data).

### ExternalSource

`{ provider, function }`: `provider` is `'Bloomberg'`, `'LSEG'`,
`'FactSet'`, `'S&P Capital IQ'`, `'Morningstar'`, `'Excel'`, `'add-in'`
(another XLL function), `'link'` (another workbook) or a provider from
`liveFunctions`; `function` is the function's name (`'BDP'`) or the link
(`'[1]'`).

### LIVE_FUNCTIONS

`Record<provider, functionNames[]>`: the data-feed functions recognised by
default. A name ending in `*` matches every function starting with it
(`CIQ*`). The list is in [spreadiff-lite → Outside
data](spreadiff-lite.md#outside-data).

---

## Labels

### tableLabeler

```ts signature
tableLabeler(table: Table, options?: LabelOptions): (row: number, col: number) => CellLabel
```

What the sheet calls a cell, from the layout around it: section headings
above the row, the row's label cells and the column headings above the
cell. Takes table coordinates (`row` into `table.rows`). The layout is
analysed once, on the first call. The rules and their failure modes are in
[spreadiff-lite → How cells are named](spreadiff-lite.md#how-cells-are-named).

### LabelOptions

| Field | Default | |
| --- | --- | --- |
| `keyColumns` | `[]` | Key columns by name (`TableDiff.keyColumns`): their values lead each row's label, and no sections are inferred (the table is records, not a model) |
| `maxRowCells` | `2` | Label cells per row |

### CellLabel

`{ sections, row, column }`, each a `string[]`: section headings
outermost first (`['India']`), the row's label cells (`['Revenue']`) and
the column headings outermost first (`['Actual', 'Q2']`, or the column
name). Any can be empty.

### sheetLabels

```ts signature
sheetLabels(table: Table, options?: LabelOptions): SheetLabels
```

Every label of a table at once. The layout is analysed when this is called
(one pass, about 50 ms per 100,000 rows); lookups are cheap after.
`DiffGrid` and `renderHTML` use it for their headers.

### SheetLabels

| Member | |
| --- | --- |
| `cell(row, col): CellLabel` | A cell's labels, as `tableLabeler` gives them |
| `row(row): RowLabel` | A row's sections and label, and whether it is a section heading |
| `column(row, col): string[]` | The column headings in force at a row: the nearest heading rows above it, or below for rows above the first, or the header |
| `headingRow(row): boolean` | The row holds column headings (years, dates or titles) and no label of its own |
| `labelColumns(): number[]` | Columns whose every non-blank cell is part of its row's label: the row labels show them |

Table coordinates throughout.

### RowLabel

`{ sections, row, heading }`: section headings outermost first, the row's
label cells, and whether the row is itself a heading.

### formatLabel

```ts signature
formatLabel(label: CellLabel): string
```

One line: sections and row joined with ` › `, then ` · ` and the column
headings: `'India › Revenue · 2025'`. Empty when there are no labels.

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

## Number formats

Showing values as the sheet shows them. The format language is Excel's,
in the en-US locale.

### displayValue

```ts signature
displayValue(v: CellValue | undefined, code: string | undefined, date1904?: boolean): string
```

A value as a cell with this format code shows it: `0.123` with `0.0%` is
`12.3%`, `"2024-01-01"` (dates are ISO text) with `d-mmm-yy` is
`1-Jan-24`. Falls back to `formatCell` with no code, or one it can't apply.

### displayCell

```ts signature
displayCell(table: Table, row: number, col: number): string
```

`displayValue` with the cell's own code from `table.numberFormats`.

### formatCodeAt

```ts signature
formatCodeAt(table: Table, row: number, col: number): string | undefined
```

A cell's format code (`'#,##0.00'`), or undefined for General or none.

### displayMagnitude

```ts signature
displayMagnitude(abs: number, code: string | undefined): string
```

The size of a change in the cell's format (`$1,200` for `$#,##0`),
unsigned. Percentages change in points (`1.2 pp`). Dates, times and changes
too small for the format to show stay plain numbers.

### formatNumberCode

```ts signature
formatNumberCode(v: number, code: string, date1904?: boolean, general?: (n: number) => string): string | null
```

Format a number with a format code: sections and conditions
(`[>=1000]`), colors, literals, `_` and `*` spacing, thousands and scaling
commas, `%`, scientific and engineering notation, fractions (`# ?/?`), dates
and times (including elapsed `[h]:mm`), and the 1900 and 1904 date systems.
Null for a code it doesn't understand, so you can fall back. The engine's
`TEXT` uses it.

### formatTextCode

```ts signature
formatTextCode(v: string, code: string): string | null
```

Format text with a code's text section (the fourth, or a lone one with
`@`); null when none applies.

### isDateFormat

```ts signature
isDateFormat(code: string): boolean
```

Whether a code shows numbers as dates or times (by its first section).

### BUILTIN_FORMATS

The format codes Excel stores by number instead of text (`14` →
`m/d/yyyy`), including the East Asian date formats (27–36, 50–58, shown
in their Japanese forms).

---

## Workbook objects

### diffWorkbookObjects

```ts signature
diffWorkbookObjects(a: Workbook, b: Workbook): ObjectChange[]
```

Changes to defined names, charts and pivot tables. `diffWorkbooks` puts
them in `WorkbookDiff.objects`. Charts pair by sheet and title (then
position), pivots by sheet and name. References are compared as written:
a chart whose range moved because rows were inserted shows as changed.
The XLSX reader fills `Workbook.charts` and `Workbook.pivots`; the other
readers don't read charts or pivot tables.

### ObjectChange

`{ kind: 'name' | 'chart' | 'pivot'; change: 'added' | 'removed' | 'modified'; name: string; details: string[] }`:
`details` says what changed, one line each (a name's `Inputs!$B$2 → Inputs!$B$3`, a pivot's `source: Data!$A$1:$D$90 → Data!$A$1:$D$120`).

### ChartInfo

`{ sheet, title, types, series: { name?, categories?, values? }[] }`: a
chart as read from the file; `types` are its plot types (`bar`, `line`).

### PivotInfo

`{ sheet, name, location, source, values }`: a pivot table, where it sits,
the range it summarizes and its value fields.

---

## Review

Accepting and rejecting changes. A decision applies to an edit (a value or
formula someone changed, or data from outside), an added or removed row,
or an added, removed or renamed column. Downstream results aren't decided:
they follow their inputs. See [Reviewing changes](review.md).

### Decision

`'accept' | 'reject'`

### SheetDecisions

`{ cells, rows, columns }`: `Map`s from a cell key (`row *
diff.columns.length + col`), an aligned row, or a column index to a
`Decision`. Anything not listed is accepted.

### noDecisions

```ts signature
noDecisions(): SheetDecisions
```

Empty decisions: everything accepted.

### isDecidable

```ts signature
isDecidable(diff: TableDiff, row: number, col: number): boolean
```

Whether a changed cell is a decision to make (an edit), rather than a result.

### decideRange

```ts signature
decideRange(diff: TableDiff, decisions: SheetDecisions, rows: number[], columns: number[], decision: Decision | undefined): SheetDecisions
```

Decide every change in a block (a `GridRange`'s `rows` and `columns`):
its edits and its added or removed rows. Returns new decisions;
`undefined` clears them.

### decisionOf

```ts signature
decisionOf(diff: TableDiff, decisions: SheetDecisions | undefined, row: number, col: number): Decision
```

How a cell was decided, through its row or column decision if it has one.

### decisionStats

```ts signature
decisionStats(diff: TableDiff, decisions: SheetDecisions): { accepted: number; rejected: number; changes: number }
```

### applyDecisions

```ts signature
applyDecisions(diff: TableDiff, decisions: SheetDecisions): Table
```

The new table with rejected changes undone: old values and formulas back,
rejected added rows dropped, rejected removed rows restored where they were.

### tableToCsv

```ts signature
tableToCsv(table: Table, delimiter?: string): string
```

RFC 4180 CSV (formulas as their values); `'\t'` writes TSV.

### applyDecisionsToXlsx

```ts signature
applyDecisionsToXlsx(newFile: ArrayBuffer | Uint8Array, diff: WorkbookDiff, decisions: (SheetDecisions | undefined)[], options?: { recalculate?: boolean }): { file: Uint8Array; applied: number; recalculated: number; skipped: SkippedDecision[] }
```

The new .xlsx with rejected cell changes written back in place, keeping
styles. Old formulas are moved to the new file's rows and columns.
Formulas that depend on a reverted cell get fresh cached values, and the
file is marked for recalculation on open. Row and column decisions, and
cells in array formulas or table headers, aren't written and come back in
`skipped`. `.xlsx`/`.xlsm` only. From `spreadiff`, `spreadiff/core`, `spreadiff/xlsx`.

### SkippedDecision

`{ sheet, target, reason }`: a decision `applyDecisionsToXlsx` couldn't write.

### translateFormula

```ts signature
translateFormula(formula: string, sheet: string, diff: WorkbookDiff): string | null
```

An old formula with its references moved to the new file's rows and
columns (`=B4*2` → `=B5*2` after a row went in above). Null when one
points at a cell the new file doesn't have.

---

## Merge

### mergeTables

```ts signature
mergeTables(base: Table, ours: Table, theirs: Table, options?: MergeOptions): MergeResult
```

Three-way merge: changes made on one side are taken, changes both made
alike are taken once, and changes that differ are conflicts, resolved
toward `prefer` and listed. Rows are matched by key, as in `diffTables`.

```ts
import { mergeTables, parseCsv } from 'spreadiff';

const base = parseCsv('id,price,qty\n1,10,5');
const ours = parseCsv('id,price,qty\n1,12,5');
const theirs = parseCsv('id,price,qty\n1,10,6');
mergeTables(base, ours, theirs).table.rows; // [[1, 12, 6]]
```

### MergeOptions

`DiffOptions` plus `prefer: 'ours' | 'theirs'` (default `'ours'`): the
side a conflict takes.

### MergeResult

| Field | Type | |
| --- | --- | --- |
| `table` | `Table` | The merged table |
| `conflicts` | `MergeConflict[]` | Each conflict and the side taken |
| `stats` | `{ ours, theirs, both }` | Changes taken without conflict: from ours, from theirs, and made alike on both |
| `diff` | `TableDiff` | The merged table against the base, to show in a grid |

### MergeConflict

| Field | Type | |
| --- | --- | --- |
| `kind` | `'cell' \| 'delete-edit' \| 'both-added'` | Both changed a cell differently; one deleted a row the other edited; both added a row with the same key |
| `row` | `number` | Row in the merged table, or -1 for a deleted row left out |
| `baseRow` | `number?` | Row in the base |
| `column` | `string?` | For `cell` |
| `base`, `ours`, `theirs` | `CellValue?` | For `cell` |
| `oursRow`, `theirsRow` | `CellValue[] \| null?` | For row conflicts; null where that side deleted it |
| `took` | `'ours' \| 'theirs'` | |

### conflictAnnotations

```ts signature
conflictAnnotations(result: MergeResult): { row: number; column?: string; data: MergeConflict; id: number }[]
```

The conflicts as `DiffGrid` annotations for `result.diff`.

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
| `render({ sheet, container })` | Show a [`SheetView`](#sheetview) from `diffFiles` / `compareWorkbooks`, with its trace |
| `setDiff(diff)` | Show another diff |
| `setOptions(partial)` | Change options (cheap when nothing layout-related changed) |
| `getOptions()` | Current options |
| `nextChange()` / `prevChange()` | Select the next / previous block of changes; `false` at the end |
| `scrollToRow(row, align?)` | Scroll an aligned row into view (`'start' \| 'center' \| 'nearest'`) |
| `select(row, col?)` / `select(null)` | Select a cell; `false` if out of range |
| `selectRange(row, col, focusRow, focusCol)` | Select a block of cells; `false` if out of range |
| `getRange()` | The selected [`GridRange`](#gridrange), or `null` |
| `expandAll()` / `collapseAll()` | Unfold / refold unchanged regions |
| `cellInfo(row, col)` | A cell's [`CellInfo`](#cellinfo); throws `RangeError` if it doesn't exist |
| `refreshTheme()` | Re-read CSS variables |
| `drawNow()` | Paint synchronously (tests, benchmarks) |
| `cleanUp()` | Unmount and release observers |
| `host` | The element inserted into your container |

### DiffGridOptions

All optional; see the [options table](grid.md#options) for defaults and
behavior: `layout`, `cellDisplay`, `collapseUnchanged`, `context`,
`hideUnchangedColumns`, `frozenColumns`, `columnWidths`, `rowHeaders`,
`columnHeaders`, `rowHeaderWidth`, `foldLabels`, `rowNumberOffset`,
`theme`, `minimap`, `tooltips`, `highlightComputed`, `trace`, `traceFocus`,
`annotations`, `renderAnnotation`, `numberFormats`, `decisions`,
`onCellClick`, `onCellHover`, `onSelectionChange`, `onRangeChange`,
`accessibility`, `ariaLabel`, `strings`, `locale`.

### DiffStrings

Every word the grid and `SpreadsheetDiff` show or announce: the
screen-reader text, the popover, the separators, the sheet tabs, the
outputs panel and the review bar. Pass a partial set as the `strings`
option; the rest stay English. Messages with a number get it twice:
formatted for `locale` ("1.234" in `de-DE`), and as a number, to pick a
plural form. See [Translating the grid](grid.md#translating-the-grid).

### ENGLISH_STRINGS

The English [`DiffStrings`](#diffstrings): the defaults, and a starting
point for a translation.

### DiffGridRenderProps

`{ diff: TableDiff; container: HTMLElement }` or `{ sheet: SheetView;
container: HTMLElement }`: the argument to `render`.

### GridRange

| Field | Type | |
| --- | --- | --- |
| `anchor` | `{ row, col }` | Where the selection started (the active cell) |
| `focus` | `{ row, col }` | The opposite corner; equal to `anchor` for one cell |
| `rows` | `number[]` | Aligned rows in the range, top to bottom; rows folded away aren't in it |
| `columns` | `number[]` | Column indexes in the range, left to right; hidden columns aren't in it |
| `a1` | `{ old?, new? }` | The range in A1 notation on each side (`"B2:D5"`), when its cells are one block there |

From a drag, shift-click or shift + arrow keys, `selectRange()`, and
`onRangeChange`. Pass `rows` and `columns` to [`decideRange`](#deciderange).

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

### SpreadsheetDiff

```ts signature
new SpreadsheetDiff<A = unknown>(options?: SpreadsheetDiffOptions<A>)
```

The whole review in one component: give it two files and it compares them
(in a worker if you pass one), then shows sheet tabs, the grid with
downstream changes traced, an outputs panel (each changed result and a
waterfall of what drove it) and changes to names, charts and pivot tables.
In review mode it accepts and rejects changes and downloads the result.

```ts
import { SpreadsheetDiff } from 'spreadiff';
declare const container: HTMLElement;

const view = new SpreadsheetDiff({ layout: 'split', review: true });
await view.render({ oldFile: new URL('/v1.xlsx', location.href), newFile: new URL('/v2.xlsx', location.href), container });
```

URL inputs are fetched for you; `File`s and bytes are read directly. The component
renders inside a shadow root, so page CSS doesn't reach its internals:
theme it with the same CSS variables as the grid ([Theming](theming.md)),
and in tests query through `view.host.shadowRoot` (Playwright's
`getByText` and `getByRole` see into it).

| Member | |
| --- | --- |
| `render(props)` | Mount and show two files, or a comparison; resolves when shown, rejects if reading fails (the error is shown too) |
| `showSheet(index)` | Show a sheet |
| `reveal(ref)` | Switch to a `CellRef`'s sheet, select it and scroll it into view |
| `getDecisions()` / `setDecisions(d)` | Review decisions, one `SheetDecisions` per sheet |
| `decideSelection(decision)` | Accept or reject every change in the selected range (`undefined` clears) |
| `exportXlsx()` | The new .xlsx with decisions applied (see `applyDecisionsToXlsx`); needs the new file's bytes |
| `exportCsv(sheet?)` / `mergedTable(sheet?)` | A sheet with decisions applied |
| `setOptions(partial)` | Change options; grid options go to the grid |
| `comparison`, `sheet` | The comparison shown, and the index of the sheet shown |
| `grid` | The `DiffGrid` inside |
| `host` | The element inserted into your container |
| `cleanUp()` | Unmount and release everything |

### SpreadsheetDiffOptions

Every `DiffGridOptions` field except `trace` (the component supplies it),
plus:

| Field | Type | Default | |
| --- | --- | --- | --- |
| `compare` | `CompareOptions` | | `diff`, `parse`, `trace` (see [`DiffFilesOptions`](#difffilesoptions)) |
| `worker` | `DiffWorker \| (() => Worker)` | none | Compare in a worker |
| `outputs` | `boolean \| 'auto'` | `'auto'` | The outputs panel; `'auto'` shows it when tracing found changed results |
| `review` | `boolean` | `false` | Accept/reject buttons (and `a` / `r` / `u` keys in the grid) and a download button: `.xlsx` for an .xlsx/.xlsm new file, else the shown sheet as `.csv` |
| `renderSheetHeader` | `(sheet, comparison) => HTMLElement \| string \| null` | | Content above each sheet's grid |
| `onLoad` | `(comparison) => void` | | The files were compared |
| `onError` | `(error) => void` | | Reading or comparing failed |
| `onSheetChange` | `(sheet) => void` | | Another sheet is shown |
| `onDecisionsChange` | `(decisions) => void` | | Review decisions changed |
| `decisions` | `SheetDecisions[]` | | Decisions, one per sheet, applied when files are compared and whenever a different array is passed; see [Keeping decisions](review.md#keeping-decisions) |

### SpreadsheetDiffRenderProps

`{ oldFile: SpreadsheetInput; newFile: SpreadsheetInput; container }`, or
`{ comparison: Comparison; newFile?: { name, data }; container }` for a
comparison you made (`newFile` enables the .xlsx download).

---

## React

### DiffGrid (React)

```ts signature
<DiffGrid diff={…} {...DiffGridOptions} className? style? ref? />
```

From `spreadiff/react`. Renders a container `<div>` with the canvas grid
inside; option props update the grid in place. `renderAnnotation` may
return JSX (portaled into place).

### SpreadsheetDiff (React)

```ts signature
<SpreadsheetDiff oldFile={…} newFile={…} {...SpreadsheetDiffOptions} className? style? ref? />
```

From `spreadiff/react`. The `SpreadsheetDiff` component; pass `comparison`
instead of the files if you have one. New files are compared again; option
props update in place. `renderSheetHeader` may return JSX.

```tsx
import { SpreadsheetDiff } from 'spreadiff/react';

export const Review = ({ before, after }: { before: File; after: File }) => (
  <SpreadsheetDiff oldFile={before} newFile={after} review style={{ height: 600 }} />
);
```

### SpreadsheetDiffProps

`SpreadsheetDiffOptions` plus `oldFile`, `newFile`, `comparison`,
`className`, `style`, and `renderSheetHeader: (sheet, comparison) => ReactNode`.

### SpreadsheetDiffHandle

The `ref` value: `showSheet`, `reveal`, `getDecisions`, `setDecisions`,
`decideSelection`, `exportXlsx`, `exportCsv`, `mergedTable`, and
`instance` (the underlying `SpreadsheetDiff`, or `null` before mount).

### DiffGridBaseProps

`DiffGridOptions` plus `className`, `style`, and a React-flavored
`renderAnnotation: (a) => ReactNode`: the props other than how the diff is
given.

### DiffGridProps

`DiffGridBaseProps` plus either `diff` or `sheet` (a `SheetView`).

### DiffGridHandle

The `ref` value: `nextChange`, `prevChange`, `scrollToRow`, `expandAll`,
`collapseAll`, `select`, `refreshTheme`, `cellInfo`, and `instance` (the
underlying `DiffGrid`, or `null` before mount).

---

## HTML and text rendering

From `spreadiff/html`. See [spreadiff/html](html.md) for a guide.

### renderHTML

```ts signature
renderHTML(diff: TableDiff, options?: HtmlOptions): string
```

The diff as an HTML string (changes plus context). Cell content is escaped.
Style it with [`HTML_CSS`](#html_css).

### renderText

```ts signature
renderText(diff: TableDiff, options?: { context?; changedColumnsOnly?; maxRows?; maxWidth? }): string
```

An aligned plain-text diff with ASCII `+` / `-` line markers, for terminals,
logs and ```` ```diff ```` blocks. `maxWidth` (default 24) caps column width.

### HtmlOptions

| Field | Default | |
| --- | --- | --- |
| `layout` | `'split'` | `'split'` or `'unified'` |
| `display` | `'stacked'` | How changed cells read |
| `context` | `2` | Unchanged rows around each change |
| `changedColumnsOnly` | `false` | Only changed columns, plus keys (or the row-label column when rows were matched by order) |
| `maxRows` | `2000` | Rows of output before truncating |
| `kindOf` | — | `(row, col) => ChangeKind \| undefined` override |
| `rowNumberOffset` | from `origin` | Displayed number of row index 0 |
| `rowHeaders` | `'labels'` | `'labels'` (inferred row labels, or the number when a row has none), `'numbers'` or `'both'` |
| `columnHeaders` | `'labels'` | `'labels'` (headings), `'letters'` or `'both'` |
| `foldLabels` | `true` | With labels, leave out unchanged label columns and heading rows |

### HtmlLayout

`'split' | 'unified'`

### HtmlDisplay

`'stacked' | 'new' | 'delta' | 'inline'`

### HTML_CSS

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
triples) and `contributions` (NaN for null), plus `stale` and `external`
values by index. Structured-cloning one object per traced cell cost about 200 ms per
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

---

## spreadiff-lite

A separate package (`import … from 'spreadiff-lite'`), with no UI. It also
re-exports `externalSource`, `LIVE_FUNCTIONS`, `tableLabeler`,
`formatLabel`, `CellLabel` and `ExternalSource` from `spreadiff/core`. Guide:
[spreadiff-lite](spreadiff-lite.md).

### compare

```ts signature
compare(oldFile: FileInput, newFile: FileInput, options?: CompareOptions): ChangeReport
```

Parse (if needed), diff and trace two workbooks, and build the report:
edits, downstream results with their causes and contributions, shape
changes and notes. Formulas reading outside data are taken as saved.

### summarize

```ts signature
summarize(oldFile: FileInput, newFile: FileInput, options?: CompareOptions & MarkdownOptions): string
```

`toMarkdown(compare(oldFile, newFile, options), options)`.

### reportOf

```ts signature
reportOf(diff: WorkbookDiff, trace: WorkbookTrace, oldName?: string, newName?: string): ChangeReport
```

The report for a diff and trace you already have, for example from a
`spreadiff` worker.

### toMarkdown

```ts signature
toMarkdown(report: ChangeReport, options?: MarkdownOptions): string
```

The report as Markdown. In order:
- a summary line;
- tables of edits, final outputs and intermediate results (or one
  *Downstream* table when every result is an output);
- rows, columns and sheets that changed;
- notes.

Empty sections are left out. Identical files give `No changes.`

### FileInput

`Workbook | { name: string; data: ArrayBuffer | Uint8Array }`: a parsed
workbook, or a file's name (which picks the parser: `.xlsx`, `.xlsm`,
`.csv`, `.tsv`) and bytes.

### CompareOptions

| Field | |
| --- | --- |
| `diff` | [`DiffOptions`](#diffoptions): `keys`, `numericTolerance`, … |
| `parse` | [`ParseFileOptions`](#parsefileoptions): `header`, `delimiter`. Ignored for workbooks passed in parsed |
| `trace` | [`TraceOptions`](#traceoptions): `recalculate`, `liveFunctions`, budgets |
| `sheets` | Only these sheets, by name |

### MarkdownOptions

| Field | Default | |
| --- | --- | --- |
| `maxRows` | `50` | Rows per table; the rest are counted under it. `Infinity` for all |
| `maxCauses` | `3` | Causes listed per downstream result |
| `headingLevel` | `2` | Level of the title; sections are one deeper |
| `digits` | `15` | Significant digits for values and changes (15 is exact). Contributions show at most 4 |
| `intermediate` | `true` | List intermediate results as well as final outputs |

### ChangeReport

| Field | Type | |
| --- | --- | --- |
| `old`, `new` | `string` | File names |
| `totals` | `{ inputs, formulas, external, downstream, outputs, recalculated, unexplained, stale }` | Counts |
| `edits` | `CellChange[]` | Input, formula and outside-data changes, in sheet order |
| `downstream` | `CellChange[]` | Results that moved: final outputs first, then sheet order |
| `structure` | `StructureChange[]` | Rows, columns and sheets added, removed or moved |
| `notes` | `string[]` | Caveats: outside data, stale files, results traced by references, truncation |
| `diff`, `trace` | `WorkbookDiff`, `WorkbookTrace` | The underlying diff and trace. Not enumerable, so `JSON.stringify` leaves them out |

### CellChange

| Field | Type | |
| --- | --- | --- |
| `id` | `string` | `Sheet!A1` (new file; old file for removed cells) |
| `sheet`, `cell` | `string` | Sheet name and A1 address |
| `label`, `labels` | `string`, `CellLabel` | What the sheet calls it ([How cells are named](spreadiff-lite.md#how-cells-are-named)) |
| `kind` | `'input' \| 'formula' \| 'external' \| 'computed'` | Typed over, formula edited, outside data refreshed, or moved downstream |
| `before`, `after` | `CellValue` | Values (the recalculated one when the new file's was stale) |
| `formula` | `{ before, after }?` | Formula text, when either side has one |
| `delta`, `ratio` | `number?` | `after − before` and `delta / \|before\|`, for numbers |
| `external` | `ExternalSource?` | For `external` edits |
| `impact` | `number?` | For edits: downstream results it moved |
| `causes` | `Cause[]?` | For `computed`: the edits that moved it, largest contribution first |
| `causeCount` | `number?` | Exact number of causes |
| `method` | `'recalculated' \| 'references'?` | How the causes were found |
| `attribution` | `'but-for' \| 'shapley'?` | What the causes' contributions measure (see `CellTrace.attribution`) |
| `output` | `boolean?` | No other changed cell reads it: a final output |
| `unexplained` | `'volatile' \| 'dynamic' \| 'unknown'?` | Why no cause was found |
| `stale` | `{ saved }?` | The new file saved this value without recalculating |

### Cause

`{ id, label, contribution }`: the edit's `Sheet!A1`, its label, and the
result now minus the result without that edit (`null` when traced by
references or not numeric).

### StructureChange

`{ sheet, kind, rows?, labels?, column?, impact? }`. `kind` is one of:
- `'sheet-added'`, `'sheet-removed'`;
- `'rows-added'`, `'rows-removed'`, `'rows-moved'`;
- `'column-added'`, `'column-removed'`, `'column-renamed'`.

The other fields:
- `rows`: spreadsheet row numbers;
- `labels`: their row labels;
- `column`: the column name (`'old → new'` for a rename);
- `impact`: results the added or removed cells moved.

