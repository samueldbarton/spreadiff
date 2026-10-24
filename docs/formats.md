# File formats

`parseFile` (and `diffFiles`, `SpreadsheetDiff`, the CLI and the action)
reads these, choosing the reader from the file's contents, so a file with
the wrong extension, or none, still reads:

| Format | Reader | Values and formulas | Also read | Not read |
| --- | --- | --- | --- | --- |
| `.xlsx`, `.xlsm` | `parseXlsx` | Cached values; formula text, shared and array formulas expanded | Number formats, dates, names, Excel tables, merges, hidden and filtered rows, charts and pivot tables (for [object changes](api.md#diffworkbookobjects)) | Styles beyond number formats, comments, macros |
| `.xlsb` | `parseXlsb` | Cached values; formulas decoded from their tokens | Number formats, dates, names, merges, hidden and filtered rows | Structured references (`Sales[Amount]`: the cell keeps its value, without a formula), charts, pivot tables |
| `.xls` (Excel 97-2003) | `parseXls` | Cached values; formulas decoded from their tokens, including shared, array and data-table formulas | Number formats, dates, names, merges, hidden rows | Excel 5 and older files, charts, pivot tables |
| `.ods` | `parseOds` | Values; formulas converted from OpenFormula to Excel syntax | Dates, names, merges, hidden and filtered rows | Number formats, charts |
| `.csv`, `.tsv` | `parseCsv` | Text, with numbers converted when lossless | | |

A password-protected workbook throws an error saying so: its contents are
encrypted. Save a copy without the password.

Only the .xlsx writer exists: reviewing an .xlsb, .xls or .ods file works,
but the result downloads as CSV (see [Reviewing changes](review.md)).

## How well the readers agree with LibreOffice

Each reader was compared cell by cell with LibreOffice's reading of the
same file: LibreOffice converts the file to .xlsx, and both are read. A
cell agrees when its value matches (numbers to 1 part in 10⁹) and its
formula text matches after normalizing spelling differences (`_xlfn.`
prefixes, quoting of sheet names, `A1:A1` for `A1`).

| Reader | Files | Values agree | Formulas agree |
| --- | --- | --- | --- |
| `.xls` | 119 of 120 Apache POI test files (one LibreOffice can't open) | 99.84% of 71,627 cells | 98.97% of 4,253 |
| `.xlsb` | 10 of 11 POI test files (one is encrypted) | 464 of 467 cells | 281 of 284 |
| `.ods` | 331 of 353 POI .xlsx files, converted to .ods by LibreOffice (it can't open the rest) | 99.91% of 1,105,498 cells | 85.02% of 27,317; 99.69% leaving out structured references, which ODS can't hold |

Every disagreement was looked at. Checking the .ods figure found two
reader bugs, now fixed: long runs of repeated cells were cut at 10,000
(one file lost 6,384 cells) and cells after a long empty run were
misplaced. With those fixed, none of the remaining disagreements is the
reader misreading the file:

- **LibreOffice recalculates on load, spreadiff reads the stored value.**
  `TODAY()` cells, results saved stale by an older Excel, `YEAR(0)` (Excel
  1900, LibreOffice 1899) and a few functions where the two disagree. Since
  a diff should show what the file says, the stored value is the right one.
- **LibreOffice's own writing.** It truncates text to 32,767 characters
  (one file stores 32,770), stores plain booleans as `=TRUE()` formulas, and
  spells data tables and external-link numbers its own way (43 of the 44
  .xls formula differences).
- **ODS can't hold structured references** (`Sales[Amount]`), so
  LibreOffice converts them to ranges: 4,018 of the 4,091 .ods formula
  differences.
- **The conversion itself.** LibreOffice regenerates pivot tables in its
  own language and layout when it opens a file (most of the 1,000 .ods
  value differences), and drops cells from a few damaged files.

### How this can be wrong

- **The .xlsb figure rests on a small corpus**: 467 cells, mostly
  references, arithmetic and `SUM`. Functions, names, 3-D references,
  array constants and array formulas are covered by unit tests built from
  the format specification, not by files Excel wrote. Treat .xlsb formulas
  as less proven than .xls ones.
- **The oracle is another reader.** Where LibreOffice and spreadiff make
  the same mistake, it doesn't show. The .xls and .xlsb readers share the
  formula decoder, so a decoding error would show up in both.
- **The POI files are test cases**: many are small and built to exercise
  one feature or bug. They're broader than typical workbooks but less
  dense.

Rerun with `npm run bench:trace` (needs LibreOffice); it fetches the POI
files and prints these tables.

## Number formats

Cells are shown as the sheet formats them: `0.123` with `0.0%` is
`12.3%`, `1234.5` with `$#,##0.00` is `$1,234.50`, a date with `d-mmm-yy`
is `1-Jan-24`. The grid, the HTML renderer and the git textconv do
this (the Markdown report shows stored values); `numberFormats: false` on the grid shows the stored values. Changes
are shown in the cell's format too: `+$200`, and `+1.2 pp` (points) for a
percentage, so a move from 10% to 11.2% isn't read as 12%.

The formatter handles sections and conditions (`[>=1000]`), colors,
literals, `_` and `*` spacing, thousands and scaling commas, `%`,
scientific and engineering notation, fractions, dates, times and elapsed
times (`[h]:mm`), both date systems, and Excel's built-in format numbers,
including the East Asian date formats. The locale is en-US.

Against LibreOffice's "as shown" export of every formatted cell in the
test corpus, 7,590 of 7,603 cells (99.8%) display the same, over 32
distinct formats. The rest:

- **Serial 0 as a date**: Excel shows it as 1900-01-00 (so `DAY(0)` is 0),
  and spreadiff does too; LibreOffice shows 1899-12-30.
- **`-0`**: a negative number that rounds to zero, shown with or without
  its sign.

Not settled, because Excel wasn't available to check: rounding of
fractional seconds (`mm:ss.0` of 0.0789 is `53:37.0` here and `53:36.9`
in LibreOffice), and the padding of accounting formats (`_(` and `* `)
where the two agree with each other but not necessarily with Excel.

The engine's `TEXT()` uses the same formatter, and is fuzzed against
LibreOffice in CI with 35 format codes × 10 values.

Not handled: other locales' separators and month names, and `[$-…]` locale
codes beyond currency symbols (the symbol is shown; the locale's
conventions aren't applied).
