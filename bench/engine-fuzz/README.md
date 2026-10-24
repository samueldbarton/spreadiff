# Engine differential test

`npm run bench:engine` checks spreadiff's formula engine against
LibreOffice, which really recalculates:
1. `build.py` writes a workbook of tricky inputs and probe formulas. There
   are about 350 curated edge cases and, by default, 1,500 random
   expressions built from the engine's functions. LibreOffice recalculates
   it and saves the results.
2. `compare.ts` evaluates every probe with spreadiff's engine and prints
   each difference.

`npm run bench:engine -- 3000 11` runs 3,000 random probes with seed 11.

CI runs `npm run bench:engine -- 1500 7 --check bench/engine-fuzz/expected.json`
(see `.github/workflows/ci.yml`, on Ubuntu 24.04's LibreOffice 24.2). It
fails on any difference that isn't in `expected.json`, the 107 known ones
listed below by kind. When a change fixes some, the run lists them as
gone; when LibreOffice, not the engine, is wrong about a new one, record
it with `--write-baseline` (same count and seed) and add its kind to the
table below.

## Reading the differences

spreadiff follows **Excel**. LibreOffice is the oracle, but it differs
from Excel in known ways, so some differences are expected. The engine is
right in these:

| LibreOffice | Excel (and the engine) |
| --- | --- |
| Booleans are numbers in ranges (`SUM`, `COUNT`, `AVERAGE`, `MEDIAN`, `COUNTIF(">=0")` count them) | Booleans in ranges are skipped |
| `ISNUMBER(TRUE)` is TRUE; `TRUE>1` is FALSE | Booleans aren't numbers; they sort above all numbers |
| `"a"&TRUE` is `"a1"` | `"aTRUE"` |
| `IF("5", …)` reads text as a number | `#VALUE!` |
| `0^-1` is `#NUM!`; `SQRT(-1)` is `#VALUE!` | `#DIV/0!`; `#NUM!` |
| `1E+20&""` is `"1E+020"` (three exponent digits) | `"1E+20"` |
| No 29 February 1900 (`DAY(59)` is 27) | Excel keeps Lotus 1-2-3's fictional leap day (`DAY(59)` is 28) |
| `FIND("","abc")` is `#VALUE!`; `CODE("")` is 0 | 1; `#VALUE!` |
| `XLOOKUP` is unknown before LibreOffice 24.8 | Supported |
| `SUM(A1:A3*B1:B3)` without array entry is `#VALUE!` | Dynamic arrays: evaluated element-wise |
| `TEXT(-0.004,"0.00")` is `0.00` | `-0.00`: a negative rounded to zero keeps its sign |
| `TEXT(5,"#.##")` is `5` | `5.`: the point stays |
| Dates count from 1899-12-30, so `TEXT(0.5,"d-mmm-yy")` is `30-Dec-99` | `0-Jan-00`: serial 0 is 1900-01-00 |
| A date past 9999-12-31 shows as `#FMT` | Excel shows `#####`; the engine leaves it unsupported |

Also expected: `MATCH(x, ascending, -1)`. Its input breaks the function's
contract (`-1` needs descending data), and the result isn't defined.

## Found and fixed by this test

- `CEILING` / `FLOOR` rounded the wrong way with a negative number and a
  negative significance, and `.MATH` ignored its `mode`.
- Text like `"1,234"`, `"$1,234"` and `"(5)"` wasn't read as a number.
- Date and time text (`"2024-03-01" - "2024-02-01"`, `DAYS(text, text)`,
  `"2:30 PM"`) gave `#VALUE!`. Slash dates (`1/2/2024`) depend on locale,
  so they are unsupported rather than guessed.
- Numbers past 15 digits turned into text in full (`1E+20&""`).
- Text compared by code point, so `"~" > "b"`. Excel orders punctuation
  before digits before letters, and accents with their letter.

Then, checked against Excel's saved results in the real-workbook corpus
(`bench/trace-eval/engine-check.ts`):

- `TEXT` ignored the fourth (text) section and printed empty sections, and
  read `TRUE` as 1. LibreOffice reads it as 1 too; Excel formats it as the
  text `TRUE`.
- `TEXT` dropped fractional seconds and printed `AM/PM` as `A1/P1`.
- Serial 0 was 1899-12-31. Excel calls it 1900-01-00, so `DAY(0)` is 0.

## What it can't tell you

- Where LibreOffice and Excel *both* differ from the engine in the same
  way, nothing shows.
- It runs in one locale (en-US).
- Excel's general number format for very small numbers (`0.0000001&""`)
  isn't verified: LibreOffice and the engine differ, and Excel wasn't
  available to settle it.
- Nor is rounding of fractional seconds: `TEXT(0.0789,"mm:ss.0")` is
  `53:37.0` in the engine (it rounds 36.96 s) and `53:36.9` in LibreOffice
  (it truncates).

The probes include `TEXT` over 35 format codes × 10 values (finance and
accounting codes, scaling commas, conditions, fractions, scientific,
dates). The grid displays cells through the same formatter. For how cells
display in real files, `bench/trace-eval/format-check.ts` compares every
formatted cell of the corpus with LibreOffice's "as shown" CSV export:
7,590 of 7,603 cells in 32 formats match, and the 13 others are the
serial-0 and `-0` rows above.
