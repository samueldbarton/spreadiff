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
