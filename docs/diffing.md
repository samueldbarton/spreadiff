# How diffing works

A spreadsheet diff answers three questions:
1. Which columns correspond?
2. Which rows correspond?
3. Which cells of corresponding rows differ?

spreadiff answers them in that order. This page explains each step, the
options that steer it, and the cases where the answer may not be what you
expect.

## The result: aligned rows and columns

`diffTables(old, new)` returns a [`TableDiff`](api.md#tablediff): a list of
**aligned rows** over a list of **aligned columns**. An aligned row points
at a row of the old table, a row of the new table, or both:

| Aligned row | `oldRows[r]` | `newRows[r]` | `status[r]` |
| --- | --- | --- | --- |
| unchanged | old index | new index | `equal` |
| edited | old index | new index | `modified`, with a mask in `changedCells` |
| deleted | old index | `-1` | `removed` |
| inserted | `-1` | new index | `added` |

Aligned rows come in the new table's order, with removed rows placed where
they used to be. Columns work the same way (`oldIndex` / `newIndex`, -1 when
missing), with one more status: `renamed`.

Row data lives in typed arrays (`Int32Array`, `Uint8Array`), so the diff of
a million rows costs a few megabytes on top of the tables themselves.

<!-- example: anatomy -->
```ts
import { diffTables, parseCsv, ROW_STATUS_NAMES } from 'spreadiff';

const d = diffTables(parseCsv('k,v\na,1\nb,2\nc,3'), parseCsv('k,v\na,1\nb,5\nd,4'), { keys: ['k'] });
for (let r = 0; r < d.rowCount; r++) {
  console.log(ROW_STATUS_NAMES[d.status[r]], d.oldRows[r], d.newRows[r], d.changedCells.get(r) ?? '');
}
```

<!-- output: anatomy -->
```text
equal 0 0 
modified 1 1 {"0":0,"1":1}
removed 2 -1 
added -1 2 
```

## Step 1: columns

Columns are matched **by name**, so reordering columns is not a change. A
name present only in the new table is `added`; only in the old, `removed`.

**Renames.** A removed and an added column are paired as `renamed` when at
least 80% of their values agree across matched rows. Rows where both cells
are blank are skipped, so two mostly-empty columns don't pair on their
blanks. `detectRenames: false` turns this off.

**No header row?** Parse with `header: false`. Columns are then named `A`,
`B`, … and match by position.

## Step 2: rows

There are two strategies.

### By key (when rows have an identity)

When rows have an id, matching on it is exact: rows can be reordered,
inserted and deleted, and edits are still found.

- `keys: ['id']` uses the columns you name. Composite keys work (`['region', 'month']`). Duplicate keys match in order.
- `keys: 'auto'` (the default) uses the **leftmost id-like column**: `id`, `ID`, `key`, `sku`, `code`, `uuid`, `guid`, `Order ID`, `order_id`, `customerId`, … The column must also be unique and non-empty on both sides, and match at least half the rows. A column called `name` is never picked automatically: renaming a product would then read as a deletion plus an insertion.
- `keys: 'none'` always matches by order.

`diff.keyColumns` tells you which strategy ran: empty means "by order".

### By order (when they don't)

Without keys, rows are aligned like lines in a text diff:

1. Each row is reduced to its content, and the longest common subsequence of
   identical rows is found (Myers' algorithm; it falls back to
   patience-style anchoring when the tables are huge and very different).
2. Between two matched rows, the leftover removed and added rows are
   **paired by similarity**. A removed and an added row become one
   `modified` row when at least half their cells (`rowSimilarity`, default
   0.5) are equal. A single removed and a single added row at the same
   position pair whenever any cell matches: that is almost always an edit.

**Cost.** Matching by key is the faster path: on a million rows with ~0.8%
changed, about 1–3 s by key against 3–5 s by order, in Node. See
[Performance](performance.md#keyed-vs-unkeyed).

**Formula cells don't count against similarity.** When you change one
input, every formula in that row may change too. Comparing rows on their
typed-in values keeps the row recognisable as "the same row, edited".

## Step 3: cells

Two cells are equal when:

- they are the same value, or
- with `numericTolerance > 0`, both read as numbers within that tolerance, or
- their normalized text matches. By default whitespace is trimmed
  (`trimWhitespace`) and case matters (`ignoreCase: false`).

Values keep their type: `1` and `"1"` compare equal (both read as `1`), but
`"1.0"` and `1` don't, because the text differs. Use `numericTolerance` to
compare numerically.

### Formulas

When both files carry formulas (XLSX), a cell also counts as changed when
its formula's **meaning** changed, even if the cached value didn't
(`compareFormulas`, default on). Formulas are compared by what they point
at, not by their text:

| Old | New | Changed? | Why |
| --- | --- | --- | --- |
| `=B5*C5` | `=B6*C6`, after a row was inserted above | no | Both read the same aligned cells |
| `=SUM(D2:D5)` | `=SUM(D2:D6)`, after a row was inserted in the range | no | Same surviving rows; the total moved *because* a row was added |
| `=SUM(D2:D5)` | `=SUM(D2:D4)` | **yes** | It now leaves out a row that still exists |
| `=B2*0.38` | `=B2*0.42` | **yes** | Different formula |

Cells whose formula changed are listed in `diff.formulaChanged`. To learn
*why* formula results changed, see [Tracing](tracing.md).

## Workbooks

`diffWorkbooks(old, new)` pairs sheets **by name** and diffs each pair.
Sheets present on one side only are `added` / `removed`, diffed against an
empty table. Two single-sheet workbooks (two CSV files, say) are paired
regardless of name.

Sheets are aligned first and compared second, so a formula on one sheet
that references another (`=SUM(Lines!D2:D9)`) is compared through that
sheet's alignment.

## Parsing

| | CSV / TSV (`parseCsv`) | XLSX (`parseXlsx`) |
| --- | --- | --- |
| Encoding | UTF-8; UTF-16 with a BOM (Excel's "Unicode text") via `parseFile` | — |
| Delimiter | detected (`,` `\t` `;` `\|`), or Excel's `sep=;` first line, or `delimiter` | — |
| Numbers | converted only when lossless: `"007"`, `"1.50"`, `"1e5"` stay text | stored values |
| Dates | text | ISO strings, from the number format (`2026-01-31`) |
| Formulas | — | text of each cell, shared formulas expanded; values are the cached results |
| Header | first row (`header: false` for none) | first non-empty row of the used range |

`parseFile(name, bytes)` picks the parser: zip content means XLSX, `.tsv`
means tab-delimited, anything else is CSV. Legacy `.xls` (BIFF) files are
rejected with a clear error.

**Why lossless numbers?** Converting `"007"` to `7` would hide a real
change, and showing `1.50` as `1.5` would misrepresent the file. Numbers
are also much cheaper to diff and to send across a worker boundary than
strings.

## Options at a glance

| Option | Default | Effect |
| --- | --- | --- |
| `keys` | `'auto'` | `string[]`, `'auto'` or `'none'`; see [rows](#step-2-rows) |
| `numericTolerance` | `0` | Numbers within ± this are equal |
| `trimWhitespace` | `true` | Ignore leading/trailing whitespace in text |
| `ignoreCase` | `false` | Compare text case-insensitively |
| `compareFormulas` | `true` | Count formula-meaning changes, XLSX vs. XLSX |
| `rowSimilarity` | `0.5` | Share of equal cells needed to pair rows as modified |
| `detectRenames` | `true` | Pair look-alike removed/added columns as renames |

## When results surprise you

**"Every row shows as changed."** The rows were probably reordered and no
key was found. Check `diff.keyColumns`; if it's empty, pass `keys: ['your id column']`.

**"A row I edited shows as removed + added."** More than half its cells
changed, so it no longer looks like the same row. If rows have an id, match
by key; otherwise lower `rowSimilarity`.

**"Numbers that look equal are flagged."** `0.30000000000000004` vs. `0.3`
displays the same but differs. Set `numericTolerance` (e.g. `1e-9`). The
grid shows full precision when two different values would display
identically.

**"A renamed column shows as removed + added."** Fewer than 80% of its
values matched, or it was renamed *and* heavily edited. That's ambiguous by
design: spreadiff won't guess.

**"Whitespace differences are hidden."** `trimWhitespace` is on by default.
Turn it off to see them.
