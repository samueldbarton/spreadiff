# Spreadsheets in git

Git sees an .xlsx as a zip: `git diff` says "Binary files differ", and two
branches that edited different cells conflict on the whole file.
spreadiff-lite gives git three things:

| | What git does with it | Set up with |
| --- | --- | --- |
| [textconv](#readable-git-diff) | `git diff`, `git log -p` and `git show` print a sheet's rows as text and diff those lines | `diff.<driver>.textconv` |
| [diff driver](#the-full-report-in-git-diff) | `git diff` prints the full report: edits, downstream results, final outputs | `diff.<driver>.command` |
| [merge driver](#merging-csv-files) | Two branches that edited different rows of a CSV merge cleanly | `merge.<driver>.driver` |

For pull requests there's also a [GitHub Action](#pull-request-comments)
that posts the report as a comment.

All of them run `spreadiff-lite`, so install it where git runs (see
[Install](getting-started.md#install)), and check that
`spreadiff-lite --help` works.

## Readable `git diff`

A textconv turns a file into text before git diffs it. spreadiff-lite
prints each sheet under a `== Sheet ==` line and each row as its sheet row
number and cells, shown as the sheet formats them, with formulas in braces:

```text
== Summary ==
1: Region | Revenue | Share
2: North | $1,200 {=SUM(Data!B2:B9)} | 40.0% {=B2/B$5}
```

Then git's line diff shows which rows changed. Tell git which files to
convert, in `.gitattributes`:

```text
*.xlsx diff=spreadsheet
*.xlsm diff=spreadsheet
*.xlsb diff=spreadsheet
*.xls diff=spreadsheet
*.ods diff=spreadsheet
```

and what the `spreadsheet` driver runs, in your git config:

```bash
git config diff.spreadsheet.textconv "spreadiff-lite --textconv"
git config diff.spreadsheet.cachetextconv true   # convert each version once
```

Leave CSV files out: they're text already, and git diffs them well.

Textconv is line-based: a row inserted near the top shows as one added
line, but a value that moved because an input changed is just another
changed line. To see which edits caused which changes, use the diff driver.

## The full report in `git diff`

With `diff.<driver>.command`, git hands both versions to spreadiff-lite
and prints what it prints: the [report](spreadiff-lite.md#the-report),
with each changed result's causes. Use the same attributes and:

```bash
git config diff.spreadsheet.command "spreadiff-lite --git"
```

`git diff` then prints the report, and `git diff --no-ext-diff` the
textconv (or "Binary files differ" without one). `git show` and
`git log -p` use external drivers only when asked: `git log -p --ext-diff`.
Added and deleted files are compared with an empty workbook.

## Merging CSV files

A merge driver replaces git's line merge for matching files. spreadiff-lite
merges tables by row key instead of by line: two branches that edited
different cells of the same row, or added different rows, merge without
a conflict. In `.gitattributes`:

```text
*.csv merge=spreadsheet
*.tsv merge=spreadsheet
```

and in your config:

```bash
git config merge.spreadsheet.name "spreadsheet merge"
git config merge.spreadsheet.driver "spreadiff-lite merge %O %A %B -o %A --name %P"
```

When both branches changed the same cell differently, the merge takes
ours (`--prefer theirs` changes that), lists each conflict on stderr, and
exits 1 so git marks the file as conflicted: check the listed cells, then
`git add` it. Rows are matched as `diffTables` matches them (by a key column
it finds, or `--keys id,region`). See [`mergeTables`](api.md#mergetables).

Workbooks (.xlsx and the rest) aren't merged: the driver reports their
conflicts and exits 1, leaving our version in place, since writing a
workbook back would lose what the reader doesn't read (styles, charts,
macros).

## Pull-request comments

The action compares every spreadsheet a pull request changes and posts
the report as one comment, updated on later pushes (and in the job
summary):

```yaml
# .github/workflows/spreadsheets.yml
name: Spreadsheet changes
on: pull_request
permissions:
  contents: read
  pull-requests: write
jobs:
  report:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
        with:
          fetch-depth: 0 # the base commit's files are read from history
      - uses: samueldbarton/spreadiff@main
        with:
          min-change: '0.1%'
```

| Input | Default | |
| --- | --- | --- |
| `paths` | `**/*.{xlsx,xlsm,xlsb,xls,ods,csv,tsv}` | Files to compare: globs, comma- or newline-separated |
| `min-change` | none | Ignore moves smaller than this share of the number (`0.1%` or `0.001`) |
| `max-rows` | `25` | Rows per table |
| `comment` | `true` | `false` writes only the job summary |
| `base`, `head` | the pull request's | Commits to compare |
| `github-token` | `github.token` | Token to post the comment with |

Outputs: `files` (how many matching files changed) and `changed` (how many
differ in content). A comment longer than GitHub allows is cut, with the
whole report in the job summary. Pull requests from forks get a read-only
token, so only the job summary is written.

## How it can be wrong

- **Textconv hides what doesn't print.** A change of format, comment or
  style, or a value change too small for its format to show (`12.3%` before
  and after), produces no changed line. The report compares stored values,
  so it does see them.
- **The merge driver trusts the row key.** If no column identifies rows,
  rows are matched by position and content, and an inserted row with values
  like its neighbours can pair with the wrong one. Pass `--keys`.
- **The action reads files from git history.** With a shallow checkout the
  base commit isn't there; use `fetch-depth: 0`.
