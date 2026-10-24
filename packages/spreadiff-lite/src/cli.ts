#!/usr/bin/env node
/**
 * spreadiff-lite OLD NEW [options]
 *
 * Prints what changed between two spreadsheets as Markdown (or JSON), and
 * exits 0 when they're the same, 1 when they differ, 2 on an error, like diff(1).
 * Also a git diff driver (--git), textconv (--textconv) and merge driver (merge).
 */
import { readFileSync, writeFileSync } from 'node:fs';
import { basename } from 'node:path';
import { parseFile } from 'spreadiff/core';
import { gitDiff, mergeFiles, textconv } from './git.js';
import { compare, type CompareOptions } from './report.js';
import { toMarkdown, type MarkdownOptions } from './markdown.js';

const HELP = `Usage: spreadiff-lite OLD NEW [options]
       spreadiff-lite --textconv FILE
       spreadiff-lite --git PATH OLD-FILE OLD-HEX OLD-MODE NEW-FILE NEW-HEX NEW-MODE
       spreadiff-lite merge BASE OURS THEIRS [-o OUT] [--name PATH] [--prefer ours|theirs] [--keys COL,…]

Compare two spreadsheets (.xlsx, .xlsm, .xlsb, .xls, .ods, .csv, .tsv): what was edited, what
moved downstream because of it, and what changed shape.

Options:
  --json               Print the report as JSON instead of Markdown
  --all                Don't cap the tables (default: 50 rows each)
  --max-rows N         Rows per table
  --outputs-only       List final outputs, not intermediate results
  --sheet NAME         Only this sheet (repeatable)
  --keys COL[,COL]     Match rows on these columns
  --header auto|yes|no Whether row 1 holds column names (default auto)
  --no-recalc          Follow formula references instead of recalculating (faster; may over-attribute)
  --live NAME[,NAME]   More data-feed functions to treat as inputs (Provider:FN,FN or FN)
  --tolerance X        Numbers within X are equal
  --min-change P       Ignore moves smaller than P of the number (0.1% or 0.001)
  --textconv           Print FILE as line-per-row text, for git diff's textconv
  --git                Act as git's external diff driver (diff.<driver>.command)
  -h, --help           Show this help

merge: three-way merge of CSV/TSV files, for a git merge driver
(merge.<driver>.driver = spreadiff-lite merge %O %A %B -o %A --name %P).
Conflicts take ours (or --prefer theirs) and are listed on stderr.

Exit status: 0 no changes, 1 changes, 2 error. --git and --textconv exit 0;
merge exits 1 when there were conflicts (or for .xlsx, which it can't write).

See docs/git.md for the git configuration.`;

function fail(message: string): never {
  process.stderr.write(`spreadiff-lite: ${message}\n`);
  process.exit(2);
}

function main(argv: string[]) {
  const files: string[] = [];
  const options: CompareOptions = {};
  const md: MarkdownOptions = {};
  let json = false;
  let git = false;
  let conv = false;
  const sheets: string[] = [];
  const live: Record<string, string[]> = {};
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    const value = () => {
      const v = argv[++i];
      if (v === undefined) fail(`${a} needs a value`);
      return v;
    };
    switch (a) {
      case '-h':
      case '--help':
        process.stdout.write(HELP + '\n');
        process.exit(0);
      case '--json':
        json = true;
        break;
      case '--git':
        git = true;
        break;
      case '--textconv':
        conv = true;
        break;
      case '--all':
        md.maxRows = Infinity;
        break;
      case '--max-rows': {
        const n = Number(value());
        if (!(n > 0)) fail('--max-rows needs a positive number');
        md.maxRows = n;
        break;
      }
      case '--outputs-only':
        md.intermediate = false;
        break;
      case '--sheet':
        sheets.push(value());
        break;
      case '--keys':
        options.diff = { ...options.diff, keys: value().split(',').map((s) => s.trim()) };
        break;
      case '--header': {
        const v = value();
        const header = v === 'auto' ? 'auto' : v === 'yes' || v === 'true' ? true : v === 'no' || v === 'false' ? false : fail('--header is auto, yes or no');
        options.parse = { ...options.parse, header };
        break;
      }
      case '--no-recalc':
        options.trace = { ...options.trace, recalculate: false };
        break;
      case '--live': {
        const v = value();
        const [provider, names] = v.includes(':') ? v.split(':', 2) : ['custom', v];
        (live[provider] ??= []).push(...names.split(',').map((s) => s.trim()).filter(Boolean));
        break;
      }
      case '--tolerance': {
        const n = Number(value());
        if (!(n >= 0)) fail('--tolerance needs a number');
        options.diff = { ...options.diff, numericTolerance: n };
        break;
      }
      case '--min-change': {
        const v = value();
        const n = v.endsWith('%') ? Number(v.slice(0, -1)) / 100 : Number(v);
        if (!(n >= 0)) fail('--min-change needs a fraction or a percentage (0.001 or 0.1%)');
        options.diff = { ...options.diff, relativeTolerance: n };
        break;
      }
      default:
        if (a.startsWith('-')) fail(`unknown option ${a} (see --help)`);
        files.push(a);
    }
  }
  if (sheets.length) options.sheets = sheets;
  if (Object.keys(live).length) options.trace = { ...options.trace, liveFunctions: live };
  if (conv) {
    if (files.length !== 1) fail('--textconv needs one file');
    try {
      process.stdout.write(textconv(parseFile(files[0], readFileSync(files[0]), { header: 'auto', ...options.parse })));
    } catch (e) {
      fail((e as Error).message);
    }
    return;
  }
  if (git) {
    // git passes 7 arguments; an added or deleted file is /dev/null on one side.
    if (files.length !== 7) fail(`--git needs the 7 arguments git passes to a diff driver (got ${files.length})`);
    try {
      process.stdout.write(gitDiff(files, (p) => readFileSync(p), { ...options, ...md }));
    } catch (e) {
      fail((e as Error).message);
    }
    return;
  }
  if (files.length !== 2) fail(files.length ? 'needs two files' : 'needs two files (see --help)');

  const read = (path: string) => {
    try {
      return { name: basename(path), data: readFileSync(path) };
    } catch (e) {
      fail(`can't read ${path}: ${(e as Error).message}`);
    }
  };
  let report;
  try {
    report = compare(read(files[0]), read(files[1]), options);
  } catch (e) {
    fail((e as Error).message);
  }
  process.stdout.write(json ? JSON.stringify(report, null, 2) + '\n' : toMarkdown(report, md));
  const changed = report.edits.length + report.downstream.length + report.structure.length + report.objects.length > 0;
  process.exitCode = changed ? 1 : 0;
}

function mergeMain(argv: string[]) {
  const files: string[] = [];
  let out: string | undefined;
  let name: string | undefined;
  let prefer: 'ours' | 'theirs' = 'ours';
  let keys: string[] | undefined;
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    const value = () => argv[++i] ?? fail(`${a} needs a value`);
    if (a === '-o' || a === '--output') out = value();
    else if (a === '--name') name = value();
    else if (a === '--prefer') {
      const v = value();
      if (v !== 'ours' && v !== 'theirs') fail('--prefer is ours or theirs');
      prefer = v;
    } else if (a === '--keys') keys = value().split(',').map((s) => s.trim());
    else if (a.startsWith('-')) fail(`unknown merge option ${a}`);
    else files.push(a);
  }
  if (files.length !== 3) fail('merge needs BASE OURS THEIRS');
  const read = (p: string) => {
    try {
      return new Uint8Array(readFileSync(p));
    } catch (e) {
      fail(`can't read ${p}: ${(e as Error).message}`);
    }
  };
  let result;
  try {
    result = mergeFiles(name ?? files[1], read(files[0]), read(files[1]), read(files[2]), { prefer, keys });
  } catch (e) {
    fail((e as Error).message);
  }
  for (const c of result.conflicts) process.stderr.write(`spreadiff-lite: conflict: ${c}\n`);
  if (result.csv === null) {
    process.stderr.write(`spreadiff-lite: ${name ?? files[1]}: workbooks aren't merged automatically; ${result.conflicts.length} conflicting change(s) listed above.\n`);
    process.exitCode = 1;
    return;
  }
  if (out) writeFileSync(out, result.csv);
  else process.stdout.write(result.csv);
  process.exitCode = result.conflicts.length ? 1 : 0;
}

const args = process.argv.slice(2);
if (args[0] === 'merge') mergeMain(args.slice(1));
else main(args);
