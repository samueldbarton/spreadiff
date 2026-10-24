#!/usr/bin/env node
/**
 * spreadiff-lite OLD NEW [options]
 *
 * Prints what changed between two spreadsheets as Markdown (or JSON), and
 * exits 0 when they're the same, 1 when they differ, 2 on an error, like diff(1).
 */
import { readFileSync } from 'node:fs';
import { basename } from 'node:path';
import { compare, type CompareOptions } from './report.js';
import { toMarkdown, type MarkdownOptions } from './markdown.js';

const HELP = `Usage: spreadiff-lite OLD NEW [options]

Compare two spreadsheets (.xlsx, .xlsm, .csv, .tsv): what was edited, what
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
  -h, --help           Show this help

Exit status: 0 no changes, 1 changes, 2 error.`;

function fail(message: string): never {
  process.stderr.write(`spreadiff-lite: ${message}\n`);
  process.exit(2);
}

function main(argv: string[]) {
  const files: string[] = [];
  const options: CompareOptions = {};
  const md: MarkdownOptions = {};
  let json = false;
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
  if (files.length !== 2) fail(files.length ? 'needs two files' : 'needs two files (see --help)');
  if (sheets.length) options.sheets = sheets;
  if (Object.keys(live).length) options.trace = { ...options.trace, liveFunctions: live };

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
  const changed = report.edits.length + report.downstream.length + report.structure.length > 0;
  process.exitCode = changed ? 1 : 0;
}

main(process.argv.slice(2));
