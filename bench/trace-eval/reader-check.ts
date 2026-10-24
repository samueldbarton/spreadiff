// The .ods, .xls and .xlsb readers against LibreOffice: each file is read by
// spreadiff, and also converted to .xlsx by LibreOffice and read with the
// .xlsx reader; every cell's value and formula should agree.
// With --via-ods, .xlsx files are first converted to .ods (and compared with
// themselves), which gives the .ods reader a large corpus.
// Usage: npx vite-node reader-check.ts [--via-ods] FILE…   (needs soffice)
import { execFileSync } from 'node:child_process';
import { existsSync, mkdirSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { basename, extname, join } from 'node:path';
import { parseFile, type Table, type Workbook } from '../../packages/spreadiff/src/core/index';
import { columnLabel } from '../../packages/spreadiff/src/core/csv';

const args = process.argv.slice(2);
const viaOds = args.includes('--via-ods');
const files = args.filter((a) => !a.startsWith('--'));
const out = join(tmpdir(), 'spreadiff-reader-check');
const convert = (file: string, to: string): string | null => {
  rmSync(out, { recursive: true, force: true });
  mkdirSync(out, { recursive: true });
  try {
    execFileSync('soffice', ['--headless', '--convert-to', to, '--outdir', out, file], { stdio: 'ignore', timeout: 120_000 });
  } catch {
    return null;
  }
  const p = join(out, `${basename(file, extname(file))}.${to}`);
  return existsSync(p) ? p : null;
};

/** Every non-empty cell of a sheet, by A1 address. */
function cells(t: Table): Map<string, { v: unknown; f: string | null }> {
  const m = new Map<string, { v: unknown; f: string | null }>();
  const o = t.origin ?? { row: 0, col: 0 };
  t.rows.forEach((row, r) =>
    row.forEach((v, c) => {
      const f = t.formulas?.[r]?.[c] ?? null;
      if ((v === null || v === '') && !f) return;
      m.set(`${columnLabel(c + o.col)}${r + o.row + 1}`, { v, f });
    }),
  );
  return m;
}
// Cosmetic differences between writers: Excel quotes sheet names that look like cells ('S2'!A1),
// LibreOffice writes a 3-D single cell as A1:A1.
const normF = (f: string | null) =>
  f
    ? f
        .replace(/^=/, '')
        .replace(/_xlfn\.|_xlws\.|_xludf\./gi, '')
        .replace(/'([A-Za-z_][\w.]*)'!/g, '$1!')
        .replace(/!(\$?[A-Z]+\$?\d+):\1(?![\d$A-Z])/gi, '!$1')
        // LibreOffice has no boolean literal (FALSE()), and spells whole columns out to row 1048576.
        .replace(/\b(TRUE|FALSE)\(\)/gi, '$1')
        .replace(/(\$?[A-Z]+)\$?1:(\$?[A-Z]+)\$?1048576(?!\d)/gi, '$1:$2')
        .replace(/\s+/g, '')
        .toUpperCase()
    : null;
const sameV = (a: unknown, b: unknown) => {
  if (a === b) return true;
  if (typeof a === 'number' && typeof b === 'number') return Math.abs(a - b) <= 1e-9 * Math.max(1, Math.abs(a), Math.abs(b));
  if ((a === null || a === '') && (b === null || b === '')) return true;
  return false;
};

let compared = 0;
let values = 0;
let formulas = 0;
let formulaCells = 0;
let done = 0;
const failed: string[] = [];
const examples: string[] = [];
for (const file of files) {
  let ours: Workbook;
  let theirs: Workbook;
  try {
    if (viaOds) {
      const ods = convert(file, 'ods');
      if (!ods) throw new Error('LibreOffice could not convert it');
      ours = parseFile('x.ods', readFileSync(ods), { header: false });
      theirs = parseFile('x.xlsx', readFileSync(file), { header: false });
    } else {
      ours = parseFile(basename(file), readFileSync(file), { header: false });
      const x = convert(file, 'xlsx');
      if (!x) throw new Error('LibreOffice could not convert it');
      theirs = parseFile('x.xlsx', readFileSync(x), { header: false });
    }
  } catch (e) {
    failed.push(`${basename(file)}: ${(e as Error).message}`);
    continue;
  }
  done++;
  for (const t of theirs.sheets) {
    const mine = ours.sheets.find((s) => s.name === t.name);
    if (!mine) {
      if (t.rows.length) failed.push(`${basename(file)}: sheet ${t.name} missing`);
      continue;
    }
    const a = cells(mine);
    const b = cells(t);
    for (const k of new Set([...a.keys(), ...b.keys()])) {
      const x = a.get(k) ?? { v: null, f: null };
      const y = b.get(k) ?? { v: null, f: null };
      compared++;
      const vOk = sameV(x.v, y.v);
      if (vOk) values++;
      if (x.f || y.f) {
        formulaCells++;
        if (normF(x.f) === normF(y.f)) formulas++;
        else if (examples.length < (process.env.ALL ? 1e9 : 60)) examples.push(`${basename(file)} ${t.name}!${k} formula ${JSON.stringify(x.f)} vs ${JSON.stringify(y.f)}`);
      }
      if (!vOk && examples.length < (process.env.ALL ? 1e9 : 60)) examples.push(`${basename(file)} ${t.name}!${k} value ${JSON.stringify(x.v)} vs ${JSON.stringify(y.v)}`);
    }
  }
}
const pct = (n: number, d: number) => `${((100 * n) / Math.max(1, d)).toFixed(2)}%`;
console.log(`${done} of ${files.length} files read · ${compared} cells · values ${values} (${pct(values, compared)}) · formulas ${formulas} of ${formulaCells} (${pct(formulas, formulaCells)})`);
if (failed.length) console.log(`failed:\n  ${failed.join('\n  ')}`);
console.log(examples.join('\n'));
