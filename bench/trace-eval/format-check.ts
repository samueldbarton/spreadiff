// How cells display, against LibreOffice: every cell with a number format,
// shown through spreadiff's formatter, compared with LibreOffice's "as shown"
// CSV export of the same file.
// Usage: npx vite-node format-check.ts work/poi-cases work/cases   (needs soffice)
import { execFileSync } from 'node:child_process';
import { existsSync, mkdirSync, readdirSync, readFileSync, rmSync, statSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { basename, join } from 'node:path';
import { parseCsv, parseXlsx } from '../../packages/spreadiff/src/core/index';
import { displayCell, formatCodeAt } from '../../packages/spreadiff/src/core/display';

const roots = process.argv.slice(2);
const files: string[] = [];
for (const root of roots) {
  for (const d of readdirSync(root)) {
    const f = join(root, d, 'excel-original.xlsx');
    const g = join(root, d, 'old.xlsx');
    if (existsSync(f)) files.push(f);
    else if (existsSync(g) && statSync(g).isFile()) files.push(g);
  }
}
const out = join(tmpdir(), 'spreadiff-format-check');
let total = 0;
let same = 0;
const byCode = new Map<string, { n: number; bad: number; example?: string }>();
for (const f of files) {
  rmSync(out, { recursive: true, force: true });
  mkdirSync(out, { recursive: true });
  try {
    execFileSync('soffice', ['--headless', '--convert-to', 'csv:Text - txt - csv (StarCalc):44,34,76,1,,1033,false,true,true,false,false,-1', '--outdir', out, f], { stdio: 'ignore', timeout: 120_000 });
  } catch {
    continue;
  }
  const wb = parseXlsx(new Uint8Array(readFileSync(f)), { header: false, computeMissing: false });
  const stem = basename(f, '.xlsx');
  for (const t of wb.sheets) {
    if (!t.numberFormats) continue;
    const csvPath = join(out, `${stem}-${t.name}.csv`);
    if (!existsSync(csvPath)) continue;
    const lo = parseCsv(readFileSync(csvPath, 'utf8'), { header: false });
    const o = t.origin ?? { row: 0, col: 0 };
    t.rows.forEach((row, r) =>
      row.forEach((v, c) => {
        const code = formatCodeAt(t, r, c);
        if (!code || v === null) return;
        const theirs = lo.rows[r + o.row]?.[c + o.col];
        if (theirs === undefined) return;
        const expected = theirs === null ? '' : String(theirs);
        const ours = displayCell(t, r, c);
        total++;
        const k = byCode.get(code) ?? { n: 0, bad: 0 };
        k.n++;
        // LibreOffice pads with the * fill to the column width; spaces from _x vary. Compare without outer spaces.
        if (ours.trim() === expected.trim()) same++;
        else {
          k.bad++;
          k.example ??= `${basename(join(f, '..'))} ${t.name}: ${JSON.stringify(v)} → ours ${JSON.stringify(ours)} LibreOffice ${JSON.stringify(expected)}`;
        }
        byCode.set(code, k);
      }),
    );
  }
}
console.log(`${files.length} files · ${total} formatted cells in ${byCode.size} distinct formats · same as LibreOffice ${same} (${((100 * same) / Math.max(1, total)).toFixed(1)}%)`);
if (process.env.CODES) console.log([...byCode].map(([c, k]) => `${k.n}\t${c}`).join('\n'));
const worst = [...byCode].filter(([, k]) => k.bad).sort((a, b) => b[1].bad - a[1].bad);
for (const [code, k] of worst.slice(0, 40)) console.log(`${String(k.bad).padStart(6)}/${String(k.n).padEnd(6)} ${JSON.stringify(code)}\n         ${k.example}`);
