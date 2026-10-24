// Evaluate every formula with cached inputs; compare with the cached result.
import { readFileSync, readdirSync, existsSync } from 'node:fs';
import { join } from 'node:path';
import { parseXlsx } from '../../packages/spreadiff/src/core/index';
import { Engine, Unsupported } from '../../packages/spreadiff/src/core/engine/evaluate';
import { tableHost, engineValue } from '../../packages/spreadiff/src/core/engine/host';
import { isError, sameResult } from '../../packages/spreadiff/src/core/engine/values';

const files: string[] = [];
for (const root of process.argv.slice(2)) {
  for (const d of readdirSync(root)) {
    for (const f of ['old.xlsx', 'excel-original.xlsx']) if (existsSync(join(root, d, f))) files.push(join(root, d, f));
  }
}
let total = 0, ok = 0, unsupported = 0, mismatch = 0, uncached = 0;
const reasons = new Map<string, number>();
const examples: string[] = [];
const t0 = performance.now();
for (const f of files) {
  const wb = parseXlsx(new Uint8Array(readFileSync(f)));
  const host = tableHost(wb.sheets, { names: wb.names, tables: wb.tables, date1904: wb.date1904 });
  const engine = new Engine(host);
  wb.sheets.forEach((t, s) => {
    const o = t.origin!;
    t.formulas?.forEach((row, r) => row?.forEach((formula, c) => {
      if (!formula) return;
      total++;
      const cached = engineValue(t.rows[r]?.[c], wb.date1904 ?? false);
      if (cached === null) { uncached++; return; }
      try {
        const { value } = engine.evaluate(formula, t.name, r + o.row, c + o.col);
        if (sameResult(value, cached, wb.date1904)) ok++;
        else {
          mismatch++;
          if (examples.length < 40) examples.push(`${f.split('/').slice(-2).join('/')} ${t.name}!R${r + o.row + 1}C${c + o.col + 1} ${formula.slice(0, 70)} → got ${JSON.stringify(isError(value) ? value.code : value)} cached ${JSON.stringify(isError(cached) ? cached.code : cached)}`);
        }
      } catch (e) {
        if (!(e instanceof Unsupported)) throw e;
        unsupported++;
        const k = e.message.replace(/^function /, 'fn ');
        reasons.set(k, (reasons.get(k) ?? 0) + 1);
      }
    }));
  });
}
const ms = performance.now() - t0;
console.log(`${files.length} files · ${total} formulas · cached ${total - uncached} · match ${ok} (${(100 * ok / (total - uncached)).toFixed(1)}%) · mismatch ${mismatch} · unsupported ${unsupported} · ${ms.toFixed(0)} ms`);
console.log('unsupported:', [...reasons].sort((a, b) => b[1] - a[1]).slice(0, 25).map(([k, n]) => `${k} ${n}`).join(' · '));
console.log(examples.join('\n'));
