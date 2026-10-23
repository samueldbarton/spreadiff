// Values computed for an uncached file vs LibreOffice's recalculated values, cell by cell.
import { readFileSync } from 'node:fs';
import { parseXlsx } from '../../packages/spreadiff/src/core/index';
import { engineValue } from '../../packages/spreadiff/src/core/engine/host';
import { sameResult } from '../../packages/spreadiff/src/core/engine/values';
for (const c of ['three_statement', 'inventory', 'loan', 'commission']) {
  const t0 = performance.now();
  const u = parseXlsx(new Uint8Array(readFileSync(`cases/${c}/uncached.xlsx`)));
  const ms = performance.now() - t0;
  const truth = parseXlsx(new Uint8Array(readFileSync(`cases/${c}/new.xlsx`)), { computeMissing: false });
  let n = 0, ok = 0; const bad: string[] = [];
  u.sheets.forEach((t, s) => t.formulas?.forEach((row, r) => row?.forEach((f, col) => {
    if (!f) return; n++;
    const a = engineValue(t.rows[r][col], false), b = engineValue(truth.sheets[s].rows[r]?.[col], false);
    if (sameResult(a, b)) ok++; else if (bad.length < 3) bad.push(`${t.name} r${r} c${col} ${f}: ${JSON.stringify(a)} vs ${JSON.stringify(b)}`);
  })));
  console.log(`${c}: ${ok}/${n} formula values equal LibreOffice's · computed ${JSON.stringify(u.computedValues)} · parse ${ms.toFixed(0)} ms`, bad.join(' | '));
}
