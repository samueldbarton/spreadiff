import { readFileSync } from 'node:fs';
import { parseXlsx, diffWorkbooks, traceWorkbook } from '../../packages/spreadiff/src/core/index';
for (const c of process.argv.slice(2)) {
  const t0 = performance.now();
  const a = parseXlsx(new Uint8Array(readFileSync(`scale-cases/${c}/old.xlsx`))), b = parseXlsx(new Uint8Array(readFileSync(`scale-cases/${c}/new.xlsx`)));
  const t1 = performance.now();
  const d = diffWorkbooks(a, b);
  const t2 = performance.now();
  const refs = traceWorkbook(d, { recalculate: false });
  const t3 = performance.now();
  const tr = traceWorkbook(d);
  const t4 = performance.now();
  const formulas = a.sheets.reduce((n, s) => n + (s.formulas?.reduce((m, r) => m + (r?.filter(Boolean).length ?? 0), 0) ?? 0), 0);
  console.log(`${c}: ${formulas} formulas · parse ${Math.round(t1 - t0)} ms · diff ${Math.round(t2 - t1)} ms · trace (references) ${Math.round(t3 - t2)} ms · trace (recalculated) ${Math.round(t4 - t3)} ms`);
  console.log('   references:', JSON.stringify(refs.stats), refs.truncated ? 'TRUNCATED' : '');
  console.log('   recalculated:', JSON.stringify(tr.stats), tr.truncated ? 'TRUNCATED' : '');
}
