// Evaluate every probe in work/probe.xlsx with spreadiff's engine and compare
// with LibreOffice's saved result. Prints the divergences, grouped.
// Usage: npx tsx compare.ts   (after build.py)
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { parseXlsx } from '../../packages/spreadiff/src/core/index';
import { Engine, Unsupported } from '../../packages/spreadiff/src/core/engine/evaluate';
import { tableHost, engineValue } from '../../packages/spreadiff/src/core/engine/host';
import { isError, sameResult, type Scalar } from '../../packages/spreadiff/src/core/engine/values';

const file = fileURLToPath(new URL('work/probe.xlsx', import.meta.url));
const wb = parseXlsx(new Uint8Array(readFileSync(file)), { header: false, computeMissing: false });
const host = tableHost(wb.sheets, { names: wb.names, tables: wb.tables, date1904: wb.date1904 });
const engine = new Engine(host);
const p = wb.sheets.find((s) => s.name === 'P')!;
const o = p.origin!;
const show = (v: Scalar) => (isError(v) ? v.code : JSON.stringify(v));
let ok = 0;
let unsupported = 0;
const diffs: string[] = [];
const unsup = new Map<string, number>();
p.formulas!.forEach((row, r) => {
  const f = row?.[0];
  if (!f) return;
  const cached = engineValue(p.rows[r][0], false);
  let got: Scalar;
  try {
    got = engine.evaluate(f, 'P', r + o.row, o.col).value;
  } catch (e) {
    if (!(e instanceof Unsupported)) throw e;
    unsupported++;
    unsup.set(e.message, (unsup.get(e.message) ?? 0) + 1);
    return;
  }
  if (sameResult(got, cached, wb.date1904)) ok++;
  else diffs.push(`R${r + o.row + 1}  ${f}\n      engine ${show(got)}   LibreOffice ${show(cached)}`);
});
const total = ok + diffs.length + unsupported;
console.log(`${total} probes · match ${ok} · differ ${diffs.length} · unsupported ${unsupported}`);
if (unsup.size) console.log('unsupported:', [...unsup].map(([k, n]) => `${k} ×${n}`).join(' · '));
console.log(diffs.join('\n'));
