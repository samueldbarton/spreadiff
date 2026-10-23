// Evaluate traceWorkbook against the LibreOffice oracle.
// Usage: node eval.mjs <casesDir> [case...]
import { readFileSync, readdirSync, existsSync } from 'node:fs';
import { join } from 'node:path';
import { diffWorkbooks, parseXlsx, traceWorkbook, columnLabel, type WorkbookDiff, type CellRef } from '../../packages/spreadiff/src/core/index';

const dir = process.argv[2];
const only = process.argv.slice(3);
const read = (p: string) => parseXlsx(new Uint8Array(readFileSync(p)), { name: p });

/** Aligned ref → "Sheet!A1" in the new (or, if absent, old) workbook's coordinates. */
function addressOf(wd: WorkbookDiff, ref: CellRef, side: 'new' | 'old' = 'new'): string | null {
  const s = wd.sheets[ref.sheet];
  const d = s.diff;
  const t = side === 'new' ? d.newTable : d.oldTable;
  const ri = side === 'new' ? d.newRows[ref.row] : d.oldRows[ref.row];
  const ci = side === 'new' ? d.columns[ref.col].newIndex : d.columns[ref.col].oldIndex;
  if (ri < 0 || ci < 0) return null;
  const o = t.origin ?? { row: 1, col: 0 };
  return `${t.name}!${columnLabel(ci + o.col)}${ri + o.row + 1}`;
}

/** Every changed cell of a workbook diff, by address. */
function changedAddresses(wd: WorkbookDiff): Set<string> {
  const out = new Set<string>();
  wd.sheets.forEach((s, si) => {
    for (const [row, mask] of s.diff.changedCells) {
      for (let col = 0; col < mask.length; col++) if (mask[col]) {
        const a = addressOf(wd, { sheet: si, row, col });
        if (a) out.add(a);
      }
    }
  });
  return out;
}

const time = <T>(f: () => T): [T, number] => {
  const t = performance.now();
  const r = f();
  return [r, performance.now() - t];
};

const totals = { cells: 0, kindOk: 0, pairsTP: 0, pairsFP: 0, pairsFN: 0, exact: 0, computed: 0, roundtripFiles: 0, roundtripDirty: 0, roundtripCells: 0, trace: 0 };
const cases = (only.length ? only : readdirSync(dir)).filter((c) => existsSync(join(dir, c, 'manifest.json')));
for (const name of cases) {
  const base = join(dir, name);
  const m = JSON.parse(readFileSync(join(base, 'manifest.json'), 'utf8'));
  const [[oldWb, newWb], tParse] = time(() => [read(join(base, 'old.xlsx')), read(join(base, 'new.xlsx'))]);
  const [wd, tDiff] = time(() => diffWorkbooks(oldWb, newWb));
  const [tr, tTrace] = time(() => traceWorkbook(wd));
  const valueEdits = new Set<string>();
  const formulaEdits = new Set<string>();
  const editIndex = new Map<string, number>();
  m.edits.forEach((e: any, i: number) => {
    if (e.kind === 'value') valueEdits.add(`${e.sheet}!${e.newAddress}`);
    if (e.kind === 'formula') formulaEdits.add(`${e.sheet}!${e.newAddress}`);
    if (e.newAddress) editIndex.set(`${e.sheet}!${e.newAddress}`, i);
  });
  // Oracle: which edits, applied alone, move each cell.
  const movedBy = new Map<string, Set<number>>();
  m.edits.forEach((e: any, i: number) => {
    if (!e.single) return;
    const sd = diffWorkbooks(oldWb, read(join(base, e.single)));
    for (const a of changedAddresses(sd)) {
      if (!movedBy.has(a)) movedBy.set(a, new Set());
      movedBy.get(a)!.add(i);
    }
  });
  // But-for oracle: edit i caused C if C differs between new and new-without-i.
  const butFor = new Map<string, Set<number>>();
  m.edits.forEach((e: any, i: number) => {
    const loo = e.loo ?? (m.edits.length === 1 ? 'old.xlsx' : null);
    if (!loo) return;
    const ld = diffWorkbooks(read(join(base, loo)), newWb);
    for (const a of changedAddresses(ld)) {
      if (!butFor.has(a)) butFor.set(a, new Set());
      butFor.get(a)!.add(i);
    }
  });
  const truthCauses = (a: string) => {
    const bf = butFor.get(a);
    return bf && bf.size ? bf : movedBy.get(a) ?? new Set<number>();
  };
  const hasInsert = m.edits.some((e: any) => e.kind === 'insert');
  let cells = 0, kindOk = 0, tp = 0, fp = 0, fn = 0, exact = 0, computed = 0;
  const problems: string[] = [];
  wd.sheets.forEach((s, si) => {
    for (const [k, t] of tr.sheet(si)) {
      if (t.kind === 'added' || t.kind === 'removed') continue;
      const ref = { sheet: si, row: Math.floor(k / s.diff.columns.length), col: k % s.diff.columns.length };
      const a = addressOf(wd, ref)!;
      const truth = valueEdits.has(a) ? 'input' : formulaEdits.has(a) ? 'formula' : 'computed';
      cells++;
      if (truth === t.kind) kindOk++;
      else problems.push(`kind ${a}: got ${t.kind}, want ${truth}`);
      if (t.kind !== 'computed' || truth !== 'computed') continue;
      computed++;
      // Causes: compare with the counterfactual oracle (only when every edit has a single run).
      const got = new Set<number>();
      let structural = false;
      for (const c of t.causes) {
        const ca = addressOf(wd, c) ?? addressOf(wd, c, 'old');
        const ce = tr.get(c);
        if (ce?.kind === 'added' || ce?.kind === 'removed') { structural = true; continue; }
        const i = ca ? editIndex.get(ca) : undefined;
        if (i === undefined) problems.push(`cause of ${a} is not an edit: ${ca} (${ce?.kind})`);
        else got.add(i);
      }
      if (t.causeCount > t.causes.length) problems.push(`causes of ${a} truncated`);
      const want = truthCauses(a);
      if (hasInsert && (structural || want.size === 0)) continue; // insert causes have no single-run oracle
      let ok = true;
      for (const i of got) if (want.has(i)) tp++; else { fp++; ok = false; problems.push(`extra cause ${a} ← edit ${i} (${m.edits[i].sheet}!${m.edits[i].newAddress})`); }
      for (const i of want) if (!got.has(i)) { fn++; ok = false; problems.push(`missing cause ${a} ← edit ${i} (${m.edits[i].sheet}!${m.edits[i].newAddress})`); }
      if (ok) exact++;
    }
  });
  // Edits the diff should have flagged as changed.
  for (const a of [...valueEdits, ...formulaEdits]) {
    const found = [...changedAddresses(wd)].includes(a);
    if (!found) problems.push(`edit not detected: ${a}`);
  }
  totals.cells += cells; totals.kindOk += kindOk; totals.pairsTP += tp; totals.pairsFP += fp; totals.pairsFN += fn; totals.exact += exact; totals.computed += computed; totals.trace += tTrace;
  console.log(`\n## ${name}: ${cells} changed cells, ${computed} computed · kinds ${kindOk}/${cells} · causes exact ${exact}/${computed} (pairs tp ${tp} fp ${fp} fn ${fn})`);
  console.log(`   parse ${tParse.toFixed(0)} ms · diff ${tDiff.toFixed(1)} ms · trace ${tTrace.toFixed(1)} ms · stats ${JSON.stringify(tr.stats)}${tr.truncated ? ' TRUNCATED' : ''}`);
  const uniq = [...new Set(problems)];
  for (const p of uniq.slice(0, 12)) console.log('   - ' + p);
  if (uniq.length > 12) console.log(`   … ${uniq.length - 12} more`);
  if (existsSync(join(base, 'excel-original.xlsx'))) {
    const rd = diffWorkbooks(read(join(base, 'excel-original.xlsx')), oldWb);
    const rt = traceWorkbook(rd);
    const n = changedAddresses(rd).size;
    totals.roundtripFiles++;
    if (n) totals.roundtripDirty++;
    totals.roundtripCells += n;
    console.log(`   excel → libreoffice round trip, no edits: ${n} changed cells ${n ? JSON.stringify(rt.stats) + ' e.g. ' + [...changedAddresses(rd)].slice(0, 4).join(', ') : ''}`);
  }
  // Failure scenarios: stale and uncached files.
  for (const variant of ['stale', 'uncached'] as const) {
    if (!m[variant]) continue;
    const vd = diffWorkbooks(oldWb, read(join(base, m[variant])));
    const vt = traceWorkbook(vd);
    console.log(`   ${variant}: ${JSON.stringify(vt.stats)} (recalculated new has ${JSON.stringify(tr.stats)})`);
  }
}
const pct = (a: number, b: number) => (b ? ((100 * a) / b).toFixed(1) + '%' : 'n/a');
console.log(`\nround trips: ${totals.roundtripDirty}/${totals.roundtripFiles} files changed, ${totals.roundtripCells} cells · total trace time ${totals.trace.toFixed(0)} ms`);
console.log(`TOTAL kinds ${pct(totals.kindOk, totals.cells)} of ${totals.cells} · computed cells with exact causes ${pct(totals.exact, totals.computed)} of ${totals.computed} · cause precision ${pct(totals.pairsTP, totals.pairsTP + totals.pairsFP)} recall ${pct(totals.pairsTP, totals.pairsTP + totals.pairsFN)}`);
