// Copies a selection of evaluated cases into test/hosted/fixtures with their
// expected results, so `npm run test:hosted` can check tracing in a browser
// without LibreOffice. The expectations come from LibreOffice's recalculated
// files (the but-for oracle in eval.ts), never from spreadiff's own trace.
// Usage (after npm run bench:trace): npx tsx bench/trace-eval/export-fixtures.ts
import { copyFileSync, existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { columnLabel, diffWorkbooks, parseXlsx, type CellValue, type Workbook, type WorkbookDiff } from '../../packages/spreadiff/src/core/index';

const HERE = dirname(fileURLToPath(import.meta.url));
const WORK = join(HERE, 'work');
const OUT = join(HERE, '../../test/hosted/fixtures');
const CASES: [dir: string, name: string][] = [
  ...['three_statement', 'budget', 'inventory', 'loan', 'commission'].map((n): [string, string] => ['cases', n]),
  ...[
    'Tables', '55906-MultiSheetRefs', '51710', 'DataTableCities', 'MatrixFormulaEvalTestData', 'NewlineInFormulas',
    '47813', 'CVLKRA-KYC_Download_File_Structure_V3.1', 'bug66215', '48539', 'clone_sheet', 'testSharedFormulasRangeSetBlankBug', '61652', 'table-sample',
  ].map((n): [string, string] => ['poi-cases', n]),
];

const read = (p: string) => parseXlsx(new Uint8Array(readFileSync(p)), { name: p });

function addressOf(wd: WorkbookDiff, sheet: number, row: number, col: number, side: 'new' | 'old' = 'new'): string | null {
  const d = wd.sheets[sheet].diff;
  const t = side === 'new' ? d.newTable : d.oldTable;
  const ri = side === 'new' ? d.newRows[row] : d.oldRows[row];
  const ci = side === 'new' ? d.columns[col].newIndex : d.columns[col].oldIndex;
  if (ri < 0 || ci < 0) return null;
  const o = t.origin ?? { row: 1, col: 0 };
  return `${t.name}!${columnLabel(ci + o.col)}${ri + o.row + 1}`;
}

function changed(wd: WorkbookDiff): Set<string> {
  const out = new Set<string>();
  wd.sheets.forEach((s, si) => {
    for (const [row, mask] of s.diff.changedCells) for (let col = 0; col < mask.length; col++) if (mask[col]) {
      const a = addressOf(wd, si, row, col);
      if (a) out.add(a);
    }
  });
  return out;
}

/** Every formula cell's value, by address. */
function formulaValues(wb: Workbook): Record<string, CellValue> {
  const out: Record<string, CellValue> = {};
  for (const t of wb.sheets) {
    const o = t.origin ?? { row: 1, col: 0 };
    t.formulas?.forEach((row, r) => row?.forEach((f, c) => {
      if (f) out[`${t.name}!${columnLabel(c + o.col)}${r + o.row + 1}`] = t.rows[r]?.[c] ?? null;
    }));
  }
  return out;
}

rmSync(OUT, { recursive: true, force: true });
let total = 0;
for (const [dir, name] of CASES) {
  const base = join(WORK, dir, name);
  const m = JSON.parse(readFileSync(join(base, 'manifest.json'), 'utf8'));
  const oldWb = read(join(base, 'old.xlsx'));
  const newWb = read(join(base, 'new.xlsx'));
  const wd = diffWorkbooks(oldWb, newWb);
  const edits: { kind: string; sheet: string; newAddress?: string }[] = m.edits;
  const editAddr = edits.map((e) => (e.newAddress ? `${e.sheet}!${e.newAddress}` : null));
  const movedBy = new Map<string, Set<number>>();
  const butFor = new Map<string, Set<number>>();
  const note = (map: Map<string, Set<number>>, a: string, i: number) => (map.get(a) ?? map.set(a, new Set()).get(a)!).add(i);
  edits.forEach((e: any, i) => {
    if (e.single) for (const a of changed(diffWorkbooks(oldWb, read(join(base, e.single))))) note(movedBy, a, i);
    const loo = e.loo ?? (edits.length === 1 ? 'old.xlsx' : null);
    if (loo) for (const a of changed(diffWorkbooks(read(join(base, loo)), newWb))) note(butFor, a, i);
  });
  const hasInsert = edits.some((e) => e.kind === 'insert');
  const cells: Record<string, { kind: string; causes: string[] | null }> = {};
  for (const a of changed(wd)) {
    const i = editAddr.indexOf(a);
    const kind = i >= 0 && edits[i].kind === 'value' ? 'input' : i >= 0 && edits[i].kind === 'formula' ? 'formula' : 'computed';
    if (kind !== 'computed') {
      cells[a] = { kind, causes: null };
      continue;
    }
    const bf = butFor.get(a);
    const truth = bf?.size ? bf : movedBy.get(a) ?? new Set<number>();
    // An insert has no single-edit oracle: such cells are listed but not scored.
    const scored = !(hasInsert && (truth.size === 0 || [...truth].some((j) => edits[j].kind === 'insert')));
    cells[a] = { kind, causes: scored ? [...truth].map((j) => editAddr[j]!).sort() : null };
  }
  const values = formulaValues(newWb);
  // In the stale file only value edits are applied: cells caused only by them must recalculate to LibreOffice's value.
  const valueEdit = new Set(editAddr.filter((_, i) => edits[i].kind === 'value'));
  const staleCheck = Object.entries(cells)
    .filter(([a, c]) => c.kind === 'computed' && c.causes?.length && c.causes.every((x) => valueEdit.has(x)) && a in values)
    .map(([a]) => a);
  const out = join(OUT, name);
  mkdirSync(out, { recursive: true });
  const files: Record<string, string> = { old: 'old.xlsx', new: 'new.xlsx' };
  copyFileSync(join(base, 'old.xlsx'), join(out, 'old.xlsx'));
  copyFileSync(join(base, 'new.xlsx'), join(out, 'new.xlsx'));
  if (m.stale && existsSync(join(base, m.stale)) && staleCheck.length) {
    copyFileSync(join(base, m.stale), join(out, 'stale.xlsx'));
    files.stale = 'stale.xlsx';
  }
  if (dir === 'cases' && m.uncached) {
    copyFileSync(join(base, m.uncached), join(out, 'uncached.xlsx'));
    files.uncached = 'uncached.xlsx';
  }
  const expected = {
    source: dir === 'cases' ? 'generated model' : 'Apache POI test file',
    files,
    cells,
    // LibreOffice's values: all formula cells for the uncached check, else only the stale-checked ones.
    values: files.uncached ? values : Object.fromEntries(staleCheck.map((a) => [a, values[a]])),
    staleCheck: files.stale ? staleCheck : [],
  };
  writeFileSync(join(out, 'expected.json'), JSON.stringify(expected));
  total += Object.keys(cells).length;
  console.log(`${name}: ${Object.keys(cells).length} changed cells, ${Object.values(cells).filter((c) => c.causes === null && c.kind === 'computed').length} unscored, stale ${expected.staleCheck.length}, files ${Object.keys(files).join(' ')}`);
}
console.log(`${CASES.length} cases, ${total} changed cells → ${OUT}`);
