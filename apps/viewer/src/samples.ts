import type { CellValue, Table, Workbook } from 'spreadiff';

function rng(seed: number) {
  return () => {
    seed = (seed + 0x6d2b79f5) | 0;
    let t = Math.imul(seed ^ (seed >>> 15), 1 | seed);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

const ADJ = ['Quiet', 'Bright', 'Nordic', 'Linen', 'Walnut', 'Copper', 'Slate', 'Cedar', 'Ivory', 'Smoked', 'Matte', 'Woven'];
const NOUN = ['Lamp', 'Chair', 'Desk', 'Mug', 'Throw', 'Shelf', 'Vase', 'Stool', 'Rug', 'Clock', 'Planter', 'Tray'];
const CATS = ['Lighting', 'Furniture', 'Kitchen', 'Textiles', 'Decor', 'Storage'];
const WAREHOUSES = ['AMS-1', 'BER-2', 'NYC-3', 'SFO-1', 'TYO-2'];

/** A small, realistic two-sheet workbook pair that exercises every diff feature. */
export function sampleWorkbooks(): { old: Workbook; new: Workbook } {
  const r = rng(11);
  const pick = <T,>(xs: T[]) => xs[Math.floor(r() * xs.length)];
  const rows: CellValue[][] = [];
  for (let i = 0; i < 420; i++) {
    const name = `${pick(ADJ)} ${pick(NOUN)}`;
    rows.push([
      `P-${String(1000 + i)}`,
      name,
      pick(CATS),
      Math.round((5 + r() * 240) * 100) / 100,
      Math.floor(r() * 500),
      pick(WAREHOUSES),
      `2026-0${1 + Math.floor(r() * 8)}-${String(1 + Math.floor(r() * 27)).padStart(2, '0')}`,
      r() < 0.15 ? 'Seasonal' : null,
    ]);
  }
  const oldInv: Table = {
    name: 'Inventory',
    columns: ['id', 'name', 'category', 'price', 'qty', 'warehouse', 'restocked', 'notes'],
    rows,
  };
  // New version: price changes, restocks, deletions, insertions, a renamed
  // column (qty → on_hand), a dropped column (notes), a new column (margin).
  const newRows: CellValue[][] = [];
  rows.forEach((row, i) => {
    if (i % 53 === 7) return; // discontinued
    const [id, name, cat, price, qty, wh, restocked] = row as [string, string, string, number, number, string, string];
    let p = price;
    let q = qty;
    let n = name;
    let w = wh;
    let d = restocked;
    if (i % 17 === 3) p = Math.round(price * 1.08 * 100) / 100;
    if (i % 29 === 5) {
      q = qty + 120;
      d = '2026-10-14';
    }
    if (i % 61 === 9) n = `${name} (v2)`;
    if (i % 71 === 13) w = pick(WAREHOUSES);
    newRows.push([id, n, cat, p, q, w, d, Math.round((0.18 + r() * 0.4) * 1000) / 1000]);
    if (i % 97 === 50) {
      newRows.push([`P-${2000 + i}`, `${pick(ADJ)} ${pick(NOUN)}`, pick(CATS), Math.round(r() * 20000) / 100, 40, pick(WAREHOUSES), '2026-10-01', 0.3]);
    }
  });
  const newInv: Table = {
    name: 'Inventory',
    columns: ['id', 'name', 'category', 'price', 'on_hand', 'warehouse', 'restocked', 'margin'],
    rows: newRows,
  };
  const regions = (bump: boolean): Table => ({
    name: 'Regions',
    columns: ['region', 'manager', 'target', 'actual'],
    rows: [
      ['EMEA', 'Ada Okafor', 1200000, bump ? 1312400 : 1188000],
      ['NA', 'Sam Lee', 2100000, 2066500],
      ['APAC', 'Kenji Mori', 900000, bump ? 948250 : 871900],
      ...(bump ? [['LATAM', 'Lucía Pérez', 450000, 120300] as CellValue[]] : []),
    ],
  });
  const summary: Table = {
    name: 'Notes',
    columns: ['key', 'value'],
    rows: [['owner', 'ops@studio.example'], ['cadence', 'weekly']],
  };
  return {
    old: { name: 'inventory-2026-09.xlsx', sheets: [oldInv, regions(false), summary] },
    new: { name: 'inventory-2026-10.xlsx', sheets: [newInv, regions(true), summary] },
  };
}

/** Synthetic large CSV pair for stress tests (generated inside the worker). */
export function syntheticCsv(rows: number, cols: number, seed = 1): { old: string; new: string; changes: number } {
  const r = rng(seed);
  const header = ['id', ...Array.from({ length: cols - 1 }, (_, c) => `metric_${c + 1}`)].join(',');
  const a: string[] = [header];
  const b: string[] = [header];
  let changes = 0;
  for (let i = 0; i < rows; i++) {
    const vals = [String(i + 1)];
    for (let c = 1; c < cols; c++) vals.push(String(Math.floor(r() * 100000) / 100));
    const line = vals.join(',');
    a.push(line);
    const roll = r();
    if (roll < 0.002) {
      changes++;
      continue; // deleted
    }
    if (roll < 0.006) {
      const c = 1 + Math.floor(r() * (cols - 1));
      vals[c] = String(Math.floor(r() * 100000) / 100);
      b.push(vals.join(','));
      changes++;
    } else b.push(line);
    if (roll > 0.998) {
      b.push([`n${i}`, ...vals.slice(1)].join(','));
      changes++;
    }
  }
  return { old: a.join('\n'), new: b.join('\n'), changes };
}

/**
 * A small financial model with real formula dependencies, for tracing:
 *   Inputs:  Price, COGS %, Tax rate (named TaxRate)
 *   Revenue: units × price → revenue, × COGS % → cogs, → gross
 *   Summary: totals, tax, net, margin
 * The new version raises the price, revises EMEA units, gives APAC its own
 * COGS formula, adds LATAM and raises the tax rate. Everything else that
 * changes is downstream.
 */
export function sampleModel(): { old: Workbook; new: Workbook } {
  type V = { price: number; cogs: number; tax: number; regions: [string, number, number | null][] };
  const build = (v: V, name: string): Workbook => {
    const inputs: Table = {
      name: 'Inputs', columns: ['assumption', 'value', 'note'], origin: { row: 1, col: 0 },
      rows: [['Price', v.price, 'per unit'], ['COGS %', v.cogs, 'of revenue'], ['Tax rate', v.tax, 'named TaxRate']],
      formulas: [[null, null, null], [null, null, null], [null, null, null]],
    };
    const rev: Table = { name: 'Revenue', columns: ['region', 'units', 'price', 'revenue', 'cogs', 'gross'], origin: { row: 1, col: 0 }, rows: [], formulas: [] };
    v.regions.forEach(([region, units, ownCogs], i) => {
      const r = i + 2;
      const revenue = units * v.price;
      const cogs = revenue * (ownCogs ?? v.cogs);
      rev.rows.push([region, units, v.price, revenue, round(cogs), round(revenue - cogs)]);
      rev.formulas!.push([null, null, '=Inputs!$B$2', `=B${r}*C${r}`, ownCogs === null ? `=D${r}*Inputs!$B$3` : `=D${r}*${ownCogs}`, `=D${r}-E${r}`]);
    });
    const last = v.regions.length + 1;
    const totRev = rev.rows.reduce((s, r) => s + (r[3] as number), 0);
    const totCogs = round(rev.rows.reduce((s, r) => s + (r[4] as number), 0));
    const gross = round(totRev - totCogs);
    const tax = round(gross * v.tax);
    const net = round(gross - tax);
    const summary: Table = {
      name: 'Summary', columns: ['line', 'value'], origin: { row: 1, col: 0 },
      rows: [['Revenue', totRev], ['COGS', totCogs], ['Gross profit', gross], ['Tax', tax], ['Net income', net], ['Net margin', Math.round((net / totRev) * 10000) / 10000]],
      formulas: [[null, `=SUM(Revenue!D2:D${last})`], [null, `=SUM(Revenue!E2:E${last})`], [null, '=B2-B3'], [null, '=B4*TaxRate'], [null, '=B4-B5'], [null, '=B6/B2']],
    };
    return { name, sheets: [summary, rev, inputs], names: { TaxRate: 'Inputs!$B$4' } };
  };
  const round = (x: number) => Math.round(x * 100) / 100;
  const regions: [string, number, number | null][] = [['NA', 2400, null], ['EMEA', 1200, null], ['APAC', 900, null], ['ANZ', 300, null]];
  return {
    old: build({ price: 49, cogs: 0.38, tax: 0.21, regions }, 'model-v1.xlsx'),
    new: build({ price: 52, cogs: 0.38, tax: 0.25, regions: [['NA', 2400, null], ['EMEA', 1350, null], ['APAC', 900, 0.42], ['ANZ', 300, null], ['LATAM', 450, null]] }, 'model-v2.xlsx'),
  };
}
