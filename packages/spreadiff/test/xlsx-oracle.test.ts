import { describe, expect, it } from 'vitest';
import ExcelJS from 'exceljs';
import { parseXlsx } from '../src/core';

function rng(seed: number) {
  return () => {
    seed = (seed + 0x6d2b79f5) | 0;
    let t = Math.imul(seed ^ (seed >>> 15), 1 | seed);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

const STRINGS = ['plain', 'a & b', '<tag>', '"quoted"', "it's", ' padded ', 'tab\there', 'multi\nline', 'é', '日本語', '😀 emoji', '', '0042', '1e5', 'TRUE', '=not a formula', ']]>', '&amp;'];

describe('xlsx: parse files written by ExcelJS (independent writer)', () => {
  it('round-trips values, dates, booleans, formulas and sheet names', async () => {
    const r = rng(42);
    for (let iter = 0; iter < 40; iter++) {
      const wb = new ExcelJS.Workbook();
      const expected: Record<string, unknown[][]> = {};
      const expectedF: Record<string, (string | null)[][]> = {};
      const nSheets = 1 + Math.floor(r() * 3);
      for (let s = 0; s < nSheets; s++) {
        const name = ['Data', 'Q&A', 'Sales <2026>', 'Ünïcode', "O'Brien"][s + (iter % 3)] ?? `S${s}`;
        const ws = wb.addWorksheet(name);
        const w = 1 + Math.floor(r() * 4);
        const h = Math.floor(r() * 8);
        ws.addRow(Array.from({ length: w }, (_, c) => `col${c}`));
        const rows: unknown[][] = [];
        const forms: (string | null)[][] = [];
        for (let i = 0; i < h; i++) {
          const row: unknown[] = [];
          const frow: (string | null)[] = [];
          for (let c = 0; c < w; c++) {
            const k = r();
            if (k < 0.35) {
              const v = STRINGS[Math.floor(r() * STRINGS.length)];
              row.push(v === '' ? null : v); // ExcelJS drops empty strings
              frow.push(null);
            } else if (k < 0.6) {
              const v = Math.round((r() - 0.5) * 1e6) / 100;
              row.push(v);
              frow.push(null);
            } else if (k < 0.7) {
              row.push(r() < 0.5);
              frow.push(null);
            } else if (k < 0.8) {
              const day = 1 + Math.floor(r() * 27);
              row.push(new Date(Date.UTC(2020 + Math.floor(r() * 6), Math.floor(r() * 12), day)));
              frow.push(null);
            } else if (k < 0.9) {
              row.push({ formula: 'A1&"x"', result: 'col0x' });
              frow.push('=A1&"x"');
            } else {
              row.push(null);
              frow.push(null);
            }
          }
          rows.push(row);
          forms.push(frow);
        }
        rows.forEach((row) => {
          const added = ws.addRow(row.map((v) => (v === null ? undefined : v)));
          added.eachCell((cell) => {
            if (cell.value instanceof Date) cell.numFmt = 'yyyy-mm-dd';
          });
        });
        // Expected values as our parser should report them.
        expected[name] = rows.map((row) =>
          row.map((v) => (v instanceof Date ? v.toISOString().slice(0, 10) : v && typeof v === 'object' ? 'col0x' : v)),
        );
        expectedF[name] = forms;
      }
      const buf = await wb.xlsx.writeBuffer();
      // The formulas read A1, so 'auto' would (rightly) treat row 1 as data; this test is about values.
      const parsed = parseXlsx(new Uint8Array(buf as ArrayBuffer), { header: true });
      for (const sheet of parsed.sheets) {
        const exp = expected[sheet.name];
        expect(exp, `unexpected sheet ${sheet.name}`).toBeDefined();
        // Trailing all-empty rows are trimmed by design.
        const trimmed = [...exp];
        const tf = [...expectedF[sheet.name]];
        while (trimmed.length && trimmed[trimmed.length - 1].every((v) => v === null)) { trimmed.pop(); tf.pop(); }
        const width = sheet.columns.length;
        const got = sheet.rows.map((row) => row.slice(0, width));
        const want = trimmed.map((row) => [...row, ...Array(Math.max(0, width - row.length)).fill(null)].slice(0, width));
        expect(got, `iter ${iter} sheet ${sheet.name}`).toEqual(want);
        tf.forEach((frow, i) => frow.forEach((f, c) => {
          if (f) expect(sheet.formulas?.[i]?.[c], `formula iter ${iter} ${sheet.name} r${i} c${c}`).toBe(f);
        }));
      }
      expect(parsed.sheets.map((s) => s.name).sort()).toEqual(Object.keys(expected).sort());
    }
  });
});
