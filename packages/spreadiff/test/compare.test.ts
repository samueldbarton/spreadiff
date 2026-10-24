import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { compareWorkbooks, diffFiles, diffWorkbooks, parseCsv, parseFile, reportOf, traceWorkbook } from '../src/core';

const fixture = (name: string) => new Uint8Array(readFileSync(fileURLToPath(new URL(`../../../test/hosted/fixtures/commission/${name}`, import.meta.url))));

describe('diffFiles / compareWorkbooks', () => {
  it('reads every kind of input and gives the same result', async () => {
    const old = fixture('old.xlsx');
    const now = fixture('new.xlsx');
    const ref = compareWorkbooks(parseFile('old.xlsx', old), parseFile('new.xlsx', now));
    const inputs = [
      [{ name: 'old.xlsx', data: old }, { name: 'new.xlsx', data: now }],
      [new Blob([old]), new Blob([now])],
      [old, now], // sniffed: zip → XLSX
      [old.slice().buffer, now.slice().buffer],
    ] as const;
    for (const [a, b] of inputs) {
      const cmp = await diffFiles(a, b);
      expect(cmp.sheets.map((s) => [s.name, s.status])).toEqual(ref.sheets.map((s) => [s.name, s.status]));
      expect(cmp.trace?.stats).toEqual(ref.trace?.stats);
    }
    // The caller's bytes are left alone (a worker would take a copy, not these).
    expect(old.byteLength).toBeGreaterThan(0);
  });

  it('binds the trace to each sheet, and traces only when there are formulas', async () => {
    const cmp = await diffFiles({ name: 'a.xlsx', data: fixture('old.xlsx') }, { name: 'b.xlsx', data: fixture('new.xlsx') });
    expect(cmp.trace).not.toBeNull();
    cmp.sheets.forEach((s, i) => {
      expect(s.index).toBe(i);
      expect(s.diff).toBe(cmp.workbook.sheets[i].diff);
      expect(s.trace).toEqual({ trace: cmp.trace, workbook: cmp.workbook, sheet: i });
    });
    expect(cmp.names).toEqual({ old: 'a.xlsx', new: 'b.xlsx' });
    expect(cmp.sheets[cmp.firstChanged].status).not.toBe('equal');

    const csv = await diffFiles({ name: 'a.csv', data: 'id,x\n1,2' }, { name: 'b.csv', data: 'id,x\n1,3' });
    expect(csv.trace).toBeNull();
    expect(csv.sheets[0].trace).toBeUndefined();
    expect(csv.sheets[0].diff.stats.cellsChanged).toBe(1);
    expect((await diffFiles({ name: 'a.xlsx', data: fixture('old.xlsx') }, { name: 'b.xlsx', data: fixture('new.xlsx') }, { trace: false })).trace).toBeNull();
  });

  it('builds the report once, the same as reportOf', async () => {
    const a = parseFile('old.xlsx', fixture('old.xlsx'));
    const b = parseFile('new.xlsx', fixture('new.xlsx'));
    const cmp = compareWorkbooks(a, b);
    const r = cmp.report();
    expect(cmp.report()).toBe(r);
    const wd = diffWorkbooks(a, b);
    expect(JSON.stringify(r)).toBe(JSON.stringify(reportOf(wd, traceWorkbook(wd), 'old.xlsx', 'new.xlsx')));
    expect(r.totals.downstream).toBeGreaterThan(0);
  });

  it('accepts tables, and reports CSV changes as edits without a trace', () => {
    const cmp = compareWorkbooks(parseCsv('id,x\n1,2\n2,5', { name: 'T' }), parseCsv('id,x\n1,3\n2,5', { name: 'T' }));
    expect(cmp.sheets).toHaveLength(1);
    expect(cmp.report().edits.map((e) => [e.cell, e.before, e.after])).toEqual([['B2', 2, 3]]);
  });
});
