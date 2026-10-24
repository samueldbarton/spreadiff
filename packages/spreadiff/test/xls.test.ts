import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { parseFile } from '../src/core';
import { readCfb } from '../src/core/cfb';
import { decodeFormula, type PtgContext } from '../src/core/ptg';

const fixture = (name: string) => readFileSync(new URL(`./fixtures/poi/${name}`, import.meta.url));
const ctx = (over: Partial<PtgContext> = {}): PtgContext => ({
  flavor: 'biff8',
  row: 0,
  col: 0,
  sheetOf: (i) => [{ first: 'Data', last: 'Data' }, { first: 'Jan', last: 'Mar' }, { first: 'S2', last: 'S2' }][i] ?? null,
  nameOf: (i) => (i === 1 ? 'Rate' : null),
  nameXOf: () => null,
  ...over,
});
const bytes = (...parts: (number | number[])[]) => new Uint8Array(parts.flat());
const u16 = (n: number) => [n & 0xff, n >> 8];
const u32 = (n: number) => [n & 0xff, (n >> 8) & 0xff, (n >> 16) & 0xff, n >>> 24];
const REL = 0xc000; // column field flags: row and column relative

describe('formula tokens', () => {
  it('decodes references, operators and functions (BIFF8)', () => {
    // A1+2
    expect(decodeFormula(bytes(0x24, u16(0), u16(REL), 0x1e, u16(2), 0x03), ctx())).toBe('=A1+2');
    // SUM($B$2:C10)
    expect(decodeFormula(bytes(0x25, u16(1), u16(9), u16(1), u16(2 | REL), 0x22, 1, u16(4)), ctx())).toBe('=SUM($B$2:C10)');
    // "a"&Rate
    expect(decodeFormula(bytes(0x17, 1, 0, 0x61, 0x23, u16(1), u16(0), 0x08), ctx())).toBe('="a"&Rate');
  });

  it('decodes 3-D references, quoting sheet names that look like cells', () => {
    expect(decodeFormula(bytes(0x3a, u16(0), u16(0), u16(REL)), ctx())).toBe('=Data!A1');
    expect(decodeFormula(bytes(0x3a, u16(1), u16(4), u16(1)), ctx())).toBe('=Jan:Mar!$B$5');
    expect(decodeFormula(bytes(0x3a, u16(2), u16(0), u16(REL)), ctx())).toBe("='S2'!A1");
  });

  it('decodes BIFF12 references with wider rows', () => {
    expect(decodeFormula(bytes(0x24, u32(1_000_000), u16(REL)), ctx({ flavor: 'biff12' }))).toBe('=A1000001');
  });

  it('gives up on tokens it does not know rather than guessing', () => {
    expect(decodeFormula(bytes(0x24, u16(0), u16(REL), 0x7f), ctx())).toBeNull();
    expect(decodeFormula(bytes(0x24, u16(0)), ctx())).toBeNull(); // truncated
    expect(decodeFormula(bytes(0x03), ctx())).toBeNull(); // operator with no operands
  });
});

describe('.xls files', () => {
  it('reads values, formulas and 3-D references', () => {
    const wb = parseFile('3dFormulas.xls', fixture('3dFormulas.xls'), { header: false });
    expect(wb.sheets.map((s) => s.name)).toEqual(['Sheet1', 'S2', 'Sh3']);
    const [s1, s2, s3] = wb.sheets;
    expect(s1.rows[1]).toEqual([11, 22, 33]);
    expect(s1.formulas?.[1]).toEqual(['=Sheet1!A1', '=Sheet1!B1', '=Sheet1!C1']);
    expect(s1.formulas?.[4]?.slice(0, 2)).toEqual(["='S2'!A1", "='S2'!B1"]);
    expect(s1.rows[4].slice(0, 2)).toEqual(['S2-A1', 'S2-A2']);
    expect(s2.rows).toEqual([['S2-A1', 'S2-A2']]);
    expect(s3.rows).toEqual([[44], [55], [66]]);
  });

  it('reads dates in the 1904 date system', () => {
    const wb = parseFile('1904DateWindowing.xls', fixture('1904DateWindowing.xls'), { header: false });
    expect(wb.sheets[0].rows[0]).toEqual(['2000-01-01', '2000-01-01']);
  });

  it('is routed by content, not by name', () => {
    const wb = parseFile('upload.bin', fixture('3dFormulas.xls'), { header: false });
    expect(wb.sheets).toHaveLength(3);
  });

  it('rejects a damaged compound file', () => {
    const data = new Uint8Array(fixture('3dFormulas.xls'));
    expect(() => readCfb(data.slice(0, 600))).toThrow();
  });
});
