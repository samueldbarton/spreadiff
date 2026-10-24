import { readFileSync } from 'node:fs';
import { strToU8, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { parseFile } from '../src/core';
import { parseXlsb } from '../src/core/xlsb';

const fixture = (name: string) => readFileSync(new URL(`./fixtures/poi/${name}`, import.meta.url));

// A tiny BIFF12 writer, for the records the POI files don't have.
const u8 = (n: number) => [n & 0xff];
const u16 = (n: number) => [n & 0xff, (n >> 8) & 0xff];
const u32 = (n: number) => [n & 0xff, (n >> 8) & 0xff, (n >> 16) & 0xff, (n >>> 24) & 0xff];
const f64 = (n: number) => [...new Uint8Array(new Float64Array([n]).buffer)];
const wstr = (s: string) => [...u32(s.length), ...[...s].flatMap((ch) => u16(ch.charCodeAt(0)))];
const formula = (rgce: number[], extra: number[] = []) => [...u32(rgce.length), ...rgce, ...u32(extra.length), ...extra];
function rec(type: number, body: number[] = []): number[] {
  const head = type < 0x80 ? [type] : [(type & 0x7f) | 0x80, type >> 7];
  const len: number[] = [];
  let n = body.length;
  do {
    len.push((n & 0x7f) | (n > 0x7f ? 0x80 : 0));
    n >>= 7;
  } while (n);
  return [...head, ...len, ...body];
}
const cell = (col: number) => [...u32(col), 0, 0, 0, 0];
const REL = 0xc000;

function xlsb(sheet: number[]): Uint8Array {
  const workbook = [
    ...rec(0x99, [...u32(0), ...u32(0)]),
    ...rec(0x9c, [...u32(0), ...u32(1), ...wstr('rId1'), ...wstr('Data')]),
    ...rec(0x9c, [...u32(0), ...u32(2), ...wstr('rId2'), ...wstr('Rates')]),
    ...rec(0x165),
    ...rec(0x16a, [...u32(2), ...u32(0), ...u32(0), ...u32(0), ...u32(0), ...u32(1), ...u32(1)]),
    // Rate = Rates!$A$1 (workbook scope), Local = 7 (on Data)
    ...rec(0x27, [...u32(0), 0, ...u32(0xffffffff), ...wstr('Rate'), ...formula([0x3a, ...u16(1), ...u32(0), ...u16(0)])]),
    ...rec(0x27, [...u32(0), 0, ...u32(0), ...wstr('Local'), ...formula([0x1e, ...u16(7)])]),
  ];
  const rels = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.bin"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.bin"/></Relationships>`;
  const rates = [...rec(0x00, [...u32(0), ...u32(0), ...u16(300), 0, 0]), ...rec(0x05, [...cell(0), ...f64(0.5)])];
  return zipSync({
    'xl/workbook.bin': new Uint8Array(workbook),
    'xl/_rels/workbook.bin.rels': strToU8(rels),
    'xl/worksheets/sheet1.bin': new Uint8Array(sheet),
    'xl/worksheets/sheet2.bin': new Uint8Array(rates),
  });
}
const row = (r: number, hidden = false) => rec(0x00, [...u32(r), ...u32(0), ...u16(300), 0, hidden ? 0x10 : 0]);

describe('.xlsb files', () => {
  it('reads values and shared formulas', () => {
    const wb = parseFile('bug66682.xlsb', fixture('bug66682.xlsb'), { header: false });
    const [t] = wb.sheets;
    expect(t.name).toBe('test2');
    expect(t.rows.slice(0, 6).map((r) => r[1])).toEqual(['Literal', true, false, 1, 1.5, 'abcd']);
    expect(t.formulas?.slice(1, 6).map((r) => r?.[2])).toEqual(['=B2', '=B3', '=B4', '=B5', '=B6']);
    expect(t.rows.slice(1, 6).map((r) => r[2])).toEqual([true, false, 1, 1.5, 'abcd']);
  });

  it('reads dates from number formats', () => {
    const wb = parseFile('date.xlsb', fixture('date.xlsb'), { header: false });
    expect(wb.sheets[0].rows[0][0]).toBe('2013-01-12');
  });

  it('reads names, 3-D references, array constants, array formulas, text and filters', () => {
    const sheet = [
      ...row(0),
      // A1 = Rate*2
      ...rec(0x09, [...cell(0), ...f64(1), ...u16(0), ...formula([0x23, ...u32(1), 0x1e, ...u16(2), 0x05])]),
      // B1 = SUM({1,2;3,4})
      ...rec(0x09, [...cell(1), ...f64(10), ...u16(0), ...formula([0x60, ...new Array(14).fill(0), 0x22, 1, ...u16(4)], [...u32(2), ...u32(2), ...[1, 2, 3, 4].flatMap((n) => [0, ...f64(n)])])]),
      // C1:C2 = {A1:A2*2}, an array formula
      ...rec(0x09, [...cell(2), ...f64(2), ...u16(0), ...formula([0x01, ...u32(0)])]),
      ...rec(0x1aa, [...u32(0), ...u32(1), ...u32(2), ...u32(2), 0, ...formula([0x25, ...u32(0), ...u32(1), ...u16(REL), ...u16(REL), 0x1e, ...u16(2), 0x05])]),
      // D1 text, E1:F1 merged
      ...rec(0x06, [...cell(3), ...wstr('hi')]),
      ...row(1),
      ...rec(0x05, [...cell(0), ...f64(3)]),
      ...rec(0x09, [...cell(2), ...f64(6), ...u16(0), ...formula([0x01, ...u32(0)])]),
      // D2 = "a"&D1
      ...rec(0x08, [...cell(3), ...wstr('ahi'), ...u16(0), ...formula([0x17, ...u16(1), ...u16(0x61), 0x24, ...u32(0), ...u16(3 | REL), 0x08])]),
      ...row(2, true),
      ...rec(0x05, [...cell(0), ...f64(4)]),
      ...rec(0xb0, [...u32(0), ...u32(0), ...u32(4), ...u32(5)]),
      ...rec(0xa1, [...u32(0), ...u32(4), ...u32(0), ...u32(5)]),
      ...rec(0xa3, [...u32(0), ...u16(0)]),
    ];
    const wb = parseXlsb(xlsb(sheet), { header: false });
    expect(wb.names).toEqual({ Rate: 'Rates!$A$1' });
    const [data, rates] = wb.sheets;
    expect(data.names).toEqual({ Local: '7' });
    expect(data.formulas?.[0]?.slice(0, 3)).toEqual(['=Rate*2', '=SUM({1,2;3,4})', '=A1:A2*2']);
    expect(data.formulas?.[1]?.[2]).toBe('=A1:A2*2');
    expect(data.formulas?.[1]?.[3]).toBe('="a"&D1');
    expect(data.rows[0][3]).toBe('hi');
    expect(data.rows[1][3]).toBe('ahi');
    expect(data.hiddenRows).toEqual({ rows: [2], filtered: [2] });
    expect(rates.rows).toEqual([[0.5]]);
  });

  it('says so when a workbook is password-protected', () => {
    expect(() => parseFile('protected_passtika.xlsb', fixture('protected_passtika.xlsb'))).toThrow(/password-protected/);
  });

  it('is routed by content, not by name', () => {
    expect(() => parseFile('book.xlsb', new Uint8Array(8))).toThrow(/not a valid \.xlsb file/);
    expect(parseFile('upload.zip', fixture('date.xlsb'), { header: false }).sheets[0].rows[0][0]).toBe('2013-01-12');
  });
});
