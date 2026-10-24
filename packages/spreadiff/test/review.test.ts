import { describe, expect, it } from 'vitest';
import { strToU8, zipSync } from 'fflate';
import { diffTables, parseCsv, parseFile, parseXlsx, toUnifiedCsv, uniqueNames, formatCell } from '../src/core';
import type { Table } from '../src/core';

const csv = (s: string) => parseCsv(s.trim() + '\n');

function xlsx(sheetXml: string, extra: Record<string, string> = {}) {
  return zipSync({
    'xl/workbook.xml': strToU8(extra.workbook ?? '<workbook><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>'),
    'xl/_rels/workbook.xml.rels': strToU8('<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>'),
    'xl/sharedStrings.xml': strToU8(extra.sst ?? '<sst></sst>'),
    'xl/styles.xml': strToU8(extra.styles ?? '<styleSheet><cellXfs><xf numFmtId="0"/><xf numFmtId="14"/></cellXfs></styleSheet>'),
    'xl/worksheets/sheet1.xml': strToU8(sheetXml),
  });
}
const ws = (rows: string) => `<worksheet><sheetData>${rows}</sheetData></worksheet>`;

describe('review: core', () => {
  it('uniqueNames never produces duplicates', () => {
    const out = uniqueNames(['x', 'x', 'x (2)']);
    expect(new Set(out).size).toBe(3);
  });

  it('does not call sparse, unrelated columns a rename', () => {
    const rows = (f: (i: number) => string) => Array.from({ length: 40 }, (_, i) => f(i)).join('\n');
    const a = csv(`id,notes\n${rows((i) => `${i},${i === 3 ? 'urgent' : ''}`)}`);
    const b = csv(`id,discount\n${rows((i) => `${i},${i === 9 ? '0.5' : ''}`)}`);
    const d = diffTables(a, b);
    expect(d.columns.map((c) => c.status)).toEqual(['equal', 'removed', 'added']);
  });

  it('auto-detects ID / Order ID / customerId keys', () => {
    for (const name of ['ID', 'Order ID', 'customerId', 'order_id']) {
      const d = diffTables(csv(`${name},v\n1,a\n2,b\n3,c`), csv(`${name},v\n3,c\n1,a\n2,z`));
      expect(d.keyColumns, name).toEqual([name]);
    }
    // ...but not words that merely end in "id".
    const d = diffTables(csv(`paid,v\n1,a\n2,b`), csv(`paid,v\n2,b\n1,a`));
    expect(d.keyColumns).toEqual([]);
  });

  it('handles duplicate column names in hand-built tables', () => {
    const t = (rows: (string | null)[][]): Table => ({ name: 's', columns: ['x', 'x'], rows });
    const d = diffTables(t([['1', '2']]), t([['1', '3']]));
    expect(d.columns.map((c) => [c.oldIndex, c.newIndex])).toEqual([[0, 0], [1, 1]]);
    expect(d.stats.cellsChanged).toBe(1);
  });

  it('formats numbers without hiding real differences', () => {
    expect(formatCell(0.1 + 0.2)).toBe('0.3');
    expect(formatCell(1.2345678901234)).toBe('1.2345678901234');
  });

  it('sanitizes exported cells that spreadsheets would execute', () => {
    const d = diffTables(csv('id,v\n1,a'), csv('id,v\n1,=HYPERLINK("x")'));
    const out = toUnifiedCsv(d);
    expect(out).toContain(`"'a->=HYPERLINK(""x"")"`.replace(`'a`, `a`)); // cell starting with 'a' is safe
    const d2 = diffTables(csv('id,v\n1,ok'), csv('id,v\n2,@SUM(A1)'));
    expect(toUnifiedCsv(d2)).toContain("'@SUM(A1)");
    const d3 = diffTables(csv('id,v\n1,5'), csv('id,v\n1,-5'));
    expect(toUnifiedCsv(d3)).toContain('5->-5');
  });
});

describe('review: xlsx', () => {
  it('handles self-closing <t/> in rich text without swallowing runs', () => {
    const wb = parseXlsx(xlsx(ws('<row r="1"><c r="A1" t="s"><v>0</v></c></row><row r="2"><c r="A2" t="s"><v>1</v></c></row>'), {
      sst: '<sst><si><t>h</t></si><si><r><t/></r><r><t>abc</t></r></si></sst>',
    }));
    expect(wb.sheets[0].rows).toEqual([['abc']]);
  });

  it('supports namespace prefixes, single quotes and newline-separated attributes', () => {
    const sheet = `<x:worksheet xmlns:x="ns"><x:sheetData><x:row r='1'><x:c r='A1' t='inlineStr'><x:is><x:t>h</x:t></x:is></x:c></x:row><x:row r="2"><x:c\n r="A2"><x:v>7</x:v></x:c></x:row></x:sheetData></x:worksheet>`;
    const wb = parseXlsx(xlsx(sheet, { workbook: `<x:workbook><x:sheets><x:sheet name='S' sheetId='1' r:id='rId1'/></x:sheets></x:workbook>` }));
    expect(wb.sheets[0].columns).toEqual(['h']);
    expect(wb.sheets[0].rows).toEqual([[7]]);
  });

  it('does not throw on out-of-range date serials, and handles pre-1900-03 dates', () => {
    const wb = parseXlsx(xlsx(ws('<row r="1"><c r="A1" t="inlineStr"><is><t>d</t></is></c></row><row r="2"><c r="A2" s="1"><v>1e12</v></c></row><row r="3"><c r="A3" s="1"><v>1</v></c></row><row r="4"><c r="A4" s="1"><v>61</v></c></row>')));
    expect(wb.sheets[0].rows).toEqual([[1e12], ['1900-01-01'], ['1900-03-01']]);
  });

  it('shifts cross-sheet references in shared formulas', () => {
    const wb = parseXlsx(xlsx(ws('<row r="1"><c r="A1" t="inlineStr"><is><t>f</t></is></c></row><row r="2"><c r="A2"><f t="shared" ref="A2:A3" si="0">Data!B2+1</f><v>1</v></c></row><row r="3"><c r="A3"><f t="shared" si="0"/><v>2</v></c></row>')));
    expect(wb.sheets[0].formulas![1][0]).toBe('=Data!B3+1');
  });

  it('always records formulas so formula → value changes are detected', () => {
    const withF = parseXlsx(xlsx(ws('<row r="1"><c r="A1" t="inlineStr"><is><t>v</t></is></c></row><row r="2"><c r="A2"><f>1+1</f><v>2</v></c></row>')));
    const noF = parseXlsx(xlsx(ws('<row r="1"><c r="A1" t="inlineStr"><is><t>v</t></is></c></row><row r="2"><c r="A2"><v>2</v></c></row>')));
    expect(diffTables(withF.sheets[0], noF.sheets[0]).stats.cellsChanged).toBe(1);
  });

  it('stays linear on malformed input (unclosed cells)', () => {
    const bad = ws('<row r="1">' + '<c r="A1"><v>1</v>'.repeat(60_000) + '</row>');
    const t0 = performance.now();
    try { parseXlsx(xlsx(bad)); } catch {}
    expect(performance.now() - t0).toBeLessThan(2000);
  });

  it('records where the data starts for spreadsheet row/column labels', () => {
    const wb = parseXlsx(xlsx(ws('<row r="3"><c r="C3" t="inlineStr"><is><t>h</t></is></c></row><row r="4"><c r="C4"><v>1</v></c></row>')));
    expect(wb.sheets[0].origin).toEqual({ row: 3, col: 2 });
  });
});

describe('review: parseFile', () => {
  it('decodes UTF-16 text exports (Excel "Unicode Text")', () => {
    const text = '﻿a\tb\r\n1\tü\r\n';
    const buf = new Uint8Array(text.length * 2);
    for (let i = 0; i < text.length; i++) { buf[2 * i] = text.charCodeAt(i) & 0xff; buf[2 * i + 1] = text.charCodeAt(i) >> 8; }
    const wb = parseFile('export.txt', buf);
    expect(wb.sheets[0].columns).toEqual(['a', 'b']);
    expect(wb.sheets[0].rows).toEqual([[1, 'ü']]);
  });
  it('rejects legacy .xls with a clear error instead of parsing binary as CSV', () => {
    const ole = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, 0, 0]);
    expect(() => parseFile('old.xls', ole)).toThrow(/xls/i);
  });
});
