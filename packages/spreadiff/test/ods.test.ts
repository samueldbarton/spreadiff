import { strToU8, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { parseFile } from '../src/core';
import { openFormulaToExcel, parseOds } from '../src/core/ods';

/** A minimal .ods around one table's rows. */
function ods(rows: string, extra = ''): Uint8Array {
  const content = `<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0">
<office:body><office:spreadsheet>${extra}<table:table table:name="Data">${rows}</table:table></office:spreadsheet></office:body></office:document-content>`;
  return zipSync({ mimetype: [strToU8('application/vnd.oasis.opendocument.spreadsheet'), { level: 0 }], 'content.xml': strToU8(content) });
}
const cell = (attrs: string, text = '') => `<table:table-cell ${attrs}>${text ? `<text:p>${text}</text:p>` : ''}</table:table-cell>`;

describe('OpenDocument spreadsheets', () => {
  it('converts OpenFormula to Excel syntax', () => {
    expect(openFormulaToExcel('of:=[.A1]+[.B2]*2')).toBe('=A1+B2*2');
    expect(openFormulaToExcel('of:=SUM([.A1:.A9];[$Other.$B$2])')).toBe('=SUM(A1:A9,Other!$B$2)');
    expect(openFormulaToExcel("of:=['My sheet'.A1]&\"a;b\"")).toBe("='My sheet'!A1&\"a;b\"");
    expect(openFormulaToExcel('of:=SUM([$Jan.A1:$Mar.A1])')).toBe('=SUM(Jan:Mar!A1:A1)'); // as LibreOffice's own .xlsx export writes it
    expect(openFormulaToExcel('of:=COM.MICROSOFT.CEILING([.A1];1)')).toBe('=CEILING(A1,1)');
    expect(openFormulaToExcel('of:=TRANSPOSE({1;2|3;4})')).toBe('=TRANSPOSE({1,2;3,4})');
  });

  it('reads values, formulas, repeats, merges, hidden rows; not comments or shapes', () => {
    const rows = [
      `<table:table-row>${cell('office:value-type="string"', 'item')}${cell('office:value-type="string"', 'value')}</table:table-row>`,
      `<table:table-row>${cell('office:value-type="string"', 'a')}${cell('office:value-type="float" office:value="2.5"', '2.5')}</table:table-row>`,
      `<table:table-row table:visibility="filter">${cell('office:value-type="string"', 'b')}${cell('office:value-type="float" office:value="5" table:formula="of:=[.B2]*2"', '5')}</table:table-row>`,
      `<table:table-row table:visibility="collapse">${cell('office:value-type="date" office:date-value="2024-01-15"', '01/15/24')}${cell('office:value-type="boolean" office:boolean-value="true"', 'TRUE')}</table:table-row>`,
      `<table:table-row>${cell('office:value-type="string" table:number-columns-spanned="2"', 'merged<office:annotation><text:p>a note</text:p></office:annotation>')}<table:covered-table-cell/></table:table-row>`,
      `<table:table-row>${cell('office:value-type="time" office:time-value="PT30H15M00S"', '30:15')}${cell('calcext:value-type="error" table:formula="of:=1/0"', 'Err:532')}</table:table-row>`,
      `<table:table-row>${cell('office:value-type="date" office:date-value="1900-01-01"', '')}<table:table-cell><draw:custom-shape><text:p>a box</text:p></draw:custom-shape></table:table-cell></table:table-row>`,
      `<table:table-row table:number-rows-repeated="1048000"><table:table-cell table:number-columns-repeated="1024"/></table:table-row>`,
    ].join('');
    const wb = parseFile('book.ods', ods(rows));
    const t = wb.sheets[0];
    expect(t.columns).toEqual(['item', 'value']);
    expect(t.rows).toEqual([
      ['a', 2.5],
      ['b', 5],
      ['2024-01-15', true],
      ['merged', null],
      ['30:15:00', '#DIV/0!'],
      // LibreOffice's 1900-01-01 is its serial 2, which its .xlsx export writes and Excel shows as 1900-01-02
      // (Excel counts a fictional 1900-02-29). The same workbook as .ods and as .xlsx then diffs equal.
      ['1900-01-02', null],
    ]);
    expect(t.formulas?.[1]?.[1]).toBe('=B2*2');
    expect(t.formulas?.[4]?.[1]).toBe('=1/0');
    expect(t.merges).toEqual([{ r0: 4, c0: 0, r1: 4, c1: 1 }]);
    expect(t.hiddenRows).toEqual({ rows: [2, 3], filtered: [2] });
  });

  it('is chosen by name or by its mimetype', () => {
    const bytes = ods(`<table:table-row>${cell('office:value-type="float" office:value="1"', '1')}</table:table-row>`);
    expect(parseFile('noext', bytes, { header: false }).sheets[0].rows).toEqual([[1]]);
    expect(parseOds(bytes, { header: false }).name).toBe('workbook.ods');
  });
});
