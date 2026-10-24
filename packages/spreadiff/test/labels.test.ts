import { describe, expect, it } from 'vitest';
import { formatLabel, parseCsv, tableLabeler, type Table } from '../src/core';

/** A model sheet from CSV, without a header row, as a financial model is read. */
const model = (csv: string) => parseCsv(csv, { header: false });
const at = (t: Table, row: number, col: number) => formatLabel(tableLabeler(t)(row, col));

describe('cell labels', () => {
  it('qualifies a row by the section heading above it', () => {
    const t = model([',FY2024,FY2025', 'India,,', 'Revenue,100,120', 'Costs,60,70', 'China,,', 'Revenue,80,95', 'Costs,50,55'].join('\n'));
    expect(at(t, 2, 2)).toBe('India › Revenue · FY2025');
    expect(at(t, 5, 1)).toBe('China › Revenue · FY2024');
    expect(at(t, 6, 2)).toBe('China › Costs · FY2025');
  });

  it('nests a heading directly under another, and ends a group at the next one', () => {
    const t = model(
      [',2024,2025', 'India,,', 'Revenue,,', 'Product A,10,11', 'Product B,20,21', 'Costs,,', 'Staff,5,6', 'China,,', 'Revenue,,', 'Product A,30,31'].join('\n'),
    );
    expect(at(t, 3, 2)).toBe('India › Revenue › Product A · 2025');
    expect(at(t, 6, 1)).toBe('India › Costs › Staff · 2024');
    expect(at(t, 9, 2)).toBe('China › Revenue › Product A · 2025');
  });

  it('closes a section at its total', () => {
    const t = model([',Q1', 'EMEA,', 'Units,5', 'Total EMEA,5', 'Price,3'].join('\n'));
    expect(at(t, 2, 1)).toBe('EMEA › Units · Q1');
    expect(at(t, 3, 1)).toBe('EMEA › Total EMEA · Q1');
    expect(at(t, 4, 1)).toBe('Price · Q1');
  });

  it('uses indentation from the file when there is some', () => {
    // Every label in column A, depth from the cell's indent: no blank-heading guesswork.
    const t = model([',2025', 'Asia,', 'India,', 'Revenue,1', 'China,', 'Revenue,2', 'Europe,', 'Revenue,3'].join('\n'));
    const o = t.origin!;
    const indent = (row: number, n: number) => ({ r: row + o.row, c: o.col, indent: n });
    t.labelStyles = [indent(2, 1), indent(3, 2), indent(4, 1), indent(5, 2), indent(7, 1)];
    expect(at(t, 3, 1)).toBe('Asia › India › Revenue · 2025');
    expect(at(t, 5, 1)).toBe('Asia › China › Revenue · 2025');
    expect(at(t, 7, 1)).toBe('Europe › Revenue · 2025');
  });

  it('reads leading spaces as indentation', () => {
    const t = model([',2025', 'Opex,', '  Rent,4', '  Staff,5', 'Capex,', '  Kit,9'].join('\n'));
    expect(at(t, 3, 1)).toBe('Opex › Staff · 2025');
    expect(at(t, 5, 1)).toBe('Capex › Kit · 2025');
  });

  it('stacks column headings, an upper one spanning to its right', () => {
    const t = model([',Actual,,Forecast,', ',Q1,Q2,Q1,Q2', 'Sales,1,2,3,4'].join('\n'));
    expect(at(t, 2, 2)).toBe('Sales · Actual › Q2');
    expect(at(t, 2, 4)).toBe('Sales · Forecast › Q2');
  });

  it('uses a merged title for every column under it', () => {
    const t = model([',Plan,,', ',2024,2025,2026', 'Units,1,2,3'].join('\n'));
    t.merges = [{ r0: t.origin!.row, c0: 1, r1: t.origin!.row, c1: 3 }];
    expect(at(t, 2, 3)).toBe('Units · Plan › 2026');
  });

  it('takes the nearest heading block when a sheet repeats them', () => {
    const t = model([',2024', 'Units,1', ',', ',Jan', 'Units,2'].join('\n'));
    expect(at(t, 1, 1)).toBe('Units · 2024');
    expect(at(t, 4, 1)).toBe('Units · Jan');
  });

  it('labels a records table by its key and first text column, and the header', () => {
    const t = parseCsv('id,name,price\nP-1,Lamp,10\nP-2,Desk,30');
    const label = tableLabeler(t, { keyColumns: ['id'] });
    expect(formatLabel(label(1, 2))).toBe('P-2 › Desk · price');
    expect(label(1, 2).sections).toEqual([]);
  });

  it("doesn't take a numbers row for headings, or years in a data row", () => {
    const t = model([',2024,2025', 'Units,2000,2010', 'Price,3.5,4'].join('\n'));
    expect(at(t, 2, 2)).toBe('Price · 2025'); // "Units" holds year-like numbers, but it's data: the heading stays 2024/2025
    const decreasing = model([',2024,2025', 'Units,2010,2000', 'Price,3.5,4'].join('\n'));
    expect(at(decreasing, 2, 1)).toBe('Price · 2024');
  });

  it('gives nothing for a sheet of numbers', () => {
    const t = model('1,2\n3,4');
    expect(tableLabeler(t)(1, 1)).toEqual({ sections: [], row: [], column: [] });
  });
});

describe('cell labels: dates', () => {
  it('takes a row of month-ends as column headings, but not a lone date', () => {
    const t = model([',2024-01-31,2024-02-29', 'Cash,5,6'].join('\n'));
    expect(at(t, 1, 2)).toBe('Cash · 2024-02-29');
    const lone = model(['2024-01-31', '5', '6'].join('\n'));
    expect(tableLabeler(lone)(2, 0).column).toEqual([]);
  });
});
