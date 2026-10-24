import { describe, expect, it } from 'vitest';
import { formatNumberCode, formatTextCode, generalText, isDateFormat } from '../src/core/numfmt';
import { Engine } from '../src/core/engine/evaluate';
import { tableHost } from '../src/core/engine/host';

const f = (v: number, code: string) => formatNumberCode(v, code);

describe('number formats', () => {
  it('formats the codes finance models use', () => {
    expect(f(1234.5678, '#,##0')).toBe('1,235');
    expect(f(1234.5678, '#,##0.00')).toBe('1,234.57');
    expect(f(-1234.5678, '#,##0_);(#,##0)')).toBe('(1,235)');
    expect(f(1234.5678, '#,##0_);(#,##0)')).toBe('1,235 '); // _) leaves room for the parenthesis
    expect(f(-1234.5, '$#,##0.00_);[Red]($#,##0.00)')).toBe('($1,234.50)');
    expect(f(0.0789, '0.0%')).toBe('7.9%');
    expect(f(12345678, '#,##0.0,,"m"')).toBe('12.3m');
    expect(f(12345678, '#,##0,"k"')).toBe('12,346k');
    expect(f(1.5, '0.0"x"')).toBe('1.5x');
    expect(f(0, '#,##0.00;(#,##0.00);"-"')).toBe('-');
    expect(f(1234.5, '[$€-407] #,##0.00')).toBe('€ 1,234.50');
    expect(f(1500, '[>=1000]#,##0,"k";0')).toBe('2k');
    expect(f(999, '[>=1000]#,##0,"k";0')).toBe('999');
  });

  it('formats the accounting code: spaces for _x, nothing for the * fill', () => {
    const acct = '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)';
    expect(f(1234.5, acct)).toBe(' $1,234.50 ');
    expect(f(-1234.5, acct)).toBe(' $(1,234.50)');
    expect(f(0, acct)).toBe(' $-   ');
    expect(formatTextCode('n/a', acct)).toBe(' n/a ');
  });

  it('formats scientific, engineering and fractions', () => {
    expect(f(12345678, '0.00E+00')).toBe('1.23E+07');
    expect(f(0.000123, '0.00E+00')).toBe('1.23E-04');
    expect(f(12345678, '##0.0E+0')).toBe('12.3E+6');
    expect(f(999.995, '##0.0E+0')).toBe('1.0E+3'); // rounding carries into the next step
    expect(f(1.5, '# ?/?')).toBe('1 1/2');
    expect(f(0.333, '# ??/??')).toBe('  1/3 ');
    expect(f(1.5, '?/8')).toBe('12/8');
    expect(f(2, '# ?/?')).toBe('2    ');
  });

  it('keeps Excel quirks LibreOffice smooths over', () => {
    // Excel shows a small negative rounded to zero as -0, and keeps the point of "#.##".
    expect(f(-0.004, '0.00')).toBe('-0.00');
    expect(f(5, '#.##')).toBe('5.');
    // Serial 0 is 1900-01-00, and weekdays before March 1900 run a day behind.
    expect(f(0.5, 'yyyy-mm-dd hh:mm')).toBe('1900-01-00 12:00');
    expect(f(1, 'dddd')).toBe('Sunday');
    expect(f(61, 'dddd')).toBe('Thursday'); // 1900-03-01, back on the calendar
  });

  it('formats dates and times, and refuses ones Excel shows as #####', () => {
    expect(f(45292.75, 'd-mmm-yy')).toBe('1-Jan-24');
    expect(f(45292.75, 'mmmm d, yyyy')).toBe('January 1, 2024');
    expect(f(45292.75, 'h:mm AM/PM')).toBe('6:00 PM');
    expect(f(1.25, '[h]:mm:ss')).toBe('30:00:00');
    expect(f(45292.75, '[$-409]mmm-yy')).toBe('Jan-24');
    expect(f(-1, 'd-mmm-yy')).toBeNull();
    expect(f(3e6, 'yyyy')).toBeNull();
    expect(isDateFormat('d-mmm-yy')).toBe(true);
    expect(isDateFormat('#,##0.0,,"m"')).toBe(false);
    expect(isDateFormat('0.0"x"')).toBe(false);
    expect(isDateFormat('[h]:mm')).toBe(true);
  });

  it('shows General as a cell does: at most 11 characters', () => {
    expect(generalText(1 / 3)).toBe('0.333333333');
    expect(generalText(1234567.891)).toBe('1234567.891');
    expect(generalText(123456789012)).toBe('1.23457E+11');
    expect(generalText(-0.5)).toBe('-0.5');
    expect(f(-12.5, 'General" units"')).toBe('-12.5 units');
  });

  it('is what TEXT uses, and WEEKDAY agrees on early 1900', () => {
    const e = new Engine(tableHost([]));
    const calc = (x: string) => e.evaluate(x, 'S', 0, 0).value;
    expect(calc('=TEXT(-1234.5,"#,##0.00_);(#,##0.00)")')).toBe('(1,234.50)');
    expect(calc('=WEEKDAY(1)')).toBe(1);
    expect(calc('=WEEKDAY(59)')).toBe(3);
    expect(calc('=WEEKDAY(61)')).toBe(5);
  });
});

describe('number formats from an .xlsx', () => {
  it('reads each cell’s format, shows it, and survives the worker transfer', async () => {
    const ExcelJS = (await import('exceljs')).default;
    const { parseXlsx, diffTables, packTable, unpackTable } = await import('../src/core');
    const { displayCell } = await import('../src/core/display');
    const { renderHTML } = await import('../src/html');
    const build = async (rate: number) => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('S');
      ws.addRow(['item', 'value']);
      ws.addRow(['Rate', rate]).getCell(2).numFmt = '0.0%';
      ws.addRow(['Revenue', 1234567]).getCell(2).numFmt = '$#,##0';
      ws.addRow(['Date', new Date(Date.UTC(2024, 0, 15))]).getCell(2).numFmt = 'd-mmm-yy';
      ws.addRow(['Plain', 1.5]);
      return new Uint8Array(await wb.xlsx.writeBuffer());
    };
    const a = parseXlsx(await build(0.05)).sheets[0];
    const b = parseXlsx(await build(0.062)).sheets[0];
    expect([0, 1, 2, 3].map((r) => displayCell(b, r, 1))).toEqual(['6.2%', '$1,234,567', '15-Jan-24', '1.5']);
    // Text cells don't carry a format; nor do General ones.
    expect(b.numberFormats!.rows[3]).toBeUndefined();
    const back = unpackTable(packTable(b));
    expect([0, 1, 2, 3].map((r) => displayCell(back, r, 1))).toEqual(['6.2%', '$1,234,567', '15-Jan-24', '1.5']);
    // The HTML renderer shows formatted values, and the change in points.
    const html = renderHTML(diffTables(a, b), { display: 'delta' });
    expect(html).toContain('6.2%');
    expect(html).toContain('+1.2 pp');
    expect(renderHTML(diffTables(a, b), { numberFormats: false })).toContain('0.062');
  });
});
