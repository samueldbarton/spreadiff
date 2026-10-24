import { formatCell } from './text.js';
import { formatNumberCode, formatTextCode, isDateFormat } from './numfmt.js';
import { isoToSerial } from './serial.js';
import type { CellValue, Table } from './types.js';

/** A cell's number format code (`'#,##0.00'`), or undefined for General or no format. */
export function formatCodeAt(table: Table, row: number, col: number): string | undefined {
  const nf = table.numberFormats;
  const i = nf?.rows[row]?.[col];
  return i ? nf!.codes[i] : undefined;
}

/**
 * A value as a cell with this format code shows it: `0.123` with `0.0%` is
 * `12.3%`, a date (kept as ISO text) with `d-mmm-yy` is `1-Jan-24`. Falls
 * back to `formatCell` without a code, or for one it can't apply.
 */
export function displayValue(v: CellValue | undefined, code: string | undefined, date1904 = false): string {
  if (v === undefined || v === null) return '';
  if (code) {
    if (typeof v === 'number') {
      const s = formatNumberCode(v, code, date1904, formatCell);
      if (s !== null) return s;
    } else if (typeof v === 'string') {
      const serial = isDateFormat(code) && /^\d/.test(v) ? isoToSerial(v, date1904) : null;
      if (serial !== null) {
        const s = formatNumberCode(serial, code, date1904, formatCell);
        if (s !== null) return s;
      } else {
        const s = formatTextCode(v, code);
        if (s !== null) return s;
      }
    }
  }
  return formatCell(v);
}

/** A cell's value as the sheet shows it (see {@link displayValue}). */
export function displayCell(table: Table, row: number, col: number): string {
  return displayValue(table.rows[row]?.[col], formatCodeAt(table, row, col), table.numberFormats?.date1904);
}

/**
 * The size of a change in a cell's own format (`$1,200` for a `$#,##0`
 * cell), unsigned. A percentage cell's change is in points (`1.2 pp`), so
 * it isn't read as a relative change. Dates and times keep the plain number (a difference of
 * dates is days, not a date), and so does a change too small for the
 * format to show.
 */
export function displayMagnitude(abs: number, code: string | undefined): string {
  if (code && !isDateFormat(code)) {
    const s = formatNumberCode(abs, code, false, formatCell);
    // A change too small for the format ("0" for 0.4 in a whole-number cell) shows as the plain number.
    if (s !== null && /[1-9]/.test(s)) {
      const t = s.trim();
      return /%/.test(code.replace(/"[^"]*"|\\./g, '')) && t.endsWith('%') ? `${t.slice(0, -1)} pp` : t;
    }
  }
  return formatCell(abs);
}
