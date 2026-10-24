/**
 * Excel date serials ↔ calendar dates, in both date systems. Shared by the
 * engine and the number formatter; kept apart so the HTML renderer can
 * format dates without loading the formula engine.
 */

const DAY = 86400000;
const EPOCH = Date.UTC(1899, 11, 30); // serial 0 in the (bug-compatible) 1900 system, for serials ≥ 61

/** Excel serial of a calendar date; month and day overflow roll over like `DATE()`. */
export function serialFromParts(y: number, m: number, d: number, date1904 = false): number {
  // DATE() normalizes month/day overflow like JS Date.
  const ms = Date.UTC(y, m - 1, d);
  let s = Math.round((ms - EPOCH) / DAY);
  if (date1904) return s - 1462;
  if (s < 61) s -= 1; // before 1900-03-01: Excel counts the fictitious 1900-02-29
  return s;
}

/** Calendar date and weekday (0 = Sunday) of an Excel serial, as Excel reckons them. */
export function partsFromSerial(serial: number, date1904 = false): { y: number; m: number; d: number; dow: number } {
  let s = Math.floor(serial);
  if (date1904) s += 1462;
  else if (s < 61) {
    if (s === 60) return { y: 1900, m: 2, d: 29, dow: 3 };
    if (s === 0) return { y: 1900, m: 1, d: 0, dow: 6 }; // Excel's "1900-01-00"
    // Before the fictional 29 February, Excel's weekdays run a day behind the calendar (1900-01-01 is a "Sunday").
    if (s < 1) s += 1;
    else {
      const dt = new Date(EPOCH + (s + 1) * DAY);
      return { y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1, d: dt.getUTCDate(), dow: (s - 1) % 7 };
    }
  }
  const dt = new Date(EPOCH + s * DAY);
  return { y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1, d: dt.getUTCDate(), dow: dt.getUTCDay() };
}

const ISO = /^(\d{4})-(\d{2})-(\d{2})(?: (\d{2}):(\d{2}):(\d{2}(?:\.\d+)?))?$/;
const TIME = /^(\d{2}):(\d{2}):(\d{2}(?:\.\d+)?)$/;

/** The parser's ISO rendering of a date-formatted cell, back to its serial (or null if it isn't one). */
export function isoToSerial(s: string, date1904 = false): number | null {
  let m = ISO.exec(s);
  if (m) {
    const day = serialFromParts(Number(m[1]), Number(m[2]), Number(m[3]), date1904);
    const t = m[4] ? (Number(m[4]) * 3600 + Number(m[5]) * 60 + Number(m[6])) / 86400 : 0;
    return day + t;
  }
  m = TIME.exec(s);
  if (m) return (Number(m[1]) * 3600 + Number(m[2]) * 60 + Number(m[3])) / 86400;
  return null;
}
