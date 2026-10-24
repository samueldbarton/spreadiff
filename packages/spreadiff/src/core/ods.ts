/**
 * OpenDocument spreadsheets (.ods, LibreOffice's own format): values,
 * formulas (converted from OpenFormula to Excel's A1 syntax), merged cells
 * and hidden rows. Number formats aren't read; values show as stored.
 */
import { strFromU8, unzipSync } from 'fflate';
import { rawSheet, tablesFromRaw, type RawSheet } from './sheet.js';
import { attr, decode, eachElement } from './xml.js';
import { serialToIso } from './serial.js';
import type { CellValue, Table, Workbook } from './types.js';

/** Options for {@link parseOds}. */
export interface ParseOdsOptions {
  /** Workbook name. Default `workbook.ods`. */
  name?: string;
  /** Whether the first row holds column names. Default `'auto'` (as for .xlsx). */
  header?: boolean | 'auto';
}

/** LibreOffice error codes → Excel errors (as LibreOffice's .xlsx export maps them; others are #VALUE!). */
const LO_ERRORS: Record<number, string> = { 503: '#NUM!', 519: '#VALUE!', 521: '#NULL!', 524: '#REF!', 525: '#NAME?', 532: '#DIV/0!', 32767: '#N/A' };

// A repeated run of empty cells or rows (trailing ones run to the sheet's end) only moves the
// position; a run with content is expanded, up to the sheet's size.
const MAX_ROWS = 1_048_576;
const MAX_COLS = 16_384;
const count = (v: string | undefined, max: number) => Math.max(1, Math.min(Number(v ?? 1) || 1, max));

/** A reference into another file, `'file:///dir/Book.xlsx'#$Sheet.A1`, as Excel writes it: `'[Book.xlsx]Sheet'!A1`. */
function externalRef(inner: string): string | null {
  const m = /^'((?:[^']|'')*)'#(.*)$/.exec(inner);
  if (!m) return null;
  let file = m[1].replace(/''/g, "'");
  try {
    file = decodeURIComponent(file);
  } catch {
    // keep it as written
  }
  file = file.replace(/^.*[\\/]/, '');
  const ref = refToExcel(m[2]);
  const bang = ref.lastIndexOf('!');
  if (bang < 0) return `[${file}]${ref}`;
  const sheet = ref.slice(0, bang).replace(/^'(.*)'$/, '$1').replace(/''/g, "'");
  return `'${`[${file}]${sheet}`.replace(/'/g, "''")}'${ref.slice(bang)}`;
}

/**
 * An OpenFormula expression in Excel's syntax: `of:=[.A1]+['Data'.B2:.C3]`
 * → `=A1+Data!B2:C3`, `;` between arguments → `,`.
 */
export function openFormulaToExcel(f: string): string {
  let s = f.replace(/^of:/, '');
  if (!s.startsWith('=')) s = '=' + s;
  let out = '';
  let braces = 0;
  for (let i = 0; i < s.length; i++) {
    const ch = s[i];
    if (ch === '"') {
      const end = s.indexOf('"', i + 1);
      const j = end < 0 ? s.length - 1 : end;
      out += s.slice(i, j + 1);
      i = j;
    } else if (ch === '[') {
      const end = s.indexOf(']', i);
      const inner = s.slice(i + 1, end < 0 ? s.length : end);
      i = end < 0 ? s.length : end;
      // [1]Sheet!A1: Excel's own external-link syntax, which LibreOffice keeps as written.
      if (/^\d+$/.test(inner)) out += `[${inner}]`;
      else out += externalRef(inner) ?? refToExcel(inner);
    } else if (ch === ';') out += ',';
    else if (ch === '|' && braces > 0) out += ';'; // array constant rows: {1;2|3;4} → {1,2;3,4}
    else {
      if (ch === '{') braces++;
      else if (ch === '}') braces--;
      out += ch;
    }
  }
  // Functions LibreOffice keeps in Excel's namespace (COM.MICROSOFT.CEILING) are Excel's own.
  return out.replace(/\bCOM\.MICROSOFT\.(?=[A-Z][A-Z0-9.]*\()/g, '');
}

/**
 * `.A1`, `Sheet.A1:.B2`, `$'My sheet'.$A$1`, `Jan.A1:Mar.A1` → `A1`,
 * `Sheet!A1:B2`, `'My sheet'!$A$1`, `Jan:Mar!A1` (a 3-D reference).
 */
function refToExcel(ref: string): string {
  const quote = (sh: string) => (sh && !/^'.*'$/.test(sh) && /[^A-Za-z0-9_]/.test(sh) ? `'${sh.replace(/'/g, "''")}'` : sh);
  const sheets: string[] = [];
  const cells = ref.split(':').map((p) => {
    const dot = p.lastIndexOf('.');
    const sh = (dot > 0 ? p.slice(0, dot) : '').replace(/^\$/, '');
    sheets.push(sh);
    return dot >= 0 ? p.slice(dot + 1) : p;
  });
  const [first, last] = [sheets[0], sheets[sheets.length - 1]];
  const prefix = first ? (last && last !== first ? `${quote(first)}:${quote(last)}!` : `${quote(first)}!`) : '';
  return prefix + cells.join(':');
}

/** The text of a cell: its paragraphs joined by newlines, with `<text:s>` spaces and tabs. Comments and shapes aren't its text. */
function cellText(body: string): string {
  const paras: string[] = [];
  const own = body.replace(/<office:annotation\b[\s\S]*?<\/office:annotation>/g, '').replace(/<draw:([\w-]+)(?=[\s>\/])[\s\S]*?<\/draw:\1>/g, '');
  eachElement(own, 'text:p', (_, b) => {
    const t = (b ?? '')
      .replace(/<text:s(?:\s+text:c="(\d+)")?\s*\/>/g, (_m, n: string | undefined) => ' '.repeat(Number(n ?? 1)))
      .replace(/<text:tab\s*\/>/g, '\t')
      .replace(/<text:line-break\s*\/>/g, '\n')
      .replace(/<[^>]+>/g, '');
    paras.push(decode(t));
  });
  return paras.join('\n');
}

/**
 * A cell value from `office:date-value` (`2024-01-15`, `2024-01-15T12:30:00`),
 * as the XLSX reader would give it: ISO text in Excel's reckoning. Before
 * 1900-03-01 Excel's serials run a day apart from the calendar (its fictional
 * 29 February), and a date before its serial 0 is just a number.
 */
function odsDate(v: string): CellValue {
  const iso = v.replace('T', ' ').replace(/(\.\d*?)0+$/, '$1').replace(/\.$/, '').replace(/ 00:00:00$/, '');
  if (iso >= '1900-03-01') return iso;
  const ms = Date.parse(v.length <= 10 ? `${v}T00:00:00Z` : `${v}Z`);
  if (!Number.isFinite(ms)) return iso;
  // LibreOffice's serials count from 1899-12-30 and agree with Excel's from 1900-03-01.
  const serial = (ms - Date.UTC(1899, 11, 30)) / 86400000;
  return serialToIso(serial, false) ?? serial;
}

/** `PT12H30M00S` → `12:30:00` (hours can exceed 24 for durations). */
function isoDuration(v: string): string {
  const m = /^-?P(?:(\d+)D)?T?(?:(\d+)H)?(?:(\d+)M)?(?:([\d.]+)S)?$/.exec(v);
  if (!m) return v;
  const h = Number(m[1] ?? 0) * 24 + Number(m[2] ?? 0);
  const sec = Number(m[4] ?? 0);
  const whole = Math.floor(sec + 1e-9);
  const ms = Math.round((sec - whole) * 1000);
  return `${String(h).padStart(2, '0')}:${String(Number(m[3] ?? 0)).padStart(2, '0')}:${String(whole).padStart(2, '0')}${ms ? `.${String(ms).padStart(3, '0').replace(/0+$/, '')}` : ''}`;
}

/** Parse an .ods file into a workbook. */
export function parseOds(data: ArrayBuffer | Uint8Array, options: ParseOdsOptions = {}): Workbook {
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
  const files = unzipSync(bytes, { filter: (f) => f.name === 'content.xml' || f.name === 'mimetype' });
  if (!files['content.xml']) throw new Error('Not an OpenDocument spreadsheet (missing content.xml)');
  const xml = strFromU8(files['content.xml']);
  const raws: RawSheet[] = [];
  const names: Record<string, string> = {};
  eachElement(xml, 'table:named-range', (a) => {
    const name = attr(a, 'table:name');
    const range = attr(a, 'table:cell-range-address');
    if (name && range) names[decode(name)] = openFormulaToExcel(`[${decode(range)}]`).slice(1);
  });
  eachElement(xml, 'table:table', (tableAttrs, body) => {
    const name = decode(attr(tableAttrs, 'table:name') ?? `Sheet${raws.length + 1}`);
    const grid: CellValue[][] = [];
    const formulas: (string | null)[][] = [];
    const merges: Table['merges'] = [];
    const hidden: number[] = [];
    const filters: { r0: number; r1: number }[] = [];
    const arrays: { r0: number; c0: number; r1: number; c1: number; formula: string }[] = [];
    let r = 0;
    // Rows can sit in header-row and row-group wrappers: scan every row element in order.
    const rowRe = /<table:table-row\b([^>]*?)(\/>|>([\s\S]*?)<\/table:table-row>)/g;
    for (const m of (body ?? '').matchAll(rowRe)) {
      const rowAttrs = m[1];
      const rowBody = m[3] ?? '';
      const repeat = count(attr(rowAttrs, 'table:number-rows-repeated'), MAX_ROWS);
      const visibility = attr(rowAttrs, 'table:visibility') ?? '';
      const isHidden = visibility === 'collapse' || visibility === 'filter';
      const row: CellValue[] = [];
      const frow: (string | null)[] = [];
      let c = 0;
      let any = false;
      const cellRe = /<table:(table-cell|covered-table-cell)\b([^>]*?)(\/>|>([\s\S]*?)<\/table:\1>)/g;
      for (const cm of rowBody.matchAll(cellRe)) {
        const a = cm[2];
        const cb = cm[4] ?? '';
        const reps = count(attr(a, 'table:number-columns-repeated'), MAX_COLS);
        let value: CellValue = null;
        const type = attr(a, 'office:value-type') ?? attr(a, 'calcext:value-type');
        if (cm[1] === 'table-cell') {
          switch (type) {
            case 'float':
            case 'percentage':
            case 'currency':
              value = Number(attr(a, 'office:value'));
              break;
            case 'date':
              value = odsDate(attr(a, 'office:date-value') ?? '');
              break;
            case 'time':
              value = isoDuration(attr(a, 'office:time-value') ?? '');
              break;
            case 'boolean':
              value = attr(a, 'office:boolean-value') === 'true';
              break;
            case 'string':
              value = attr(a, 'office:string-value') !== undefined ? decode(attr(a, 'office:string-value')!) : cellText(cb);
              break;
            default:
              if (cb.includes('<text:p')) value = cellText(cb) || null;
          }
          // An error result is stored as text with a calcext type of error; LibreOffice's own
          // codes (Err:502) become the Excel errors its .xlsx export writes for them.
          if (attr(a, 'calcext:value-type') === 'error') {
            const t = cellText(cb);
            const code = /^Err:(\d+)$/.exec(t)?.[1];
            value = code ? (LO_ERRORS[Number(code)] ?? '#VALUE!') : t;
          }
        }
        const f = attr(a, 'table:formula');
        let formula = f ? openFormulaToExcel(decode(f)) : null;
        // An error constant is written as the formula =#VALUE!: it's a value (as the .xlsx reader has it).
        if (formula && /^=#(?:VALUE!|N\/A|REF!|DIV\/0!|NUM!|NAME\?|NULL!)$/.test(formula) && (value === null || value === formula.slice(1))) {
          value = formula.slice(1);
          formula = null;
        }
        // LibreOffice stores a typed TRUE or FALSE as the formula =TRUE() / =FALSE(): it's a value, as in Excel.
        if (formula && /^=(TRUE|FALSE)\(\)$/i.test(formula) && typeof value === 'boolean') formula = null;
        // An array formula sits on its top-left cell; every cell of the block holds it (as in .xlsx).
        const mc = Number(attr(a, 'table:number-matrix-columns-spanned') ?? 0);
        const mr = Number(attr(a, 'table:number-matrix-rows-spanned') ?? 0);
        if (formula && (mc > 1 || mr > 1)) arrays.push({ r0: r, c0: c, r1: r + Math.max(1, mr) - 1, c1: c + Math.max(1, mc) - 1, formula });
        const cs = Number(attr(a, 'table:number-columns-spanned') ?? 1);
        const rs = Number(attr(a, 'table:number-rows-spanned') ?? 1);
        if (cs > 1 || rs > 1) merges.push({ r0: r, c0: c, r1: r + rs - 1, c1: c + cs - 1 });
        if (value === null && !formula) c += reps;
        else
          for (let k = 0; k < reps && c < MAX_COLS; k++) {
            row[c] = value;
            frow[c] = formula;
            any = true;
            c++;
          }
      }
      // Empty rows only move the position: hiding one changes no total, so it isn't recorded.
      if (!any) r += repeat;
      else
        for (let k = 0; k < repeat && r < MAX_ROWS; k++) {
          if (isHidden) hidden.push(r);
          if (visibility === 'filter') filters.push({ r0: r, r1: r });
          grid[r] = row.slice();
          formulas[r] = frow.slice();
          r++;
        }
    }
    for (const b of arrays) {
      for (let rr = b.r0; rr <= b.r1; rr++) {
        const fr = (formulas[rr] ??= []);
        for (let cc = b.c0; cc <= b.c1; cc++) {
          fr[cc] ??= b.formula;
          (grid[rr] ??= [])[cc] ??= null;
        }
      }
    }
    raws.push(rawSheet(name, grid, formulas, { merges, hidden, filters, arrays: arrays.map(({ r0, c0, r1, c1 }) => ({ r0, c0, r1, c1 })) }));
  });
  const sheets = tablesFromRaw(raws, options.header ?? 'auto', names);
  return { name: options.name ?? 'workbook.ods', sheets, ...(Object.keys(names).length ? { names } : {}) };
}
