/**
 * A small, fast XLSX reader: cached values, formulas and dates. It does not
 * evaluate formulas or keep styling; for diffing, the stored values are what
 * you want. ~4 KB on top of fflate, versus ~900 KB for a full SheetJS build.
 */
import { unzipSync, strFromU8 } from 'fflate';
import { columnLabel, uniqueNames } from './csv.js';
import type { CellValue, Table, Workbook } from './types.js';

export interface ParseXlsxOptions {
  /** Treat each sheet's first non-empty row as column names. Default true. */
  header?: boolean;
  /** Only parse these sheets (by name). */
  sheets?: string[];
  name?: string;
}

export function parseXlsx(data: ArrayBuffer | Uint8Array, options: ParseXlsxOptions = {}): Workbook {
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
  const files = unzipSync(bytes, {
    filter: (f) => f.name.startsWith('xl/') && f.name.endsWith('.xml') || f.name.endsWith('.rels'),
  });
  const text = (path: string) => (files[path] ? strFromU8(files[path]) : '');

  const workbookXml = text('xl/workbook.xml');
  if (!workbookXml) throw new Error('Not an XLSX file (missing xl/workbook.xml)');
  const date1904 = /<workbookPr\b[^>]*date1904="(1|true)"/.test(workbookXml);
  const rels = new Map<string, string>();
  for (const m of text('xl/_rels/workbook.xml.rels').matchAll(/<Relationship\b([^>]*)\/?>/g)) {
    const id = attr(m[1], 'Id');
    const target = attr(m[1], 'Target');
    if (id && target) rels.set(id, target.startsWith('/') ? target.slice(1) : `xl/${target}`);
  }
  const strings = parseSharedStrings(text('xl/sharedStrings.xml'));
  const dateStyles = parseDateStyles(text('xl/styles.xml'));

  const sheets: Table[] = [];
  for (const m of workbookXml.matchAll(/<sheet\b([^>]*)\/?>/g)) {
    const name = decode(attr(m[1], 'name') ?? `Sheet${sheets.length + 1}`);
    if (options.sheets && !options.sheets.includes(name)) continue;
    const rid = attr(m[1], 'r:id');
    const path = rid ? rels.get(rid) : undefined;
    if (!path || !files[path]) continue;
    sheets.push(parseSheet(name, strFromU8(files[path]), strings, dateStyles, date1904, options.header ?? true));
  }
  return { name: options.name ?? 'workbook.xlsx', sheets };
}

function attr(attrs: string, name: string): string | undefined {
  const i = attrs.indexOf(` ${name}="`);
  if (i < 0) return undefined;
  const start = i + name.length + 3;
  return attrs.slice(start, attrs.indexOf('"', start));
}

const ENTITY = /&(amp|lt|gt|quot|apos|#x[0-9a-fA-F]+|#\d+);/g;
function decode(s: string): string {
  if (s.indexOf('&') < 0) return s;
  return s.replace(ENTITY, (_, e: string) => {
    switch (e) {
      case 'amp': return '&';
      case 'lt': return '<';
      case 'gt': return '>';
      case 'quot': return '"';
      case 'apos': return "'";
      default:
        return String.fromCodePoint(e[1] === 'x' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10));
    }
  });
}

function parseSharedStrings(xml: string): string[] {
  const out: string[] = [];
  for (const si of xml.matchAll(/<si>([\s\S]*?)<\/si>/g)) {
    // Concatenate every <t> (rich text runs), ignoring phonetic <rPh> runs.
    const body = si[1].replace(/<rPh\b[\s\S]*?<\/rPh>/g, '');
    let s = '';
    for (const t of body.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)) s += t[1];
    out.push(decode(s));
  }
  return out;
}

const BUILTIN_DATE_FORMATS = new Set([14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47]);

/** Set of cellXfs indices whose number format renders as a date/time. */
function parseDateStyles(xml: string): Set<number> {
  const custom = new Map<number, string>();
  for (const m of xml.matchAll(/<numFmt\b([^>]*)\/?>/g)) {
    const id = Number(attr(m[1], 'numFmtId'));
    custom.set(id, decode(attr(m[1], 'formatCode') ?? ''));
  }
  const isDate = (id: number) => {
    if (BUILTIN_DATE_FORMATS.has(id)) return true;
    const code = custom.get(id);
    if (!code) return false;
    const stripped = code.replace(/"[^"]*"|\[[^\]]*\]|\\./g, '');
    return /[dmyhs]/i.test(stripped) && !/^[#0.,%\s]*$/.test(stripped);
  };
  const out = new Set<number>();
  const xfs = /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(xml);
  if (!xfs) return out;
  let i = 0;
  for (const m of xfs[1].matchAll(/<xf\b([^>]*)\/?>/g)) {
    if (isDate(Number(attr(m[1], 'numFmtId') ?? 0))) out.add(i);
    i++;
  }
  return out;
}

function serialToIso(serial: number, date1904: boolean): string {
  const ms = Math.round((serial - (date1904 ? -1462 : 0) - 25569) * 86400000);
  const iso = new Date(ms).toISOString();
  if (Number.isInteger(serial)) return iso.slice(0, 10);
  if (serial < 1) return iso.slice(11, 19);
  return `${iso.slice(0, 10)} ${iso.slice(11, 19)}`;
}

function decodeRef(ref: string): [row: number, col: number] {
  let col = 0;
  let i = 0;
  for (; i < ref.length; i++) {
    const c = ref.charCodeAt(i);
    if (c < 65 || c > 90) break;
    col = col * 26 + (c - 64);
  }
  return [Number(ref.slice(i)) - 1, col - 1];
}

/** Shift relative A1 references in a shared formula by a row/column offset. */
function shiftFormula(formula: string, dRow: number, dCol: number): string {
  return formula.replace(
    /("[^"]*")|(?<![A-Za-z0-9_.!])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\d(A-Za-z_])/g,
    (all, str, cAbs, col, rAbs, row) => {
      if (str) return all;
      let c = 0;
      for (const ch of col as string) c = c * 26 + (ch.charCodeAt(0) - 64);
      const nc = cAbs ? c - 1 : c - 1 + dCol;
      const nr = rAbs ? Number(row) : Number(row) + dRow;
      if (nc < 0 || nr < 1) return all;
      return `${cAbs}${columnLabel(nc)}${rAbs}${nr}`;
    },
  );
}

const CELL = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
const ROW = /<row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/g;

function parseSheet(
  name: string,
  xml: string,
  strings: string[],
  dateStyles: Set<number>,
  date1904: boolean,
  header: boolean,
): Table {
  const grid: CellValue[][] = [];
  const formulaGrid: (string | null)[][] = [];
  const shared = new Map<string, { formula: string; row: number; col: number }>();
  let hasFormulas = false;
  let minCol = Infinity;
  let maxCol = -1;
  let implicitRow = 0;
  for (const rm of xml.matchAll(ROW)) {
    const rAttr = attr(rm[1], 'r');
    const r = rAttr ? Number(rAttr) - 1 : implicitRow;
    implicitRow = r + 1;
    if (!rm[2]) continue;
    let implicitCol = 0;
    for (const cm of rm[2].matchAll(CELL)) {
      const attrs = cm[1];
      const body = cm[2] ?? '';
      const ref = attr(attrs, 'r');
      const c = ref ? decodeRef(ref)[1] : implicitCol;
      implicitCol = c + 1;
      const t = attr(attrs, 't') ?? 'n';
      const vm = /<v>([\s\S]*?)<\/v>/.exec(body);
      const raw = vm ? decode(vm[1]) : null;
      let value: CellValue = null;
      switch (t) {
        case 's':
          value = raw === null ? null : strings[Number(raw)] ?? null;
          break;
        case 'inlineStr': {
          let s = '';
          for (const tm of body.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)) s += tm[1];
          value = decode(s);
          break;
        }
        case 'str':
        case 'e':
          value = raw;
          break;
        case 'b':
          value = raw === '1';
          break;
        case 'd':
          value = raw;
          break;
        default:
          if (raw !== null && raw !== '') {
            const num = Number(raw);
            const style = Number(attr(attrs, 's') ?? -1);
            value = dateStyles.has(style) ? serialToIso(num, date1904) : num;
          }
      }
      let formula: string | null = null;
      const fm = /<f\b([^>]*?)(?:\/>|>([\s\S]*?)<\/f>)/.exec(body);
      if (fm) {
        const fAttrs = fm[1];
        const text = fm[2] ? decode(fm[2]) : '';
        const si = attr(fAttrs, 'si');
        if (attr(fAttrs, 't') === 'shared' && si !== undefined) {
          if (text) {
            shared.set(si, { formula: text, row: r, col: c });
            formula = text;
          } else {
            const master = shared.get(si);
            if (master) formula = shiftFormula(master.formula, r - master.row, c - master.col);
          }
        } else if (text) {
          formula = text;
        }
      }
      if (value === null && formula === null) continue;
      (grid[r] ??= [])[c] = value;
      if (formula !== null) {
        hasFormulas = true;
        (formulaGrid[r] ??= [])[c] = '=' + formula;
      }
      if (c < minCol) minCol = c;
      if (c > maxCol) maxCol = c;
    }
  }
  if (maxCol < 0) return { name, columns: [], rows: [] };

  let first = 0;
  while (first < grid.length && !grid[first]) first++;
  const width = maxCol - minCol + 1;
  const dense = (src: CellValue[] | undefined): CellValue[] => {
    const out = new Array<CellValue>(width);
    for (let c = 0; c < width; c++) out[c] = src?.[minCol + c] ?? null;
    return out;
  };
  const denseF = (src: (string | null)[] | undefined): (string | null)[] => {
    const out = new Array<string | null>(width);
    for (let c = 0; c < width; c++) out[c] = src?.[minCol + c] ?? null;
    return out;
  };
  const columns = header
    ? uniqueNames(dense(grid[first]).map((v) => (v === null ? '' : String(v))))
    : Array.from({ length: width }, (_, c) => columnLabel(minCol + c));
  const start = header ? first + 1 : first;
  const rows: CellValue[][] = [];
  const formulas: (string | null)[][] = [];
  for (let r = start; r < grid.length; r++) {
    rows.push(dense(grid[r]));
    if (hasFormulas) formulas.push(denseF(formulaGrid[r]));
  }
  // Trim trailing empty rows.
  while (rows.length && rows[rows.length - 1].every((v) => v === null)) {
    rows.pop();
    if (hasFormulas) formulas.pop();
  }
  return hasFormulas ? { name, columns, rows, formulas } : { name, columns, rows };
}
