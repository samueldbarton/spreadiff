/**
 * A small, fast XLSX reader: cached values, formulas and dates. It does not
 * evaluate formulas or keep styling; for diffing, the stored values are what
 * you want. ~4 KB on top of fflate, versus ~900 KB for a full SheetJS build.
 *
 * Elements are found with `indexOf`, never with backtracking regexes, so a
 * malformed or hostile file costs linear time.
 */
import { unzipSync, strFromU8 } from 'fflate';
import { columnLabel, uniqueNames } from './csv.js';
import type { CellValue, Table, Workbook } from './types.js';

/** Options for {@link parseXlsx}. */
export interface ParseXlsxOptions {
  /** Treat each sheet's first non-empty row as column names. Default true. */
  header?: boolean;
  /** Only parse these sheets (by name). */
  sheets?: string[];
  /** Name for the resulting workbook (default `workbook.xlsx`). */
  name?: string;
}

/**
 * Read an `.xlsx` / `.xlsm` file into a {@link Workbook}.
 *
 * Reads cached values (formulas aren't evaluated), formula text (shared
 * formulas are expanded per cell), dates (from the cell's number format),
 * booleans, errors (as `#DIV/0!` strings), rich text and workbook-scoped
 * defined names. Each sheet is trimmed to its used range; `origin` records
 * where it starts. Styles, merged cells and charts are ignored.
 *
 * Throws if the data isn't a zip with `xl/workbook.xml`.
 */
export function parseXlsx(data: ArrayBuffer | Uint8Array, options: ParseXlsxOptions = {}): Workbook {
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
  const files = unzipSync(bytes, {
    filter: (f) =>
      f.name === 'xl/workbook.xml' ||
      f.name === 'xl/sharedStrings.xml' ||
      f.name === 'xl/styles.xml' ||
      f.name === 'xl/_rels/workbook.xml.rels' ||
      f.name.startsWith('xl/worksheets/'),
  });
  const text = (path: string) => (files[path] ? stripPrefixes(strFromU8(files[path])) : '');

  const workbookXml = text('xl/workbook.xml');
  if (!workbookXml) throw new Error('Not an XLSX file (missing xl/workbook.xml)');
  let date1904 = false;
  eachElement(workbookXml, 'workbookPr', (a) => {
    const v = attr(a, 'date1904');
    date1904 = v === '1' || v === 'true';
  });
  const rels = new Map<string, string>();
  eachElement(text('xl/_rels/workbook.xml.rels'), 'Relationship', (a) => {
    const id = attr(a, 'Id');
    const target = attr(a, 'Target');
    if (id && target) rels.set(id, target.startsWith('/') ? target.slice(1) : `xl/${target}`);
  });
  const strings = parseSharedStrings(text('xl/sharedStrings.xml'));
  const dateStyles = parseDateStyles(text('xl/styles.xml'));

  const sheets: Table[] = [];
  eachElement(workbookXml, 'sheet', (a) => {
    const name = decode(attr(a, 'name') ?? `Sheet${sheets.length + 1}`);
    if (options.sheets && !options.sheets.includes(name)) return;
    // The relationships prefix is conventionally `r:` but may be anything.
    const rid = attr(a, 'r:id') ?? /\s[\w.-]+:id\s*=\s*["']([^"']*)/.exec(a)?.[1];
    const path = rid ? rels.get(rid) : undefined;
    if (!path || !files[path]) return;
    sheets.push(parseSheet(name, text(path), strings, dateStyles, date1904, options.header ?? true));
  });
  // Defined names (TaxRate → Inputs!$B$2) let formula tracing see through them.
  const names: Record<string, string> = {};
  eachElement(workbookXml, 'definedName', (a, body) => {
    const name = attr(a, 'name');
    // Workbook-scoped only; skip built-ins like _xlnm.Print_Area.
    if (!name || !body || name.startsWith('_xlnm.') || attr(a, 'localSheetId') !== undefined) return;
    names[decode(name)] = decode(body);
  });
  return { name: options.name ?? 'workbook.xlsx', sheets, ...(Object.keys(names).length ? { names } : {}) };
}

/** Drop namespace prefixes (`<x:c>` → `<c>`), as written by e.g. the OpenXML SDK. */
function stripPrefixes(xml: string): string {
  if (!/<[A-Za-z_][\w.-]*:(?:workbook|worksheet|sst|styleSheet|Relationships)\b/.test(xml.slice(0, 4096))) return xml;
  return xml.replace(/<(\/?)[A-Za-z_][\w.-]*:/g, '<$1');
}

/**
 * Call `fn(attrs, body)` for each `<tag …>body</tag>` or `<tag …/>` in
 * `xml`. Stops at the first unclosed element instead of rescanning.
 */
function eachElement(xml: string, tag: string, fn: (attrs: string, body: string | null) => void): void {
  const open = `<${tag}`;
  const close = `</${tag}>`;
  let i = xml.indexOf(open);
  while (i >= 0) {
    const next = xml.charCodeAt(i + open.length);
    // Must be the whole tag name: followed by whitespace, '>' or '/'.
    if (next !== 62 && next !== 47 && next !== 32 && next !== 9 && next !== 10 && next !== 13) {
      i = xml.indexOf(open, i + 1);
      continue;
    }
    const gt = xml.indexOf('>', i);
    if (gt < 0) return;
    if (xml.charCodeAt(gt - 1) === 47) {
      fn(xml.slice(i + open.length, gt - 1), null);
      i = xml.indexOf(open, gt + 1);
      continue;
    }
    const end = xml.indexOf(close, gt + 1);
    if (end < 0) return;
    fn(xml.slice(i + open.length, gt), xml.slice(gt + 1, end));
    i = xml.indexOf(open, end + close.length);
  }
}

/** Concatenated text of every `<t>` in `body`, skipping phonetic `<rPh>` runs. */
function textRuns(body: string): string {
  let s = '';
  const clean = body.indexOf('<rPh') < 0 ? body : removeElements(body, 'rPh');
  eachElement(clean, 't', (_, b) => {
    if (b) s += b;
  });
  return decode(s);
}

function removeElements(xml: string, tag: string): string {
  let out = '';
  let pos = 0;
  const open = `<${tag}`;
  const close = `</${tag}>`;
  for (;;) {
    const i = xml.indexOf(open, pos);
    if (i < 0) break;
    const end = xml.indexOf(close, i);
    if (end < 0) break;
    out += xml.slice(pos, i);
    pos = end + close.length;
  }
  return out + xml.slice(pos);
}

/** Attribute value by name; accepts either quote style and any whitespace. */
function attr(attrs: string, name: string): string | undefined {
  let i = attrs.indexOf(name);
  while (i >= 0) {
    const before = attrs.charCodeAt(i - 1);
    let j = i + name.length;
    if (before === 32 || before === 9 || before === 10 || before === 13) {
      while (attrs.charCodeAt(j) === 32) j++;
      if (attrs.charCodeAt(j) === 61) {
        j++;
        while (attrs.charCodeAt(j) === 32) j++;
        const q = attrs[j];
        if (q === '"' || q === "'") {
          const end = attrs.indexOf(q, j + 1);
          if (end >= 0) return attrs.slice(j + 1, end);
        }
      }
    }
    i = attrs.indexOf(name, i + 1);
  }
  return undefined;
}

const ENTITY = /&(amp|lt|gt|quot|apos|#x[0-9a-fA-F]+|#\d+);/g;
function decode(s: string): string {
  if (s.indexOf('&') < 0) return s;
  return s.replace(ENTITY, (all, e: string) => {
    switch (e) {
      case 'amp': return '&';
      case 'lt': return '<';
      case 'gt': return '>';
      case 'quot': return '"';
      case 'apos': return "'";
      default: {
        const cp = e[1] === 'x' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
        return cp >= 0 && cp <= 0x10ffff ? String.fromCodePoint(cp) : all;
      }
    }
  });
}

function parseSharedStrings(xml: string): string[] {
  const out: string[] = [];
  eachElement(xml, 'si', (_, body) => out.push(body ? textRuns(body) : ''));
  return out;
}

const BUILTIN_DATE_FORMATS = new Set([14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47]);

/** Set of cellXfs indices whose number format renders as a date/time. */
function parseDateStyles(xml: string): Set<number> {
  const custom = new Map<number, string>();
  eachElement(xml, 'numFmt', (a) => {
    custom.set(Number(attr(a, 'numFmtId')), decode(attr(a, 'formatCode') ?? ''));
  });
  const isDate = (id: number) => {
    if (BUILTIN_DATE_FORMATS.has(id)) return true;
    const code = custom.get(id);
    if (!code) return false;
    const stripped = code.replace(/"[^"]*"|\[[^\]]*\]|\\./g, '');
    return /[dmyhs]/i.test(stripped) && !/^[#0.,%\s]*$/.test(stripped);
  };
  const out = new Set<number>();
  eachElement(xml, 'cellXfs', (_, body) => {
    let i = 0;
    eachElement(body ?? '', 'xf', (a) => {
      if (isDate(Number(attr(a, 'numFmtId') ?? 0))) out.add(i);
      i++;
    });
  });
  return out;
}

/** Excel serial → ISO date/time, or null when it isn't a representable date. */
function serialToIso(serial: number, date1904: boolean): string | null {
  if (!Number.isFinite(serial) || serial < 0 || serial > 2958465) return null; // > 9999-12-31
  // Excel's 1900 system pretends 1900-02-29 existed (serial 60); serials
  // below it are one day later than the naive conversion.
  const s = date1904 ? serial + 1462 : serial < 60 ? serial + 1 : serial;
  const ms = Math.round((s - 25569) * 86400000);
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
    // String literals and quoted sheet names pass through untouched. A ref
    // may follow `!` (Sheet!A1) but not be followed by `!` (it's a sheet name).
    /("[^"]*")|('[^']*')|(?<![A-Za-z0-9_.])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\d(A-Za-z_!])/g,
    (all, str, sheet, cAbs, col, rAbs, row) => {
      if (str || sheet) return all;
      let c = 0;
      for (const ch of col as string) c = c * 26 + (ch.charCodeAt(0) - 64);
      const nc = cAbs ? c - 1 : c - 1 + dCol;
      const nr = rAbs ? Number(row) : Number(row) + dRow;
      if (nc < 0 || nr < 1) return all;
      return `${cAbs}${columnLabel(nc)}${rAbs}${nr}`;
    },
  );
}

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
  let minCol = Infinity;
  let maxCol = -1;
  let implicitRow = 0;
  eachElement(xml, 'row', (rowAttrs, rowBody) => {
    const rAttr = attr(rowAttrs, 'r');
    const r = rAttr ? Number(rAttr) - 1 : implicitRow;
    implicitRow = r + 1;
    if (!rowBody || !(r >= 0)) return;
    let implicitCol = 0;
    eachElement(rowBody, 'c', (attrs, cellBody) => {
      const body = cellBody ?? '';
      const ref = attr(attrs, 'r');
      const c = ref ? decodeRef(ref)[1] : implicitCol;
      implicitCol = c + 1;
      if (!(c >= 0)) return;
      const t = attr(attrs, 't') ?? 'n';
      let raw: string | null = null;
      eachElement(body, 'v', (_, v) => {
        raw = decode(v ?? '');
      });
      let value: CellValue = null;
      switch (t) {
        case 's':
          value = raw === null ? null : strings[Number(raw)] ?? null;
          break;
        case 'inlineStr':
          value = textRuns(body);
          break;
        case 'str':
        case 'e':
        case 'd':
          value = raw;
          break;
        case 'b':
          value = raw === '1';
          break;
        default:
          if (raw !== null && raw !== '') {
            const num = Number(raw);
            const style = Number(attr(attrs, 's') ?? -1);
            value = dateStyles.has(style) ? serialToIso(num, date1904) ?? num : num;
          }
      }
      let formula: string | null = null;
      eachElement(body, 'f', (fAttrs, fBody) => {
        const text = fBody ? decode(fBody) : '';
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
      });
      if (value === null && formula === null) return;
      (grid[r] ??= [])[c] = value;
      if (formula !== null) (formulaGrid[r] ??= [])[c] = '=' + formula;
      if (c < minCol) minCol = c;
      if (c > maxCol) maxCol = c;
    });
  });
  if (maxCol < 0) return { name, columns: [], rows: [], formulas: [], origin: { row: 0, col: 0 } };

  let first = 0;
  while (first < grid.length && !grid[first]) first++;
  const width = maxCol - minCol + 1;
  const dense = <T,>(src: (T | null)[] | undefined): (T | null)[] => {
    const out = new Array<T | null>(width);
    for (let c = 0; c < width; c++) out[c] = src?.[minCol + c] ?? null;
    return out;
  };
  const columns = header
    ? uniqueNames(dense(grid[first]).map((v) => (v === null ? '' : String(v))))
    : Array.from({ length: width }, (_, c) => columnLabel(minCol + c));
  const start = header ? first + 1 : first;
  const rows: CellValue[][] = [];
  // Formulas are always present (sparse) for XLSX, so a formula replaced by
  // a constant counts as a change.
  const formulas: (string | null)[][] = [];
  for (let r = start; r < grid.length; r++) {
    rows.push(dense(grid[r]));
    if (formulaGrid[r]) formulas[r - start] = dense(formulaGrid[r]);
  }
  // Trim trailing empty rows (rows holding only a formula are not empty).
  while (rows.length && !formulas[rows.length - 1] && rows[rows.length - 1].every((v) => v === null)) rows.pop();
  formulas.length = Math.min(formulas.length, rows.length);
  return { name, columns, rows, formulas, origin: { row: start, col: minCol } };
}
