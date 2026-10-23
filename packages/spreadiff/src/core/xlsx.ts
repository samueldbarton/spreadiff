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
import { formulaRefs } from './formula.js';
import { computeMissingValues } from './recalc.js';
import type { CellValue, Table, TableDefinition, Workbook } from './types.js';

/** Options for {@link parseXlsx}. */
export interface ParseXlsxOptions {
  /**
   * Whether each sheet's first non-empty row holds column names.
   * - `'auto'` (default): yes, unless the row has numbers (other than
   *   years next to labels), dates or formulas in it, or a formula
   *   anywhere in the workbook reads one of its cells. Then it's data, as
   *   in most financial models. If it is still mostly labels, they name
   *   the columns; otherwise columns are named by letter.
   * - `true` / `false`: always / never.
   */
  header?: boolean | 'auto';
  /** Only parse these sheets (by name). */
  sheets?: string[];
  /** Name for the resulting workbook (default `workbook.xlsx`). */
  name?: string;
  /**
   * Compute formulas the file has no value for (written by a tool that
   * doesn't calculate, like openpyxl). Default true. See
   * `Workbook.computedValues`.
   */
  computeMissing?: boolean;
}

/**
 * Read an `.xlsx` / `.xlsm` file into a {@link Workbook}.
 *
 * Reads cached values (formulas without one are computed; see
 * `computeMissing`), formula text (shared
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
      f.name.startsWith('xl/worksheets/') ||
      f.name.startsWith('xl/tables/'),
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

  const raws: RawSheet[] = [];
  const tables: TableDefinition[] = [];
  eachElement(workbookXml, 'sheet', (a) => {
    const name = decode(attr(a, 'name') ?? `Sheet${raws.length + 1}`);
    if (options.sheets && !options.sheets.includes(name)) return;
    // The relationships prefix is conventionally `r:` but may be anything.
    const rid = attr(a, 'r:id') ?? /\s[\w.-]+:id\s*=\s*["']([^"']*)/.exec(a)?.[1];
    const path = rid ? rels.get(rid) : undefined;
    if (!path || !files[path]) return;
    raws.push(parseCells(name, text(path), strings, dateStyles, date1904));
    // Excel tables on this sheet, for structured references (Sales[Amount]).
    const relsPath = path.replace(/([^/]+)$/, '_rels/$1.rels');
    eachElement(text(relsPath), 'Relationship', (ra) => {
      const target = attr(ra, 'Target');
      if (!target || !/\/table$/.test(attr(ra, 'Type') ?? '')) return;
      const tablePath = resolvePath(path, target);
      const t = parseTableDefinition(text(tablePath), name);
      if (t) tables.push(t);
    });
  });
  // Defined names (TaxRate → Inputs!$B$2) let formula tracing see through them.
  const names: Record<string, string> = {};
  eachElement(workbookXml, 'definedName', (a, body) => {
    const name = attr(a, 'name');
    // Workbook-scoped only; skip built-ins like _xlnm.Print_Area.
    if (!name || !body || name.startsWith('_xlnm.') || attr(a, 'localSheetId') !== undefined) return;
    names[decode(name)] = decode(body);
  });

  const header = options.header ?? 'auto';
  const headers = header === 'auto' ? detectHeaders(raws, names, tables) : raws.map(() => header);
  const sheets = raws.map((r, i) => buildTable(r, headers[i], header === 'auto'));
  const workbook: Workbook = {
    name: options.name ?? 'workbook.xlsx',
    sheets,
    ...(date1904 ? { date1904 } : {}),
    ...(Object.keys(names).length ? { names } : {}),
    ...(tables.length ? { tables } : {}),
  };
  if (options.computeMissing !== false) {
    const m = computeMissingValues(workbook);
    if (m.computed || m.failed) workbook.computedValues = m;
  }
  return workbook;
}

/** `xl/worksheets/sheet1.xml` + `../tables/table1.xml` → `xl/tables/table1.xml`. */
function resolvePath(from: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  const parts = from.split('/').slice(0, -1);
  for (const seg of target.split('/')) {
    if (seg === '..') parts.pop();
    else if (seg !== '.') parts.push(seg);
  }
  return parts.join('/');
}

function parseTableDefinition(xml: string, sheet: string): TableDefinition | null {
  let def: TableDefinition | null = null;
  eachElement(xml, 'table', (a) => {
    const name = attr(a, 'displayName') ?? attr(a, 'name');
    const ref = attr(a, 'ref');
    if (!name || !ref || def) return;
    const [from, to = from] = ref.split(':');
    const [r0, c0] = decodeRef(from);
    const [r1, c1] = decodeRef(to);
    const headerRows = Number(attr(a, 'headerRowCount') ?? 1);
    const totalsRows = Number(attr(a, 'totalsRowCount') ?? 0);
    def = { name: decode(name), sheet, range: { r0, c0, r1, c1 }, columns: [], headerRows, totalsRows };
  });
  if (!def) return null;
  const d: TableDefinition = def;
  eachElement(xml, 'tableColumn', (a) => {
    d.columns.push(decode(attr(a, 'name') ?? ''));
  });
  return d;
}

/**
 * Decide, per sheet, whether its first row is a header: all text, no
 * formulas, and not read by any formula in the workbook. A model's first
 * row often holds inputs; treating them as column names would hide edits
 * to them and every change they cause.
 */
function detectHeaders(raws: RawSheet[], names: Record<string, string>, tables: TableDefinition[]): boolean[] {
  const candidate = new Map<string, number>(); // sheet name (upper) → index
  const result = raws.map((r, i) => {
    const ok = headerLike(r);
    if (ok) candidate.set(r.name.toUpperCase(), i);
    return ok;
  });
  if (!candidate.size) return result;
  const sheets = raws.map((r) => r.name);
  // Only formulas that could read a candidate row need tokenizing: a
  // reference to sheet row n spells it after a column letter or $ (A1, $B$1,
  // A1:A9) or as a row range (1:1). Names and structured references can
  // hide it, so formulas that may use them are always read. Tokenizing every
  // formula made parsing 2× slower on a 10,000-formula model.
  // A range covering row n starts at n or above, so every row number up to
  // the lowest candidate counts; past 64 of them, read everything.
  const top = Math.max(...[...candidate.values()].map((i) => raws[i].first + 1));
  const rowNumbers = Array.from({ length: top }, (_, i) => i + 1).join('|');
  const spells = top <= 64 ? new RegExp(`(?:[A-Za-z$](?:${rowNumbers})(?!\\d))|(?:(?:^|[^\\w.$])\\$?(?:${rowNumbers}):)`) : /./;
  const nameList = Object.keys(names).map((n) => n.toUpperCase());
  const mayHide = (f: string) => f.includes('[') || (nameList.length > 0 && (nameList.length > 64 || nameList.some((n) => f.toUpperCase().includes(n))));
  for (const r of raws) {
    for (let row = 0; row < r.formulaGrid.length && candidate.size; row++) {
      const fr = r.formulaGrid[row];
      if (!fr) continue;
      for (const f of fr) {
        if (!f) continue;
        if (!spells.test(f) && !mayHide(f)) continue;
        for (const rg of formulaRefs(f, names, { tables, row, sheets }).ranges) {
          const key = (rg.sheet ?? r.name).toUpperCase();
          const i = candidate.get(key);
          if (i === undefined || rg.r0 < 0) continue; // whole columns (A:A) don't single out the header
          const first = raws[i].first;
          if (rg.r0 <= first && first <= rg.r1) {
            result[i] = false;
            candidate.delete(key);
          }
        }
        if (!candidate.size) break;
      }
    }
  }
  return result;
}

const ISO_DATE = /^\d{4}-\d{2}-\d{2}/;

function headerLike(r: RawSheet): boolean {
  if (r.maxCol < 0) return false;
  const row = r.grid[r.first];
  if (r.formulaGrid[r.first]?.some(Boolean)) return false;
  let n = 0;
  for (let c = r.minCol; c <= r.maxCol; c++) {
    const v = row?.[c];
    if (v === null || v === undefined || v === '') continue;
    // Whole years next to labels are column names (Region | 2023 | 2024);
    // a formula reading them still makes the row data (detectHeaders).
    if (isYear(v)) continue;
    // Other numbers, booleans and dates make it a data row.
    if (typeof v !== 'string' || ISO_DATE.test(v)) return false;
    n++;
  }
  return n > 0;
}

const isYear = (v: unknown) => typeof v === 'number' && Number.isInteger(v) && v >= 1900 && v <= 2100;

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
  // XML parsers must read a literal CR LF (or lone CR) as LF; an escaped
  // &#13; survives. Excel writes CR LF, LibreOffice LF, for the same text.
  if (s.indexOf('\r') >= 0) s = s.replace(/\r\n?/g, '\n');
  // Excel escapes control characters in text as _xHHHH_ (OOXML ST_Xstring).
  if (s.indexOf('_x') >= 0) s = s.replace(/_x([0-9A-Fa-f]{4})_/g, (_, h: string) => String.fromCharCode(parseInt(h, 16)));
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
  const colName = (abs: string, letters: string) => {
    if (abs) return abs + letters;
    let c = 0;
    for (const ch of letters) c = c * 26 + (ch.charCodeAt(0) - 64);
    const nc = c - 1 + dCol;
    return nc < 0 ? null : columnLabel(nc);
  };
  const rowNum = (abs: string, digits: string) => {
    if (abs) return abs + digits;
    const nr = Number(digits) + dRow;
    return nr < 1 ? null : String(nr);
  };
  return formula.replace(
    // String literals and quoted sheet names pass through untouched. A ref
    // may follow `!` (Sheet!A1) but not be followed by `!` (it's a sheet name).
    // Whole columns (B:B) and whole rows (2:5) shift too.
    /("[^"]*")|('[^']*')|(?<![A-Za-z0-9_.$])(\$?)([A-Z]{1,3}):(\$?)([A-Z]{1,3})(?![\d(A-Za-z_!])|(?<![A-Za-z0-9_.$:])(\$?)(\d+):(\$?)(\d+)(?![\d(A-Za-z_!.:])|(?<![A-Za-z0-9_.])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\d(A-Za-z_!])/g,
    (all, str, sheet, ca1, c1, ca2, c2, ra1, r1, ra2, r2, cAbs, col, rAbs, row) => {
      if (str || sheet) return all;
      if (c1 !== undefined) {
        const a = colName(ca1, c1);
        const b = colName(ca2, c2);
        return a && b ? `${a}:${b}` : all;
      }
      if (r1 !== undefined) {
        const a = rowNum(ra1, r1);
        const b = rowNum(ra2, r2);
        return a && b ? `${a}:${b}` : all;
      }
      const c = colName(cAbs, col);
      const r = rowNum(rAbs, row);
      if (c === null || r === null) return all;
      return `${c}${r}`;
    },
  );
}

interface RawSheet {
  name: string;
  grid: CellValue[][];
  formulaGrid: (string | null)[][];
  minCol: number;
  maxCol: number;
  /** First non-empty row. */
  first: number;
  /** Array-formula blocks (sheet coordinates). */
  arrays: { r0: number; c0: number; r1: number; c1: number }[];
}

function parseCells(name: string, xml: string, strings: string[], dateStyles: Set<number>, date1904: boolean): RawSheet {
  const grid: CellValue[][] = [];
  const formulaGrid: (string | null)[][] = [];
  const shared = new Map<string, { formula: string; row: number; col: number }>();
  // Array formulas (and dynamic-array spills) store their formula once, on
  // the anchor cell, with the range it fills; data tables likewise.
  const blocks: { ref: string; formula: string; array: boolean }[] = [];
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
        const ft = attr(fAttrs, 't');
        const si = attr(fAttrs, 'si');
        if (ft === 'shared' && si !== undefined) {
          if (text) {
            shared.set(si, { formula: text, row: r, col: c });
            formula = text;
          } else {
            const master = shared.get(si);
            if (master) formula = shiftFormula(master.formula, r - master.row, c - master.col);
          }
        } else if (ft === 'dataTable') {
          // A what-if data table: results of re-running a formula with
          // r1 (and r2) substituted. Its inputs are those cells.
          const r1 = attr(fAttrs, 'r1');
          const r2 = attr(fAttrs, 'r2');
          formula = `TABLE(${[r1, r2].filter(Boolean).join(',')})`;
          const blockRef = attr(fAttrs, 'ref');
          if (blockRef) blocks.push({ ref: blockRef, formula, array: false });
        } else if (text) {
          formula = text;
          const blockRef = attr(fAttrs, 'ref');
          if (ft === 'array' && blockRef) blocks.push({ ref: blockRef, formula: text, array: true });
        }
      });
      // LibreOffice writes a typed error constant as the formula `=#VALUE!`.
      if (formula !== null && /^#(?:VALUE!|N\/A|REF!|DIV\/0!|NUM!|NAME\?|NULL!)$/.test(formula) && value === formula) formula = null;
      if (value === null && formula === null) return;
      (grid[r] ??= [])[c] = value;
      if (formula !== null) (formulaGrid[r] ??= [])[c] = '=' + formula;
      if (c < minCol) minCol = c;
      if (c > maxCol) maxCol = c;
    });
  });
  const arrays: RawSheet['arrays'] = [];
  for (const b of blocks) {
    const [from, to = from] = b.ref.split(':');
    const [r0, c0] = decodeRef(from);
    const [r1, c1] = decodeRef(to);
    if (!(r1 - r0 < 1e6 && c1 - c0 < 2e4)) continue;
    if (b.array && (r1 > r0 || c1 > c0)) arrays.push({ r0, c0, r1, c1 });
    for (let r = r0; r <= r1; r++) {
      const fr = (formulaGrid[r] ??= []);
      for (let c = c0; c <= c1; c++) {
        if (fr[c]) continue;
        fr[c] = '=' + b.formula;
        (grid[r] ??= [])[c] ??= null;
        if (c < minCol) minCol = c;
        if (c > maxCol) maxCol = c;
      }
    }
  }
  let first = 0;
  while (first < grid.length && !grid[first]) first++;
  return { name, grid, formulaGrid, minCol, maxCol, first, arrays };
}

/**
 * A first row of labels that 'auto' kept as data (formulas read it, or a
 * parameter sits among the labels): mostly text, at least two labels.
 */
function labelRow(r: RawSheet): boolean {
  const row = r.grid[r.first];
  let text = 0;
  let filled = 0;
  for (let c = r.minCol; c <= r.maxCol; c++) {
    const v = row?.[c];
    if (v === null || v === undefined || v === '') continue;
    filled++;
    if (typeof v === 'string' && !ISO_DATE.test(v) && !r.formulaGrid[r.first]?.[c]) text++;
  }
  return text >= 2 && text * 2 >= filled;
}

function buildTable(raw: RawSheet, header: boolean, auto = false): Table {
  const { name, grid, formulaGrid, minCol, maxCol, first, arrays } = raw;
  if (maxCol < 0) return { name, columns: [], rows: [], formulas: [], origin: { row: 0, col: 0 } };
  const width = maxCol - minCol + 1;
  const dense = <T,>(src: (T | null)[] | undefined): (T | null)[] => {
    const out = new Array<T | null>(width);
    for (let c = 0; c < width; c++) out[c] = src?.[minCol + c] ?? null;
    return out;
  };
  // A row of labels kept as data still names the columns, so key detection
  // and the grid see "id", "price" rather than A, B: it stays a data row too.
  const columns = header
    ? uniqueNames(dense(grid[first]).map((v) => (v === null ? '' : String(v))))
    : auto && labelRow(raw)
      ? uniqueNames(dense(grid[first]).map((v, c) => (typeof v === 'string' && v.trim() && !formulaGrid[first]?.[minCol + c] ? v : columnLabel(minCol + c))))
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
  return {
    name,
    columns,
    rows,
    formulas,
    origin: { row: start, col: minCol },
    ...(header ? { header: dense(grid[first]) } : {}),
    ...(arrays.length ? { arrays } : {}),
  };
}
