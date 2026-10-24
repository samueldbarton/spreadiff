/**
 * A small, fast XLSX reader: cached values, formulas and dates. It does not
 * evaluate formulas or keep styling; for diffing, the stored values are what
 * you want. The reader is a few KB on top of fflate; a full SheetJS build is ~900 KB.
 *
 * Elements are found with `indexOf`, never with backtracking regexes, so a
 * malformed or hostile file costs linear time.
 */
import { unzipSync, strFromU8 } from 'fflate';
import { columnLabel } from './csv.js';
import { computeMissingValues } from './recalc.js';
import { BUILTIN_FORMATS } from './numfmt.js';
import { attr, decode, decodeRef, eachElement, removeElements, resolvePath, textRuns } from './xml.js';
import { parseObjects } from './objects.js';
import { serialToIso } from './serial.js';
import { rawSheet, tablesFromRaw, type RawSheet } from './sheet.js';

export { applyDecisionsToXlsx, type SkippedDecision } from './xlsx-write.js';
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
 * defined names (workbook- and sheet-scoped), merged ranges, and the indent
 * and bold of text cells. Each sheet is trimmed to its used range; `origin`
 * records where it starts. Other styles, number formats and charts are
 * ignored.
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
      f.name.startsWith('xl/tables/') ||
      // Charts and pivot tables (see parseObjects): their links and definitions, not their data.
      f.name.startsWith('xl/chartsheets/') ||
      f.name.startsWith('xl/drawings/_rels/') ||
      /^xl\/charts\/chart\d+\.xml$/.test(f.name) ||
      f.name.startsWith('xl/pivotTables/') ||
      /^xl\/pivotCache\/pivotCacheDefinition\d+\.xml$/.test(f.name),
  });
  // Part names are case-insensitive (OPC): some writers name a sheet Sheet1.xml and its links sheet1.xml.rels.
  let lower: Map<string, string> | null = null;
  const part = (path: string): Uint8Array | undefined => {
    if (files[path]) return files[path];
    lower ??= new Map(Object.keys(files).map((k) => [k.toLowerCase(), k]));
    const k = lower.get(path.toLowerCase());
    return k ? files[k] : undefined;
  };
  const text = (path: string) => {
    const f = part(path);
    return f ? stripPrefixes(strFromU8(f)) : '';
  };

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
  const stylesXml = text('xl/styles.xml');
  const dateStyles = parseDateStyles(stylesXml);
  const labelStyles = parseLabelStyles(stylesXml);
  const numberFormats = parseNumberFormats(stylesXml);

  const raws: RawSheet[] = [];
  const sheetParts: { name: string; path: string }[] = [];
  const tables: TableDefinition[] = [];
  const allSheets: string[] = []; // every sheet in workbook order: localSheetId indexes this
  eachElement(workbookXml, 'sheet', (a) => {
    const name = decode(attr(a, 'name') ?? `Sheet${raws.length + 1}`);
    allSheets.push(name);
    if (options.sheets && !options.sheets.includes(name)) return;
    // The relationships prefix is conventionally `r:` but may be anything.
    const rid = attr(a, 'r:id') ?? /\s[\w.-]+:id\s*=\s*["']([^"']*)/.exec(a)?.[1];
    const path = rid ? rels.get(rid) : undefined;
    if (!path || !part(path)) return;
    raws.push(parseCells(name, text(path), strings, dateStyles, date1904, labelStyles, numberFormats));
    sheetParts.push({ name, path });
    // Excel tables on this sheet, for structured references (Sales[Amount]).
    const relsPath = path.replace(/([^/]+)$/, '_rels/$1.rels');
    eachElement(text(relsPath), 'Relationship', (ra) => {
      const target = attr(ra, 'Target');
      if (!target || !/\/table$/.test(attr(ra, 'Type') ?? '')) return;
      const tablePath = resolvePath(path, target);
      const tableXml = text(tablePath);
      const t = parseTableDefinition(tableXml, name);
      if (t) tables.push(t);
      const raw = raws[raws.length - 1];
      if (raw.hidden.length) raw.filters.push(...activeFilters(tableXml));
    });
  });
  // Defined names (TaxRate → Inputs!$B$2) let formula tracing see through them.
  const names: Record<string, string> = {};
  const localNames = new Map<string, Record<string, string>>(); // sheet → its own names
  eachElement(workbookXml, 'definedName', (a, body) => {
    const name = attr(a, 'name');
    // Skip built-ins like _xlnm.Print_Area.
    if (!name || !body || name.startsWith('_xlnm.')) return;
    const local = attr(a, 'localSheetId');
    if (local === undefined) names[decode(name)] = decode(body);
    else {
      const sheet = allSheets[Number(local)];
      if (sheet === undefined) return;
      let m = localNames.get(sheet);
      if (!m) localNames.set(sheet, (m = {}));
      m[decode(name)] = decode(body);
    }
  });

  const sheets = tablesFromRaw(raws, options.header ?? 'auto', names, tables, localNames);
  const workbook: Workbook = {
    name: options.name ?? 'workbook.xlsx',
    sheets,
    ...(date1904 ? { date1904 } : {}),
    ...(Object.keys(names).length ? { names } : {}),
    ...(tables.length ? { tables } : {}),
  };
  const { charts, pivots } = parseObjects(text, sheetParts);
  if (charts.length) workbook.charts = charts;
  if (pivots.length) workbook.pivots = pivots;
  if (options.computeMissing !== false) {
    const m = computeMissingValues(workbook);
    if (m.computed || m.failed) workbook.computedValues = m;
  }
  return workbook;
}

/** `xl/worksheets/sheet1.xml` + `../tables/table1.xml` → `xl/tables/table1.xml`. */
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

function stripPrefixes(xml: string): string {
  if (!/<[A-Za-z_][\w.-]*:(?:workbook|worksheet|sst|styleSheet|Relationships)\b/.test(xml.slice(0, 4096))) return xml;
  return xml.replace(/<(\/?)[A-Za-z_][\w.-]*:/g, '<$1');
}

function parseSharedStrings(xml: string): string[] {
  const out: string[] = [];
  eachElement(xml, 'si', (_, body) => out.push(body ? textRuns(body) : ''));
  return out;
}

// 27-36 and 50-58: the built-in dates and times of East Asian Excel.
const BUILTIN_DATE_FORMATS = new Set([14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 45, 46, 47, 50, 51, 52, 53, 54, 55, 56, 57, 58]);

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

/** cellXfs index → number format code, for styles with a format other than General. */
function parseNumberFormats(xml: string): Map<number, string> {
  const custom = new Map<number, string>();
  eachElement(xml, 'numFmt', (a) => {
    custom.set(Number(attr(a, 'numFmtId')), decode(attr(a, 'formatCode') ?? ''));
  });
  const out = new Map<number, string>();
  eachElement(xml, 'cellXfs', (_, body) => {
    let i = 0;
    eachElement(body ?? '', 'xf', (a) => {
      const id = Number(attr(a, 'numFmtId') ?? 0);
      const code = custom.get(id) ?? BUILTIN_FORMATS[id];
      if (code && !/^general$/i.test(code)) out.set(i, code);
      i++;
    });
  });
  return out;
}

/** cellXfs index → indent and bold, for the styles that have either. */
function parseLabelStyles(xml: string): Map<number, { indent?: number; bold?: boolean }> {
  const bold: boolean[] = [];
  eachElement(xml, 'fonts', (_, body) => {
    eachElement(body ?? '', 'font', (_a, fb) => {
      const b = /<b(?:\s+val\s*=\s*["']([^"']*)["'])?\s*\/?>/.exec(fb ?? '');
      bold.push(!!b && b[1] !== '0' && b[1] !== 'false');
    });
  });
  const out = new Map<number, { indent?: number; bold?: boolean }>();
  eachElement(xml, 'cellXfs', (_, body) => {
    let i = 0;
    eachElement(body ?? '', 'xf', (a, xb) => {
      const indent = Number(/<alignment\b[^>]*\sindent\s*=\s*["'](\d+)["']/.exec(xb ?? '')?.[1] ?? 0);
      const b = bold[Number(attr(a, 'fontId') ?? 0)] ?? false;
      if (indent > 0 || b) out.set(i, { ...(indent > 0 ? { indent } : {}), ...(b ? { bold: true } : {}) });
      i++;
    });
  });
  return out;
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

function activeFilters(xml: string): { r0: number; r1: number }[] {
  const out: { r0: number; r1: number }[] = [];
  eachElement(xml, 'autoFilter', (a, body) => {
    const ref = attr(a, 'ref');
    if (!ref || !body || !/<(?:\w+:)?filterColumn\b/.test(body)) return;
    const [from, to = from] = ref.split(':');
    out.push({ r0: decodeRef(from)[0] + 1, r1: decodeRef(to)[0] });
  });
  return out;
}

function parseCells(
  name: string,
  xml: string,
  strings: string[],
  dateStyles: Set<number>,
  date1904: boolean,
  labelStyles: Map<number, { indent?: number; bold?: boolean }>,
  numberFormats: Map<number, string>,
): RawSheet {
  const styled: NonNullable<Table['labelStyles']> = [];
  // Format codes, numbered from 1 in order of first use; per cell, the code's number.
  const codes: string[] = [''];
  const codeIds = new Map<string, number>();
  const fmtGrid: number[][] = [];
  const grid: CellValue[][] = [];
  const formulaGrid: (string | null)[][] = [];
  const shared = new Map<string, { formula: string; row: number; col: number }>();
  // Array formulas (and dynamic-array spills) store their formula once, on
  // the anchor cell, with the range it fills; data tables likewise.
  const blocks: { ref: string; formula: string; array: boolean }[] = [];
  let minCol = Infinity;
  let maxCol = -1;
  let implicitRow = 0;
  const hidden: number[] = [];
  eachElement(xml, 'row', (rowAttrs, rowBody) => {
    const rAttr = attr(rowAttrs, 'r');
    const r = rAttr ? Number(rAttr) - 1 : implicitRow;
    implicitRow = r + 1;
    const h = attr(rowAttrs, 'hidden');
    if ((h === '1' || h === 'true') && r >= 0) hidden.push(r);
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
      if (typeof value === 'string' && labelStyles.size) {
        const ls = labelStyles.get(Number(attr(attrs, 's') ?? 0));
        if (ls) styled.push({ r, c, ...ls });
      }
      (grid[r] ??= [])[c] = value;
      if (formula !== null) (formulaGrid[r] ??= [])[c] = '=' + formula;
      if (numberFormats.size && value !== null && typeof value !== 'boolean') {
        const code = numberFormats.get(Number(attr(attrs, 's') ?? 0));
        // Text cells only matter with a text section (`;;;@` or `"Re: "@`).
        if (code && (typeof value !== 'string' || dateStyles.has(Number(attr(attrs, 's'))) || code.includes('@'))) {
          let id = codeIds.get(code);
          if (id === undefined) {
            id = codes.length;
            if (id > 0xffff) id = 0; // more distinct codes than a Uint16 holds: the rest show as General
            else {
              codes.push(code);
              codeIds.set(code, id);
            }
          }
          if (id) (fmtGrid[r] ??= [])[c] = id;
        }
      }
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
  const merges: RawSheet['merges'] = [];
  eachElement(xml, 'mergeCell', (a) => {
    const ref = attr(a, 'ref');
    if (!ref || !ref.includes(':')) return;
    const [from, to] = ref.split(':');
    const [r0, c0] = decodeRef(from);
    const [r1, c1] = decodeRef(to);
    if (r1 > r0 || c1 > c0) merges.push({ r0: Math.min(r0, r1), c0: Math.min(c0, c1), r1: Math.max(r0, r1), c1: Math.max(c0, c1) });
  });
  let first = 0;
  while (first < grid.length && !grid[first]) first++;
  return {
    name,
    grid,
    formulaGrid,
    minCol,
    maxCol,
    first,
    arrays,
    merges,
    labelStyles: styled,
    formats: codes.length > 1 ? { codes, grid: fmtGrid, date1904 } : undefined,
    hidden,
    filters: hidden.length ? activeFilters(xml) : [],
  };
}
