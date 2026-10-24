/**
 * Excel binary workbooks (.xlsb, BIFF12): the .xlsx package layout with
 * binary parts. Values, formulas (decoded from their tokens, shared and
 * array formulas included), dates, number formats, merged cells, hidden and
 * filtered rows and defined names. Structured references to tables aren't
 * decoded: those cells keep their values, without a formula.
 */
import { strFromU8, unzipSync } from 'fflate';
import { decodeFormula, type PtgContext } from './ptg.js';
import { BUILTIN_FORMATS, isDateFormat } from './numfmt.js';
import { serialToIso } from './serial.js';
import { attr, eachElement, resolvePath } from './xml.js';
import { rawSheet, tablesFromRaw, type RawSheet } from './xlsx.js';
import type { CellValue, Workbook } from './types.js';

/** Options for {@link parseXlsb}. */
export interface ParseXlsbOptions {
  /** Workbook name. Default `workbook.xlsb`. */
  name?: string;
  /** Whether the first row holds column names. Default `'auto'`. */
  header?: boolean | 'auto';
}

const ERRORS: Record<number, string> = { 0x00: '#NULL!', 0x07: '#DIV/0!', 0x0f: '#VALUE!', 0x17: '#REF!', 0x1d: '#NAME?', 0x24: '#NUM!', 0x2a: '#N/A', 0x2b: '#GETTING_DATA' };

/** Each record of a part: a 1–2 byte type and a 1–4 byte length, 7 bits per byte. */
function eachRecord(part: Uint8Array, fn: (type: number, r: Reader) => void): void {
  let p = 0;
  while (p < part.length) {
    let type = part[p++];
    if (type & 0x80) type = (type & 0x7f) | ((part[p++] & 0x7f) << 7);
    let len = 0;
    for (let k = 0; k < 4; k++) {
      const b = part[p++];
      len |= (b & 0x7f) << (7 * k);
      if (!(b & 0x80)) break;
    }
    const end = Math.min(p + len, part.length);
    try {
      fn(type, new Reader(part.subarray(p, end)));
    } catch (e) {
      // A record too short for what it claims (a damaged file) is skipped, not fatal.
      if (!(e instanceof RangeError)) throw e;
    }
    p = end;
  }
}

class Reader {
  pos = 0;
  readonly dv: DataView;
  constructor(readonly data: Uint8Array) {
    this.dv = new DataView(data.buffer, data.byteOffset, data.byteLength);
  }
  get left(): number {
    return this.data.length - this.pos;
  }
  u8(): number {
    const v = this.dv.getUint8(this.pos);
    this.pos += 1;
    return v;
  }
  u16(): number {
    const v = this.dv.getUint16(this.pos, true);
    this.pos += 2;
    return v;
  }
  u32(): number {
    const v = this.dv.getUint32(this.pos, true);
    this.pos += 4;
    return v;
  }
  i32(): number {
    const v = this.dv.getInt32(this.pos, true);
    this.pos += 4;
    return v;
  }
  f64(): number {
    const v = this.dv.getFloat64(this.pos, true);
    this.pos += 8;
    return v;
  }
  bytes(n: number): Uint8Array {
    if (n > this.left) throw new RangeError('past the record');
    const b = this.data.subarray(this.pos, this.pos + n);
    this.pos += n;
    return b;
  }
  /** XLWideString: a 32-bit count of UTF-16 characters (0xFFFFFFFF: none). */
  str(): string {
    const n = this.u32();
    if (n === 0xffffffff) return '';
    if (n * 2 > this.left) throw new RangeError('past the record');
    let s = '';
    for (let k = 0; k < n; k++) s += String.fromCharCode(this.dv.getUint16(this.pos + k * 2, true));
    this.pos += n * 2;
    return s;
  }
  /** A formula: its tokens, then the extra data (array constants) they use. */
  formula(): { rgce: Uint8Array; extra: Uint8Array } {
    const rgce = this.bytes(this.u32());
    const extra = this.left >= 4 ? this.bytes(Math.min(this.u32(), this.left)) : new Uint8Array(0);
    return { rgce, extra };
  }
  /** RfX: first and last row, first and last column. */
  range(): { r0: number; r1: number; c0: number; c1: number } {
    const r0 = this.u32();
    const r1 = this.u32();
    const c0 = this.u32();
    const c1 = this.u32();
    return { r0, r1, c0, c1 };
  }
}

/** An RK number: 30 bits of integer or float, maybe divided by 100. */
function rk(v: number): number {
  const n = v & 2 ? v >> 2 : new DataView(new Uint32Array([0, v & 0xfffffffc]).buffer).getFloat64(0, true);
  return v & 1 ? n / 100 : n;
}

/** Relationship targets of a part, by id. */
function relsOf(files: Record<string, Uint8Array>, part: string): Map<string, string> {
  const out = new Map<string, string>();
  const rels = files[part.replace(/([^/]+)$/, '_rels/$1.rels')];
  if (!rels) return out;
  eachElement(strFromU8(rels), 'Relationship', (a) => {
    const id = attr(a, 'Id');
    const target = attr(a, 'Target');
    if (id && target && attr(a, 'TargetMode') !== 'External') out.set(id, target.startsWith('/') ? target.slice(1) : resolvePath(part, target));
  });
  return out;
}

/** Parse an .xlsb (BIFF12) workbook. */
export function parseXlsb(data: ArrayBuffer | Uint8Array, options: ParseXlsbOptions = {}): Workbook {
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
  const all = unzipSync(bytes, { filter: (f) => f.name.startsWith('xl/') && !/\.(?:png|jpe?g|gif|emf|wmf|vml)$/i.test(f.name) && !f.name.includes('vbaProject') });
  // Part names are case-insensitive (OPC).
  const files: Record<string, Uint8Array> = {};
  for (const [k, v] of Object.entries(all)) files[k.toLowerCase()] = v;
  const get = (p: string) => files[p.toLowerCase()];
  const wbPart = get('xl/workbook.bin');
  if (!wbPart) throw new Error('Not an Excel binary workbook (no xl/workbook.bin)');
  const wbRels = relsOf(files, 'xl/workbook.bin');

  let date1904 = false;
  const sheets: { name: string; path?: string }[] = [];
  const nameRecs: { name: string; sheet: number; rgce: Uint8Array; extra: Uint8Array; skip: boolean }[] = [];
  const supbooks: { self: boolean; book?: number; sheets: string[]; names: string[] }[] = [];
  let externalBooks = 0;
  const xti: { book: number; first: number; last: number }[] = [];
  eachRecord(wbPart, (type, r) => {
    switch (type) {
      case 0x99: // BrtWbProp
        date1904 = (r.u32() & 1) === 1;
        break;
      case 0x9c: {
        // BrtBundleSh: a sheet's state, tab id, relationship id and name
        r.u32();
        r.u32();
        const rid = r.str();
        sheets.push({ name: r.str(), path: wbRels.get(rid) });
        break;
      }
      case 0x27: {
        // BrtName
        const flags = r.u32();
        r.u8();
        const itab = r.u32();
        const name = r.str();
        const { rgce, extra } = r.formula();
        // Built-in names (print areas, filter ranges), macros and future-function placeholders aren't the workbook's own.
        nameRecs.push({ name, sheet: itab === 0xffffffff ? -1 : itab, rgce, extra, skip: (flags & 0x2e) !== 0 || /^_xl(nm|fn|ws)\./i.test(name) });
        break;
      }
      case 0x165: // BrtSupSelf
      case 0x166: // BrtSupSame
      case 0x29b: // BrtSupAddin
        supbooks.push({ self: true, sheets: [], names: [] });
        break;
      case 0x163: // BrtSupBookSrc: another workbook
        supbooks.push({ self: false, book: ++externalBooks, sheets: [], names: [] });
        break;
      case 0x167: {
        // BrtSupTabs: that workbook's sheet names
        const book = supbooks[supbooks.length - 1];
        const n = r.u32();
        for (let k = 0; k < n && r.left >= 4; k++) book?.sheets.push(r.str());
        break;
      }
      case 0x169: // BrtPlaceholderName: an external or add-in name
        supbooks[supbooks.length - 1]?.names.push(r.str());
        break;
      case 0x16a: {
        // BrtExternSheet
        const n = r.u32();
        for (let k = 0; k < n && r.left >= 12; k++) xti.push({ book: r.u32(), first: r.i32(), last: r.i32() });
        break;
      }
    }
  });
  const allNames = sheets.map((s) => s.name);

  // Shared strings and number formats.
  const sst: string[] = [];
  const sstPart = get('xl/sharedStrings.bin');
  if (sstPart)
    eachRecord(sstPart, (type, r) => {
      if (type !== 0x13) return; // BrtSSTItem: flags, text, (runs)
      r.u8();
      sst.push(r.str());
    });
  const formats = new Map<number, string>();
  const xfFormat: number[] = [];
  const stylesPart = get('xl/styles.bin');
  if (stylesPart) {
    let inCellXfs = false;
    eachRecord(stylesPart, (type, r) => {
      if (type === 0x269) inCellXfs = true;
      else if (type === 0x26a) inCellXfs = false;
      else if (type === 0x2c) {
        const id = r.u16();
        formats.set(id, r.str());
      } else if (type === 0x2f && inCellXfs) {
        r.u16();
        xfFormat.push(r.u16());
      }
    });
  }
  const codeOf = (xf: number) => {
    const id = xfFormat[xf] ?? 0;
    return formats.get(id) ?? BUILTIN_FORMATS[id];
  };

  const baseCtx = (row: number, col: number, extra?: Uint8Array): PtgContext => ({
    flavor: 'biff12',
    row,
    col,
    extra,
    sheetOf: (i) => {
      const x = xti[i];
      const book = x && supbooks[x.book];
      if (!book || x.first < 0) return null;
      const list = book.self ? allNames : book.sheets;
      const first = list[x.first];
      if (first === undefined) return null;
      return { first, last: list[x.last] ?? first, ...(book.self ? {} : { book: book.book }) };
    },
    nameOf: (i) => nameRecs[i - 1]?.name ?? null,
    nameXOf: (i, n) => {
      const book = supbooks[xti[i]?.book ?? -1];
      const name = book?.names[n - 1];
      if (name === undefined) return null;
      return book!.self ? name : `[${book!.book}]!${name}`;
    },
  });
  const names: Record<string, string> = {};
  const localNames = new Map<string, Record<string, string>>();
  for (const n of nameRecs) {
    if (n.skip || !n.rgce.length) continue;
    const f = decodeFormula(n.rgce, baseCtx(0, 0, n.extra));
    if (!f) continue;
    if (n.sheet >= 0) {
      const sheetName = allNames[n.sheet];
      if (!sheetName) continue;
      let m = localNames.get(sheetName);
      if (!m) localNames.set(sheetName, (m = {}));
      m[n.name] = f.slice(1);
    } else names[n.name] = f.slice(1);
  }

  const raws: RawSheet[] = [];
  for (const s of sheets) {
    // Chart sheets and macro sheets have no cells to read.
    const part = s.path && /worksheets\//i.test(s.path) ? get(s.path) : undefined;
    if (!part) continue;
    const grid: CellValue[][] = [];
    const formulaGrid: (string | null)[][] = [];
    const codes: string[] = [''];
    const codeIds = new Map<string, number>();
    const fmtGrid: number[][] = [];
    const merges: RawSheet['merges'] = [];
    const arrays: RawSheet['arrays'] = [];
    const hidden: number[] = [];
    const filters: RawSheet['filters'] = [];
    type Block = { r0: number; r1: number; c0: number; c1: number; rgce: Uint8Array; extra: Uint8Array };
    const shared: Block[] = [];
    const arrayFormulas: Block[] = [];
    const pending: { r: number; c: number; rgce: Uint8Array; extra: Uint8Array }[] = [];
    let row = 0;
    let filter: { r0: number; r1: number } | null = null;
    const put = (c: number, value: CellValue, xf: number) => {
      const code = codeOf(xf);
      if (typeof value === 'number' && code && isDateFormat(code)) value = serialToIso(value, date1904) ?? value;
      (grid[row] ??= [])[c] = value;
      if (code && !/^general$/i.test(code) && value !== null && typeof value !== 'boolean') {
        let id = codeIds.get(code);
        if (id === undefined) {
          id = codes.length;
          codes.push(code);
          codeIds.set(code, id);
        }
        (fmtGrid[row] ??= [])[c] = id;
      }
    };
    eachRecord(part, (type, r) => {
      if (type === 0x00) {
        // BrtRowHdr: row number, style, height, flags (0x10 in the 12th byte: hidden)
        row = r.u32();
        if (r.left >= 8 && r.data[11] & 0x10) hidden.push(row);
        return;
      }
      if ((type >= 0x02 && type <= 0x0b) || type === 0x3e) {
        // A cell: column, style (24 bits) and flags, then its value.
        const c = r.u32();
        const xf = r.u16() | (r.u8() << 16);
        r.u8();
        let value: CellValue;
        switch (type) {
          case 0x02:
            value = rk(r.u32());
            break;
          case 0x03:
          case 0x0b:
            value = ERRORS[r.u8()] ?? '#N/A';
            break;
          case 0x04:
          case 0x0a:
            value = r.u8() === 1;
            break;
          case 0x05:
          case 0x09:
            value = r.f64();
            break;
          case 0x07:
            value = sst[r.u32()] ?? '';
            break;
          case 0x3e:
            r.u8();
            value = r.str();
            break;
          default: // 0x06 BrtCellSt, 0x08 BrtFmlaString
            value = r.str();
        }
        put(c, value, xf);
        if (type >= 0x08) {
          if (type === 0x3e) return;
          r.u16(); // flags
          pending.push({ r: row, c, ...r.formula() });
        }
        return;
      }
      switch (type) {
        case 0x1aa: {
          // BrtArrFmla: an array formula over a block
          const range = r.range();
          r.u8();
          arrayFormulas.push({ ...range, ...r.formula() });
          break;
        }
        case 0x1ab: // BrtShrFmla: a formula shared by a block, with references relative to each cell
          shared.push({ ...r.range(), ...r.formula() });
          break;
        case 0xb0: // BrtMergeCell
          merges.push(r.range());
          break;
        case 0xa1: {
          // BrtBeginAFilter: active only if a filter column follows
          const { r0, r1 } = r.range();
          filter = { r0: r0 + 1, r1 };
          break;
        }
        case 0xa3: // BrtBeginFilterColumn
          if (filter && !filters.includes(filter)) filters.push(filter);
          break;
      }
    });
    // Formulas: a lone ptgExp (its master cell's row) points at the shared or array formula the cell belongs to.
    for (const p of pending) {
      let f: string | null;
      if (p.rgce[0] === 0x01 && p.rgce.length === 5) {
        const inside = (b: Block) => p.r >= b.r0 && p.r <= b.r1 && p.c >= b.c0 && p.c <= b.c1;
        const arr = arrayFormulas.find(inside);
        const sh = arr ? undefined : shared.find(inside);
        if (arr) {
          f = decodeFormula(arr.rgce, baseCtx(arr.r0, arr.c0, arr.extra));
          if (!arrays.some((a) => a.r0 === arr.r0 && a.c0 === arr.c0)) arrays.push({ r0: arr.r0, c0: arr.c0, r1: arr.r1, c1: arr.c1 });
        } else f = sh ? decodeFormula(sh.rgce, baseCtx(p.r, p.c, sh.extra)) : null;
      } else f = decodeFormula(p.rgce, baseCtx(p.r, p.c, p.extra));
      if (f) (formulaGrid[p.r] ??= [])[p.c] = f;
    }
    raws.push(rawSheet(s.name, grid, formulaGrid, { merges, arrays, hidden, filters, formats: codes.length > 1 ? { codes, grid: fmtGrid, date1904 } : undefined }));
  }
  const tables = tablesFromRaw(raws, options.header ?? 'auto', names, [], localNames);
  return {
    name: options.name ?? 'workbook.xlsb',
    sheets: tables,
    ...(date1904 ? { date1904 } : {}),
    ...(Object.keys(names).length ? { names } : {}),
  };
}
