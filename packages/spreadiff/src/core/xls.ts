/**
 * Legacy Excel workbooks (.xls, BIFF8: Excel 97-2003): values, formulas
 * (decoded from their tokens, shared and array formulas included), dates,
 * number formats, merged cells, hidden rows and defined names. Older BIFF
 * versions (Excel 5 and before) aren't read.
 */
import { readCfb } from './cfb.js';
import { decodeFormula, type PtgContext } from './ptg.js';
import { BUILTIN_FORMATS, isDateFormat } from './numfmt.js';
import { serialToIso } from './serial.js';
import { columnLabel } from './csv.js';
import { rawSheet, tablesFromRaw, type RawSheet } from './xlsx.js';
import type { CellValue, Workbook } from './types.js';

/** Options for {@link parseXls}. */
export interface ParseXlsOptions {
  /** Workbook name. Default `workbook.xls`. */
  name?: string;
  /** Whether the first row holds column names. Default `'auto'`. */
  header?: boolean | 'auto';
}

interface Rec {
  type: number;
  data: Uint8Array;
  /** CONTINUE records that followed it. */
  more: Uint8Array[];
}

const ERRORS: Record<number, string> = {
  0x00: '#NULL!',
  0x07: '#DIV/0!',
  0x0f: '#VALUE!',
  0x17: '#REF!',
  0x1d: '#NAME?',
  0x24: '#NUM!',
  0x2a: '#N/A',
};

/**
 * The records of one substream, from its BOF to its EOF. Substreams nested
 * inside it (an embedded chart in a worksheet) are skipped whole: their
 * records aren't this sheet's cells.
 */
function records(stream: Uint8Array, start: number, stopAtEof: boolean): Rec[] {
  const dv = new DataView(stream.buffer, stream.byteOffset, stream.byteLength);
  const out: Rec[] = [];
  let p = start;
  let depth = 0;
  while (p + 4 <= stream.length) {
    const type = dv.getUint16(p, true);
    const len = dv.getUint16(p + 2, true);
    const data = stream.subarray(p + 4, p + 4 + len);
    p += 4 + len;
    if (type === 0x0809 && p - 4 - len > start) depth++;
    if (depth > 0) {
      if (type === 0x000a) depth--;
      continue;
    }
    if (type === 0x003c && out.length) out[out.length - 1].more.push(data);
    else out.push({ type, data, more: [] });
    if (stopAtEof && type === 0x000a) break;
  }
  return out;
}

/** Reads through a record and its CONTINUE records; strings resume with a fresh flags byte after a break. */
class Reader {
  private seg = 0;
  private pos = 0;
  private parts: Uint8Array[];
  constructor(rec: Rec) {
    this.parts = [rec.data, ...rec.more];
  }
  private get cur() {
    return this.parts[this.seg];
  }
  private next() {
    if (this.pos >= this.cur.length && this.seg < this.parts.length - 1) {
      this.seg++;
      this.pos = 0;
    }
  }
  /** True once every byte of the record and its continuations is read. */
  get eof(): boolean {
    return this.seg >= this.parts.length - 1 && this.pos >= this.cur.length;
  }
  u8(): number {
    this.next();
    return this.cur[this.pos++] ?? 0;
  }
  u16(): number {
    return this.u8() | (this.u8() << 8);
  }
  u32(): number {
    return (this.u16() | (this.u16() << 16)) >>> 0;
  }
  /** Skip n bytes, across continuations; a length past the end (a corrupt file) stops at the end. */
  skip(n: number): void {
    while (n > 0 && !this.eof) {
      this.next();
      const step = Math.min(n, this.cur.length - this.pos);
      this.pos += step;
      n -= step;
      if (step === 0) break;
    }
  }
  /** XLUnicodeRichExtendedString: count, flags, (runs), (extension), characters, (run data), (extension data). */
  string(cch = this.u16()): string {
    let flags = this.u8();
    const runs = flags & 0x08 ? this.u16() : 0;
    const ext = flags & 0x04 ? this.u32() : 0;
    let s = '';
    for (let k = 0; k < cch; k++) {
      // A break inside the characters: the next record starts with its own flags byte.
      if (this.pos >= this.cur.length && this.seg < this.parts.length - 1) {
        this.seg++;
        this.pos = 0;
        flags = (flags & ~1) | (this.cur[this.pos++] & 1);
      }
      s += String.fromCharCode(flags & 1 ? this.u16() : this.u8());
    }
    this.skip(runs * 4 + ext);
    return s;
  }
}

/** ShortXLUnicodeString or XLUnicodeStringNoCch from a plain byte array. */
function str(data: Uint8Array, pos: number, cch: number): [string, number] {
  const high = data[pos] & 1;
  let p = pos + 1;
  let s = '';
  for (let k = 0; k < cch; k++) {
    s += String.fromCharCode(high ? data[p] | (data[p + 1] << 8) : data[p]);
    p += high ? 2 : 1;
  }
  return [s, p];
}

/** An RK number: 30 bits of integer or float, maybe divided by 100. */
function rk(v: number): number {
  const n = v & 2 ? v >> 2 : new DataView(new Uint32Array([0, v & 0xfffffffc]).buffer).getFloat64(0, true);
  return v & 1 ? n / 100 : n;
}

/** Parse an .xls (BIFF8) workbook. */
export function parseXls(data: ArrayBuffer | Uint8Array, options: ParseXlsOptions = {}): Workbook {
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
  const streams = readCfb(bytes);
  // Office's encryption wraps the real file (.xlsx, .xlsb) in a compound file; it can't be read without the password.
  if (streams.has('EncryptedPackage') || streams.has('EncryptionInfo')) throw new Error('This workbook is password-protected. Save a copy without a password to compare it');
  const wbs = streams.get('Workbook') ?? streams.get('Book');
  if (!wbs) throw new Error('Not an Excel workbook (no Workbook stream)');
  if (!streams.has('Workbook')) throw new Error("Excel 5 and older (.xls before Excel 97) isn't supported. Save as .xlsx");
  const dv = new DataView(wbs.buffer, wbs.byteOffset, wbs.byteLength);
  const globals = records(wbs, 0, true);
  if (globals[0]?.type !== 0x0809 || dv.getUint16(4, true) !== 0x0600) throw new Error('Only Excel 97-2003 (BIFF8) .xls files are supported. Save as .xlsx');

  let date1904 = false;
  const formats = new Map<number, string>();
  const xfFormat: number[] = [];
  let sst: string[] = [];
  const sheets: { name: string; pos: number; type: number }[] = [];
  const supbooks: { self: boolean; book?: number; sheets: string[]; names: string[] }[] = [];
  let externalBooks = 0;
  const xti: { book: number; first: number; last: number }[] = [];
  const nameRecs: {
    name: string;
    sheet: number;
    rgce: Uint8Array;
    builtin: boolean;
  }[] = [];
  for (const rec of globals) {
    const d = rec.data;
    const v = new DataView(d.buffer, d.byteOffset, d.byteLength);
    // A record too short for what it claims (a damaged file) is skipped, not fatal.
    try {
      switch (rec.type) {
        case 0x0022:
          date1904 = v.getUint16(0, true) === 1;
          break;
        case 0x041e: {
          const id = v.getUint16(0, true);
          formats.set(
            id,
            new Reader({
              type: 0,
              data: d.subarray(2),
              more: rec.more,
            }).string(),
          );
          break;
        }
        case 0x00e0:
          xfFormat.push(v.getUint16(2, true));
          break;
        case 0x00fc: {
          const r = new Reader(rec);
          r.u32();
          const unique = r.u32();
          sst = [];
          // The count can be wrong (or garbage): stop at the end of the data.
          for (let k = 0; k < unique && !r.eof; k++) sst.push(r.string());
          break;
        }
        case 0x0085: {
          const [name] = str(d, 7, d[6]);
          sheets.push({ name, pos: v.getUint32(0, true), type: d[5] });
          break;
        }
        case 0x01ae:
          // SUPBOOK: 0x0401 marks this workbook (self), 0x3A01 add-in functions; otherwise an external workbook.
          {
            const ctab = v.getUint16(0, true);
            const cch = v.getUint16(2, true);
            if (cch === 0x0401 || cch === 0x3a01) supbooks.push({ self: true, sheets: [], names: [] });
            else {
              // Another workbook: its path, then the names of its sheets.
              const r = new Reader({ type: 0, data: d.subarray(4), more: rec.more });
              r.string(cch);
              const sheetNames: string[] = [];
              for (let k = 0; k < ctab && !r.eof; k++) sheetNames.push(r.string());
              supbooks.push({ self: false, book: ++externalBooks, sheets: sheetNames, names: [] });
            }
          }
          break;
        case 0x0023: {
          // EXTERNNAME: an add-in function's or an Excel 2007 function's name (with ptgNameX).
          const [name] = str(d, 7, d[6]);
          supbooks[supbooks.length - 1]?.names.push(name);
          break;
        }
        case 0x0017: {
          const n = v.getUint16(0, true);
          for (let k = 0; k < n; k++)
            xti.push({
              book: v.getUint16(2 + k * 6, true),
              first: v.getInt16(4 + k * 6, true),
              last: v.getInt16(6 + k * 6, true),
            });
          break;
        }
        case 0x0018: {
          const grbit = v.getUint16(0, true);
          const cch = d[3];
          const cce = v.getUint16(4, true);
          const itab = v.getUint16(8, true);
          const builtin = (grbit & 0x20) !== 0;
          const [name, p] = builtin ? [`_xlnm.${d[15]}`, 17] : str(d, 14, cch);
          nameRecs.push({
            name,
            sheet: itab,
            rgce: d.subarray(p, p + cce),
            builtin,
          });
          break;
        }
      }
    } catch (e) {
      if (!(e instanceof RangeError)) throw e;
    }
  }
  const workSheets = sheets.filter((s) => s.type === 0);
  const allNames = sheets.map((s) => s.name);
  const codeOf = (xf: number) => {
    const id = xfFormat[xf] ?? 0;
    return formats.get(id) ?? BUILTIN_FORMATS[id];
  };
  const baseCtx = (row: number, col: number, extra?: Uint8Array): PtgContext => ({
    flavor: 'biff8',
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
      // A name in another workbook is [1]!Name; add-in and newer functions are just their name.
      return book!.self ? name : `[${book!.book}]!${name}`;
    },
  });
  // Defined names, with formulas decoded (relative to A1).
  const names: Record<string, string> = {};
  const localNames = new Map<string, Record<string, string>>();
  for (const n of nameRecs) {
    if (n.builtin || !n.rgce.length) continue;
    const f = decodeFormula(n.rgce, baseCtx(0, 0));
    if (!f) continue;
    if (n.sheet) {
      const sheetName = allNames[n.sheet - 1];
      if (!sheetName) continue;
      let m = localNames.get(sheetName);
      if (!m) localNames.set(sheetName, (m = {}));
      m[n.name] = f.slice(1);
    } else names[n.name] = f.slice(1);
  }

  const raws: RawSheet[] = [];
  for (const s of workSheets) {
    const recs = records(wbs, s.pos, true);
    const grid: CellValue[][] = [];
    const formulaGrid: (string | null)[][] = [];
    const codes: string[] = [''];
    const codeIds = new Map<string, number>();
    const fmtGrid: number[][] = [];
    const merges: RawSheet['merges'] = [];
    const arrays: RawSheet['arrays'] = [];
    const hidden: number[] = [];
    const shared: {
      r0: number;
      c0: number;
      r1: number;
      c1: number;
      rgce: Uint8Array;
      extra: Uint8Array;
    }[] = [];
    const dataTables: { r0: number; c0: number; r1: number; c1: number; formula: string }[] = [];
    const arrayFormulas: {
      r0: number;
      c0: number;
      r1: number;
      c1: number;
      rgce: Uint8Array;
      extra: Uint8Array;
    }[] = [];
    const pending: {
      r: number;
      c: number;
      rgce: Uint8Array;
      extra: Uint8Array;
    }[] = [];
    let stringFor: { r: number; c: number } | null = null;
    const put = (r: number, c: number, value: CellValue, xf: number) => {
      const code = codeOf(xf);
      if (typeof value === 'number' && code && isDateFormat(code)) value = serialToIso(value, date1904) ?? value;
      (grid[r] ??= [])[c] = value;
      if (code && !/^general$/i.test(code) && value !== null && typeof value !== 'boolean') {
        let id = codeIds.get(code);
        if (id === undefined) {
          id = codes.length;
          codes.push(code);
          codeIds.set(code, id);
        }
        (fmtGrid[r] ??= [])[c] = id;
      }
    };
    for (const rec of recs) {
      const d = rec.data;
      const v = new DataView(d.buffer, d.byteOffset, d.byteLength);
      try {
        switch (rec.type) {
          case 0x0208: // ROW
            if (v.getUint16(12, true) & 0x20) hidden.push(v.getUint16(0, true));
            break;
          case 0x0203: // NUMBER
            put(v.getUint16(0, true), v.getUint16(2, true), v.getFloat64(6, true), v.getUint16(4, true));
            break;
          case 0x027e: // RK
            put(v.getUint16(0, true), v.getUint16(2, true), rk(v.getUint32(6, true)), v.getUint16(4, true));
            break;
          case 0x00bd: {
            // MULRK: several RKs in a row
            const r = v.getUint16(0, true);
            const c0 = v.getUint16(2, true);
            const n = (d.length - 6) / 6;
            for (let k = 0; k < n; k++) put(r, c0 + k, rk(v.getUint32(6 + k * 6, true)), v.getUint16(4 + k * 6, true));
            break;
          }
          case 0x00fd: // LABELSST
            put(v.getUint16(0, true), v.getUint16(2, true), sst[v.getUint32(6, true)] ?? '', v.getUint16(4, true));
            break;
          case 0x0204: {
            // LABEL (older writers)
            const r = new Reader({
              type: 0,
              data: d.subarray(6),
              more: rec.more,
            });
            put(v.getUint16(0, true), v.getUint16(2, true), r.string(), v.getUint16(4, true));
            break;
          }
          case 0x0205: {
            // BOOLERR
            const isErr = d[7] === 1;
            put(v.getUint16(0, true), v.getUint16(2, true), isErr ? (ERRORS[d[6]] ?? '#N/A') : d[6] === 1, v.getUint16(4, true));
            break;
          }
          case 0x0006: {
            // FORMULA: cached result, then tokens.
            const r = v.getUint16(0, true);
            const c = v.getUint16(2, true);
            const xf = v.getUint16(4, true);
            let value: CellValue;
            if (v.getUint16(12, true) === 0xffff) {
              const kind = d[6];
              if (kind === 0) {
                value = '';
                stringFor = { r, c }; // the text is in the STRING record that follows
              } else if (kind === 1) value = d[8] === 1;
              else if (kind === 2) value = ERRORS[d[8]] ?? '#N/A';
              else value = '';
            } else value = v.getFloat64(6, true);
            put(r, c, value, xf);
            const cce = v.getUint16(20, true);
            const rgce = d.subarray(22, 22 + cce);
            pending.push({ r, c, rgce, extra: d.subarray(22 + cce) });
            break;
          }
          case 0x0207: // STRING: a formula's text result
            if (stringFor) {
              (grid[stringFor.r] ??= [])[stringFor.c] = new Reader(rec).string();
              stringFor = null;
            }
            break;
          case 0x04bc: {
            // SHRFMLA: a formula shared by a block of cells, with references relative to each.
            const cce = v.getUint16(8, true);
            shared.push({
              r0: v.getUint16(0, true),
              r1: v.getUint16(2, true),
              c0: d[4],
              c1: d[5],
              rgce: d.subarray(10, 10 + cce),
              extra: d.subarray(10 + cce),
            });
            break;
          }
          case 0x0221: {
            // ARRAY: an array formula over a block.
            const cce = v.getUint16(12, true);
            arrayFormulas.push({
              r0: v.getUint16(0, true),
              r1: v.getUint16(2, true),
              c0: d[4],
              c1: d[5],
              rgce: d.subarray(14, 14 + cce),
              extra: d.subarray(14 + cce),
            });
            break;
          }
          case 0x0236: {
            // TABLE: a what-if data table; its cells hold TABLE(input cell(s)), as the .xlsx reader writes them.
            const flags = v.getUint16(6, true);
            const ref = (r: number, c: number) => `${columnLabel(c)}${r + 1}`;
            const inputs = [ref(v.getUint16(8, true), v.getUint16(10, true))];
            if (flags & 0x08) inputs.push(ref(v.getUint16(12, true), v.getUint16(14, true)));
            dataTables.push({ r0: v.getUint16(0, true), r1: v.getUint16(2, true), c0: d[4], c1: d[5], formula: `=TABLE(${inputs.join(',')})` });
            break;
          }
          case 0x00e5: {
            // MERGECELLS
            const n = v.getUint16(0, true);
            for (let k = 0; k < n; k++) {
              const o = 2 + k * 8;
              merges.push({
                r0: v.getUint16(o, true),
                r1: v.getUint16(o + 2, true),
                c0: v.getUint16(o + 4, true),
                c1: v.getUint16(o + 6, true),
              });
            }
            break;
          }
        }
      } catch (e) {
        if (!(e instanceof RangeError)) throw e;
      }
    }
    // Formulas: a lone ptgExp points at the shared or array formula it belongs to.
    for (const p of pending) {
      let f: string | null;
      if (p.rgce[0] === 0x02 && p.rgce.length === 5) {
        f = dataTables.find((t) => p.r >= t.r0 && p.r <= t.r1 && p.c >= t.c0 && p.c <= t.c1)?.formula ?? null;
      } else if (p.rgce[0] === 0x01 && p.rgce.length === 5) {
        const ar = p.rgce[1] | (p.rgce[2] << 8);
        const ac = p.rgce[3] | (p.rgce[4] << 8);
        const arr = arrayFormulas.find((a) => a.r0 === ar && a.c0 === ac && p.r >= a.r0 && p.r <= a.r1 && p.c >= a.c0 && p.c <= a.c1);
        const sh = arr ? undefined : shared.find((x) => p.r >= x.r0 && p.r <= x.r1 && p.c >= x.c0 && p.c <= x.c1);
        if (arr) {
          f = decodeFormula(arr.rgce, baseCtx(arr.r0, arr.c0, arr.extra));
          if (!arrays.some((a) => a.r0 === arr.r0 && a.c0 === arr.c0)) arrays.push({ r0: arr.r0, c0: arr.c0, r1: arr.r1, c1: arr.c1 });
        } else f = sh ? decodeFormula(sh.rgce, baseCtx(p.r, p.c, sh.extra)) : null;
      } else f = decodeFormula(p.rgce, baseCtx(p.r, p.c, p.extra));
      if (f) (formulaGrid[p.r] ??= [])[p.c] = f;
    }
    raws.push(
      rawSheet(s.name, grid, formulaGrid, {
        merges,
        arrays,
        hidden,
        formats: codes.length > 1 ? { codes, grid: fmtGrid, date1904 } : undefined,
      }),
    );
  }
  const tables = tablesFromRaw(raws, options.header ?? 'auto', names, [], localNames);
  return {
    name: options.name ?? 'workbook.xls',
    sheets: tables,
    ...(date1904 ? { date1904 } : {}),
    ...(Object.keys(names).length ? { names } : {}),
  };
}
