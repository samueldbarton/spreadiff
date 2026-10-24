/**
 * Formula bytecode (parsed tokens, "ptgs") to formula text, for the .xls
 * (BIFF8) and .xlsb (BIFF12) readers. Formulas are stored in reverse
 * Polish order with explicit parentheses, so the text is rebuilt with a
 * stack. A token this doesn't know makes the whole formula undecodable
 * (null): the cell keeps its value, and no formula is guessed.
 */
import { columnLabel } from './csv.js';
import { FIXED_ARITY, FTAB, FTAB_XLSB } from './ftab.js';

/** What a formula's tokens refer to, from the workbook around it. */
export interface PtgContext {
  /** `biff8` (.xls) or `biff12` (.xlsb): they encode references differently. */
  flavor: 'biff8' | 'biff12';
  /** The formula's own cell (0-based), for the relative references of shared formulas. */
  row: number;
  col: number;
  /** A 3-D reference's sheet(s) by its index into the workbook's sheet list (`book`: another workbook's link number), or null. */
  sheetOf(ixti: number): { first: string; last: string; book?: number } | null;
  /** A defined name by its index (1-based). */
  nameOf(index: number): string | null;
  /** An external name (an add-in function, or an Excel 2007 function in .xls) by sheet index and name index. */
  nameXOf(ixti: number, index: number): string | null;
  /** The bytes after the tokens, which hold array constants. */
  extra?: Uint8Array;
}

const ERRORS: Record<number, string> = { 0x00: '#NULL!', 0x07: '#DIV/0!', 0x0f: '#VALUE!', 0x17: '#REF!', 0x1d: '#NAME?', 0x24: '#NUM!', 0x2a: '#N/A', 0x2b: '#GETTING_DATA' };
const BINARY: Record<number, string> = { 0x03: '+', 0x04: '-', 0x05: '*', 0x06: '/', 0x07: '^', 0x08: '&', 0x09: '<', 0x0a: '<=', 0x0b: '=', 0x0c: '>=', 0x0d: '>', 0x0e: '<>', 0x0f: ' ', 0x10: ',', 0x11: ':' };

class Undecodable extends Error {}

const quoteSheet = (s: string) => (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(s) && !/^[A-Za-z]{1,3}\d+$/.test(s) ? s : `'${s.replace(/'/g, "''")}'`);
const num = (n: number) => String(n).replace(/e\+/, 'E+').replace(/e-/, 'E-');

/** Decode a formula's tokens to `=…` text, or null when a token can't be decoded. */
export function decodeFormula(rgce: Uint8Array, ctx: PtgContext): string | null {
  try {
    return decode(rgce, ctx);
  } catch (e) {
    if (e instanceof Undecodable || e instanceof RangeError) return null;
    throw e;
  }
}

function decode(rgce: Uint8Array, ctx: PtgContext): string {
  const dv = new DataView(rgce.buffer, rgce.byteOffset, rgce.byteLength);
  const b12 = ctx.flavor === 'biff12';
  const maxRow = b12 ? 0xfffff : 0xffff;
  const maxCol = b12 ? 0x3fff : 0xff;
  const stack: string[] = [];
  let extraPos = 0;
  const pop = () => {
    const v = stack.pop();
    if (v === undefined) throw new Undecodable('stack');
    return v;
  };

  /** A cell reference's text from its row and column fields (flags in the column's top bits). */
  const cell = (row: number, colField: number, relative: boolean) => {
    const colRel = (colField & 0x4000) !== 0;
    const rowRel = (colField & 0x8000) !== 0;
    let c = colField & 0x3fff;
    let r = row;
    if (relative) {
      // Shared-formula references are offsets from the formula's cell.
      if (rowRel) r = (((b12 ? r | 0 : (r << 16) >> 16) + ctx.row) % (maxRow + 1) + maxRow + 1) % (maxRow + 1);
      if (colRel) c = (((b12 ? (c << 18) >> 18 : ((c & 0xff) << 24) >> 24) + ctx.col) % (maxCol + 1) + maxCol + 1) % (maxCol + 1);
    }
    if (!b12) c &= 0xff;
    return { text: `${colRel ? '' : '$'}${columnLabel(c)}${rowRel ? '' : '$'}${r + 1}`, r, c, colRel, rowRel };
  };
  const area = (r0: number, r1: number, c0f: number, c1f: number, relative: boolean) => {
    const a = cell(r0, c0f, relative);
    const b = cell(r1, c1f, relative);
    // Whole columns (A:B) and whole rows (1:3) are written that way.
    if (a.r === 0 && b.r === maxRow) return `${a.colRel ? '' : '$'}${columnLabel(a.c)}:${b.colRel ? '' : '$'}${columnLabel(b.c)}`;
    if (a.c === 0 && b.c === maxCol) return `${a.rowRel ? '' : '$'}${a.r + 1}:${b.rowRel ? '' : '$'}${b.r + 1}`;
    return `${a.text}:${b.text}`;
  };
  const sheet3d = (ixti: number) => {
    const s = ctx.sheetOf(ixti);
    if (!s) throw new Undecodable('external reference');
    const sheets = s.first === s.last ? s.first : `${s.first}:${s.last}`;
    if (s.book === undefined) return s.first === s.last ? `${quoteSheet(s.first)}!` : `${quoteSheet(s.first)}:${quoteSheet(s.last)}!`;
    // Another workbook: [1]Sheet!A1, quoted as a whole when the sheet name needs it ('[1]My sheet'!A1).
    return quoteSheet(sheets).startsWith("'") ? `'[${s.book}]${sheets.replace(/'/g, "''")}'!` : `[${s.book}]${sheets}!`;
  };
  const str8 = (pos: number): [string, number] => {
    const cch = rgce[pos];
    const high = rgce[pos + 1] & 1;
    let s = '';
    let p = pos + 2;
    for (let i = 0; i < cch; i++) {
      s += String.fromCharCode(high ? dv.getUint16(p, true) : rgce[p]);
      p += high ? 2 : 1;
    }
    return [s, p];
  };
  /** An array constant from the extra data: BIFF8 and BIFF12 lay out sizes and values differently. */
  const arrayConstant = (): string => {
    const x = ctx.extra;
    if (!x) throw new Undecodable('array constant');
    const xv = new DataView(x.buffer, x.byteOffset, x.byteLength);
    let p = extraPos;
    let rows: number;
    let cols: number;
    if (b12) {
      rows = xv.getUint32(p, true);
      cols = xv.getUint32(p + 4, true);
      p += 8;
    } else {
      cols = x[p] + 1;
      rows = xv.getUint16(p + 1, true) + 1;
      p += 3;
    }
    if (rows * cols > x.length) throw new Undecodable('array constant size');
    const out: string[] = [];
    for (let r = 0; r < rows; r++) {
      const row: string[] = [];
      for (let c = 0; c < cols; c++) {
        // Value types: BIFF8 1 number, 2 string, 4 boolean, 16 error (each 8 bytes but strings);
        // BIFF12 0 number, 1 string, 2 boolean (1 byte), 4 error (4 bytes).
        const t = x[p++];
        const kind = b12 ? ['num', 'str', 'bool', '', 'err'][t] : { 0x01: 'num', 0x02: 'str', 0x04: 'bool', 0x10: 'err' }[t];
        if (kind === 'num') {
          row.push(num(xv.getFloat64(p, true)));
          p += 8;
        } else if (kind === 'str') {
          const cch = xv.getUint16(p, true);
          const high = b12 ? 1 : x[p + 2] & 1;
          p += b12 ? 2 : 3;
          let s = '';
          for (let i = 0; i < cch; i++) {
            s += String.fromCharCode(high ? xv.getUint16(p, true) : x[p]);
            p += high ? 2 : 1;
          }
          row.push(`"${s.replace(/"/g, '""')}"`);
        } else if (kind === 'bool') {
          row.push(x[p] ? 'TRUE' : 'FALSE');
          p += b12 ? 1 : 8;
        } else if (kind === 'err') {
          row.push(ERRORS[x[p]] ?? '#N/A');
          p += b12 ? 4 : 8;
        } else if (!b12 && t === 0x00) {
          row.push('');
          p += 8;
        } else throw new Undecodable('array value');
      }
      out.push(row.join(','));
    }
    extraPos = p;
    return `{${out.join(';')}}`;
  };

  let i = 0;
  while (i < rgce.length) {
    const ptg = rgce[i++];
    // Operand tokens come in reference/value/array classes (0x20/0x40/0x60): the same token.
    const base = ptg >= 0x20 ? (ptg & 0x1f) | 0x20 : ptg;
    switch (base) {
      case 0x01:
        throw new Undecodable('shared or array formula reference'); // resolved by the reader
      case 0x12:
        stack.push(`+${pop()}`);
        break;
      case 0x13:
        stack.push(`-${pop()}`);
        break;
      case 0x14:
        stack.push(`${pop()}%`);
        break;
      case 0x15:
        stack.push(`(${pop()})`);
        break;
      case 0x16:
        stack.push('');
        break;
      case 0x17: {
        let s: string;
        if (b12) {
          const cch = dv.getUint16(i, true);
          s = '';
          for (let k = 0; k < cch; k++) s += String.fromCharCode(dv.getUint16(i + 2 + k * 2, true));
          i += 2 + cch * 2;
        } else [s, i] = str8(i);
        stack.push(`"${s.replace(/"/g, '""')}"`);
        break;
      }
      case 0x18:
        throw new Undecodable('extended token'); // structured references (.xlsb) and the like
      case 0x19: {
        const grbit = rgce[i];
        if (grbit & 0x04) {
          const cases = dv.getUint16(i + 1, true);
          i += 3 + (cases + 1) * 2;
        } else i += 3;
        if (grbit & 0x10) stack.push(`SUM(${pop()})`); // tAttrSum: SUM of one argument
        break;
      }
      case 0x1c:
        stack.push(ERRORS[rgce[i]] ?? '#N/A');
        i += 1;
        break;
      case 0x1d:
        stack.push(rgce[i] ? 'TRUE' : 'FALSE');
        i += 1;
        break;
      case 0x1e:
        stack.push(String(dv.getUint16(i, true)));
        i += 2;
        break;
      case 0x1f:
        stack.push(num(dv.getFloat64(i, true)));
        i += 8;
        break;
      case 0x20:
        i += b12 ? 14 : 7;
        stack.push(arrayConstant());
        break;
      case 0x21:
      case 0x22: {
        let argc: number;
        let iftab: number;
        if (base === 0x21) {
          iftab = dv.getUint16(i, true);
          i += 2;
          const n = FIXED_ARITY[iftab];
          if (n === undefined) throw new Undecodable(`function ${iftab}`);
          argc = n;
        } else {
          argc = rgce[i] & 0x7f;
          iftab = dv.getUint16(i + 1, true) & 0x7fff;
          i += 3;
        }
        const args: string[] = [];
        for (let k = 0; k < argc; k++) args.unshift(pop());
        let name: string | undefined;
        if (iftab === 255) {
          // A function stored by name (an add-in, or an Excel 2007+ function in .xls): the name comes first.
          name = args.shift()?.replace(/^_xlfn\./i, '');
          if (!name) throw new Undecodable('user function');
        } else name = FTAB[iftab] ?? (b12 ? FTAB_XLSB[iftab] : undefined);
        if (!name) throw new Undecodable(`function ${iftab}`);
        stack.push(`${name}(${args.join(',')})`);
        break;
      }
      case 0x23: {
        const n = b12 ? dv.getUint32(i, true) : dv.getUint16(i, true);
        i += 4;
        const name = ctx.nameOf(n);
        if (name === null) throw new Undecodable('name');
        stack.push(name);
        break;
      }
      case 0x24:
      case 0x2c: {
        const row = b12 ? dv.getUint32(i, true) : dv.getUint16(i, true);
        const col = dv.getUint16(i + (b12 ? 4 : 2), true);
        i += b12 ? 6 : 4;
        stack.push(cell(row, col, base === 0x2c).text);
        break;
      }
      case 0x25:
      case 0x2d: {
        const [r0, r1, c0, c1] = b12
          ? [dv.getUint32(i, true), dv.getUint32(i + 4, true), dv.getUint16(i + 8, true), dv.getUint16(i + 10, true)]
          : [dv.getUint16(i, true), dv.getUint16(i + 2, true), dv.getUint16(i + 4, true), dv.getUint16(i + 6, true)];
        i += b12 ? 12 : 8;
        stack.push(area(r0, r1, c0, c1, base === 0x2d));
        break;
      }
      case 0x26:
        // A precomputed area: its cached ranges sit in the extra data, before any later array constant's.
        if (ctx.extra && extraPos < ctx.extra.length) {
          const xv = new DataView(ctx.extra.buffer, ctx.extra.byteOffset, ctx.extra.byteLength);
          extraPos += b12 ? 4 + 16 * xv.getUint32(extraPos, true) : 2 + 8 * xv.getUint16(extraPos, true);
        }
        i += 6; // the subexpression's tokens follow and evaluate as usual
        break;
      case 0x27:
      case 0x28:
        i += 6;
        break;
      case 0x29:
        i += 2;
        break;
      case 0x2a:
        i += b12 ? 6 : 4;
        stack.push('#REF!');
        break;
      case 0x2b:
        i += b12 ? 12 : 8;
        stack.push('#REF!');
        break;
      case 0x39: {
        const ixti = dv.getUint16(i, true);
        const n = b12 ? dv.getUint32(i + 2, true) : dv.getUint16(i + 2, true);
        i += 6;
        const name = ctx.nameXOf(ixti, n);
        if (name === null) throw new Undecodable('external name');
        stack.push(name);
        break;
      }
      case 0x3a: {
        const ixti = dv.getUint16(i, true);
        const row = b12 ? dv.getUint32(i + 2, true) : dv.getUint16(i + 2, true);
        const col = dv.getUint16(i + (b12 ? 6 : 4), true);
        i += b12 ? 8 : 6;
        stack.push(sheet3d(ixti) + cell(row, col, false).text);
        break;
      }
      case 0x3b: {
        const ixti = dv.getUint16(i, true);
        const p = i + 2;
        const [r0, r1, c0, c1] = b12
          ? [dv.getUint32(p, true), dv.getUint32(p + 4, true), dv.getUint16(p + 8, true), dv.getUint16(p + 10, true)]
          : [dv.getUint16(p, true), dv.getUint16(p + 2, true), dv.getUint16(p + 4, true), dv.getUint16(p + 6, true)];
        i += b12 ? 14 : 10;
        stack.push(sheet3d(ixti) + area(r0, r1, c0, c1, false));
        break;
      }
      case 0x3c:
        i += b12 ? 8 : 6;
        stack.push('#REF!');
        break;
      case 0x3d:
        i += b12 ? 14 : 10;
        stack.push('#REF!');
        break;
      default: {
        const op = BINARY[ptg];
        if (op === undefined) throw new Undecodable(`token 0x${ptg.toString(16)}`);
        const b = pop();
        const a = pop();
        stack.push(`${a}${op}${b}`);
      }
    }
  }
  if (stack.length !== 1) throw new Undecodable('unbalanced');
  return `=${stack[0]}`;
}
