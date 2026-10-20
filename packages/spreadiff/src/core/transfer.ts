import type { CellValue, Table, TableDiff, WorkbookDiff } from './types.js';

/**
 * Compact, transferable encoding of tables for crossing a worker boundary.
 * Structured-cloning `CellValue[][]` costs one allocation per cell (seconds
 * for millions of cells); this sends one type array, one number array and one
 * joined string, which move in milliseconds.
 */
export interface PackedTable {
  name: string;
  columns: string[];
  width: number;
  height: number;
  /** 0 null, 1 string, 2 number, 3 true, 4 false, row-major. */
  types: Uint8Array;
  numbers: Float64Array;
  /** String cells joined by U+0000, in row-major order. */
  strings: string;
  formulas?: (string | null)[][];
}

const SEP = '\u0000';

export function packTable(t: Table): PackedTable {
  const width = t.columns.length;
  const height = t.rows.length;
  const types = new Uint8Array(width * height);
  const numbers = new Float64Array(width * height);
  const parts: string[] = [];
  for (let r = 0; r < height; r++) {
    const row = t.rows[r];
    for (let c = 0; c < width; c++) {
      const v = row[c];
      const i = r * width + c;
      if (v === null || v === undefined) continue;
      if (typeof v === 'string') {
        types[i] = 1;
        parts.push(v.includes(SEP) ? v.replaceAll(SEP, '�') : v);
      } else if (typeof v === 'number') {
        types[i] = 2;
        numbers[i] = v;
      } else types[i] = v ? 3 : 4;
    }
  }
  return { name: t.name, columns: t.columns, width, height, types, numbers, strings: parts.join(SEP), formulas: t.formulas };
}

export function unpackTable(p: PackedTable): Table {
  const strs = p.strings.length || p.types.includes(1) ? p.strings.split(SEP) : [];
  let s = 0;
  const rows: CellValue[][] = new Array(p.height);
  const { types, numbers, width } = p;
  for (let r = 0; r < p.height; r++) {
    const row: CellValue[] = new Array(width);
    const base = r * width;
    for (let c = 0; c < width; c++) {
      const k = types[base + c];
      row[c] = k === 0 ? null : k === 1 ? strs[s++] : k === 2 ? numbers[base + c] : k === 3;
    }
    rows[r] = row;
  }
  const t: Table = { name: p.name, columns: p.columns, rows };
  if (p.formulas) t.formulas = p.formulas;
  return t;
}

export type PackedTableDiff = Omit<TableDiff, 'oldTable' | 'newTable'> & { oldTable: PackedTable; newTable: PackedTable };
export interface PackedWorkbookDiff {
  sheets: { name: string; status: WorkbookDiff['sheets'][number]['status']; diff: PackedTableDiff }[];
}

/** Pack a workbook diff; returns the message and the buffers to transfer. */
export function packWorkbookDiff(d: WorkbookDiff): { packed: PackedWorkbookDiff; transfer: ArrayBuffer[] } {
  const transfer: ArrayBuffer[] = [];
  const packed: PackedWorkbookDiff = {
    sheets: d.sheets.map((s) => {
      const oldTable = packTable(s.diff.oldTable);
      const newTable = packTable(s.diff.newTable);
      for (const p of [oldTable, newTable]) transfer.push(p.types.buffer as ArrayBuffer, p.numbers.buffer as ArrayBuffer);
      transfer.push(s.diff.oldRows.buffer as ArrayBuffer, s.diff.newRows.buffer as ArrayBuffer, s.diff.status.buffer as ArrayBuffer);
      return { name: s.name, status: s.status, diff: { ...s.diff, oldTable, newTable } };
    }),
  };
  return { packed, transfer };
}

export function unpackWorkbookDiff(p: PackedWorkbookDiff): WorkbookDiff {
  return {
    sheets: p.sheets.map((s) => ({
      name: s.name,
      status: s.status,
      diff: { ...s.diff, oldTable: unpackTable(s.diff.oldTable), newTable: unpackTable(s.diff.newTable) },
    })),
  };
}
