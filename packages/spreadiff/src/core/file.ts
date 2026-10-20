import { parseCsv } from './csv.js';
import { parseXlsx } from './xlsx.js';
import type { Workbook } from './types.js';

export interface ParseFileOptions {
  header?: boolean;
  delimiter?: string;
}

/** Parse a CSV/TSV/XLSX file by name and contents into a workbook. */
export function parseFile(name: string, data: ArrayBuffer | Uint8Array, options: ParseFileOptions = {}): Workbook {
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
  const isZip = bytes[0] === 0x50 && bytes[1] === 0x4b;
  if (isZip || /\.(xlsx|xlsm)$/i.test(name)) return parseXlsx(bytes, { name, header: options.header });
  const text = new TextDecoder().decode(bytes);
  const delimiter = options.delimiter ?? (/\.tsv$/i.test(name) ? '\t' : undefined);
  const sheet = parseCsv(text, { name: name.replace(/\.[^.]+$/, ''), delimiter, header: options.header });
  return { name, sheets: [sheet] };
}
