import { parseCsv } from './csv.js';
import { parseXlsx } from './xlsx.js';
import { parseOds } from './ods.js';
import type { Workbook } from './types.js';

/** Options for {@link parseFile}. */
export interface ParseFileOptions {
  /**
   * Whether the first (non-empty) row holds column names. Default `'auto'`:
   * for XLSX, see `ParseXlsxOptions.header`; CSV treats `'auto'` as true.
   */
  header?: boolean | 'auto';
  /** CSV delimiter; detected when omitted (`.tsv` files default to tab). */
  delimiter?: string;
}

/** Parse a CSV/TSV/XLSX file by name and contents into a workbook. */
export function parseFile(name: string, data: ArrayBuffer | Uint8Array, options: ParseFileOptions = {}): Workbook {
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
  const isZip = bytes[0] === 0x50 && bytes[1] === 0x4b && bytes[2] === 0x03 && bytes[3] === 0x04;
  if (isZip) {
    // ODF stores its mimetype first and uncompressed, so it shows in the first bytes.
    const head = String.fromCharCode(...bytes.subarray(30, 110));
    if (/\.ods$/i.test(name) || head.includes('opendocument.spreadsheet')) return parseOds(bytes, { name, header: options.header });
    return parseXlsx(bytes, { name, header: options.header });
  }
  // Legacy binary formats would otherwise be "parsed" as garbage CSV.
  if (bytes[0] === 0xd0 && bytes[1] === 0xcf && bytes[2] === 0x11 && bytes[3] === 0xe0) {
    throw new Error(`${name}: legacy .xls (BIFF) files aren't supported. Save as .xlsx or .csv.`);
  }
  if (/\.(xlsx|xlsm|xlsb|ods)$/i.test(name)) throw new Error(`${name}: not a valid ${name.slice(name.lastIndexOf('.'))} file`);
  const text = decodeText(bytes);
  const delimiter = options.delimiter ?? (/\.tsv$/i.test(name) ? '\t' : undefined);
  const sheet = parseCsv(text, { name: name.replace(/\.[^.]+$/, ''), delimiter, header: options.header !== false });
  return { name, sheets: [sheet] };
}

/** UTF-8 by default; UTF-16 when a BOM says so (Excel's "Unicode Text" export). */
function decodeText(bytes: Uint8Array): string {
  if (bytes[0] === 0xff && bytes[1] === 0xfe) return new TextDecoder('utf-16le').decode(bytes);
  if (bytes[0] === 0xfe && bytes[1] === 0xff) return new TextDecoder('utf-16be').decode(bytes);
  return new TextDecoder('utf-8').decode(bytes);
}
