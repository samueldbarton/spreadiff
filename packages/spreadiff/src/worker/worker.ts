/// <reference lib="webworker" />
import { diffWorkbooks } from '../core/diff.js';
import { parseFile } from '../core/file.js';
import { packWorkbookDiff } from '../core/transfer.js';
import type { DiffOptions } from '../core/types.js';

export interface WorkerRequest {
  id: number;
  old: { name: string; data: ArrayBuffer };
  new: { name: string; data: ArrayBuffer };
  options?: DiffOptions;
  parse?: { header?: boolean };
}

self.onmessage = (e: MessageEvent<WorkerRequest>) => {
  const { id, options, parse } = e.data;
  try {
    const t0 = performance.now();
    const a = parseFile(e.data.old.name, e.data.old.data, parse);
    const b = parseFile(e.data.new.name, e.data.new.data, parse);
    const t1 = performance.now();
    const result = diffWorkbooks(a, b, options);
    const t2 = performance.now();
    const { packed, transfer } = packWorkbookDiff(result);
    (self as unknown as Worker).postMessage({ id, packed, timings: { parse: t1 - t0, diff: t2 - t1 } }, transfer);
  } catch (err) {
    (self as unknown as Worker).postMessage({ id, error: err instanceof Error ? err.message : String(err) });
  }
};
