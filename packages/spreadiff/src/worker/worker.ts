/// <reference lib="webworker" />
import { diffWorkbooks } from '../core/diff.js';
import { parseFile } from '../core/file.js';
import { packWorkbookDiff } from '../core/transfer.js';
import type { DiffOptions } from '../core/types.js';

/**
 * A request to the diff worker (`spreadiff/worker/worker.js`). Normally sent
 * for you by {@link createDiffWorker}. Replies are `{ id, packed, timings }`
 * or `{ id, error }`.
 */
export interface WorkerRequest {
  /** Echoed back in the reply, to match it to the request. */
  id: number;
  /** The old file. */
  old: { name: string; data: ArrayBuffer };
  /** The new file. */
  new: { name: string; data: ArrayBuffer };
  /** Diff options. */
  options?: DiffOptions;
  /** Parse options: whether the first row is a header (default true). */
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
