import type { DiffOptions, WorkbookDiff } from '../core/types.js';
import type { WorkerRequest } from './worker.js';
import { unpackWorkbookDiff } from '../core/transfer.js';

export interface WorkerDiffResult {
  result: WorkbookDiff;
  timings: { parse: number; diff: number };
}

/**
 * Parse and diff two files off the main thread. Pass a Worker created from
 * `spreadiff/worker/worker.js`, e.g. with Vite:
 *
 * ```ts
 * const worker = new Worker(new URL('spreadiff/worker/worker.js', import.meta.url), { type: 'module' });
 * const { result } = await diffFilesInWorker(worker, oldFile, newFile);
 * ```
 */
export function createDiffWorker(worker: Worker) {
  let seq = 0;
  const pending = new Map<number, { resolve: (r: WorkerDiffResult) => void; reject: (e: Error) => void }>();
  worker.onmessage = (e) => {
    const p = pending.get(e.data.id);
    if (!p) return;
    pending.delete(e.data.id);
    if (e.data.error) p.reject(new Error(e.data.error));
    else p.resolve({ result: unpackWorkbookDiff(e.data.packed), timings: e.data.timings });
  };
  return {
    async diffFiles(oldFile: File, newFile: File, options?: DiffOptions, parse?: { header?: boolean }) {
      const [a, b] = await Promise.all([oldFile.arrayBuffer(), newFile.arrayBuffer()]);
      return this.diffBuffers({ name: oldFile.name, data: a }, { name: newFile.name, data: b }, options, parse);
    },
    diffBuffers(
      oldData: { name: string; data: ArrayBuffer },
      newData: { name: string; data: ArrayBuffer },
      options?: DiffOptions,
      parse?: { header?: boolean },
    ): Promise<WorkerDiffResult> {
      const id = ++seq;
      const msg: WorkerRequest = { id, old: oldData, new: newData, options, parse };
      return new Promise((resolve, reject) => {
        pending.set(id, { resolve, reject });
        worker.postMessage(msg, [oldData.data, newData.data]);
      });
    },
    terminate: () => worker.terminate(),
  };
}
