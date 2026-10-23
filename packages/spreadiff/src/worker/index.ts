import type { DiffOptions, WorkbookDiff } from '../core/types.js';
import type { WorkerRequest } from './worker.js';
import { unpackWorkbookDiff } from '../core/transfer.js';

/** What a worker diff resolves to. */
export interface WorkerDiffResult {
  /** The diff, unpacked and ready to render. */
  result: WorkbookDiff;
  /** Milliseconds spent inside the worker parsing both files and diffing them. */
  timings: { parse: number; diff: number };
}

/** A file's name (used to pick the parser) and bytes. */
export interface FileData {
  /** File name; `.xlsx`/`.xlsm` or zip content → XLSX, `.tsv` → tab-delimited, else CSV. */
  name: string;
  /** File contents. Transferred to the worker: unusable afterwards on this side. */
  data: ArrayBuffer;
}

/** A client for the diff worker, from {@link createDiffWorker}. Methods can be destructured. */
export interface DiffWorker {
  /** Read two `File`s (from an `<input>` or drop), then parse and diff them in the worker. */
  diffFiles(oldFile: File, newFile: File, options?: DiffOptions, parse?: { header?: boolean }): Promise<WorkerDiffResult>;
  /** Parse and diff two files' bytes in the worker. The buffers are transferred, not copied. */
  diffBuffers(oldData: FileData, newData: FileData, options?: DiffOptions, parse?: { header?: boolean }): Promise<WorkerDiffResult>;
  /** Stop the worker. Pending requests reject. */
  terminate(): void;
}

/**
 * Parse and diff files off the main thread, so a big file never freezes the
 * page. Pass a module `Worker` created from `spreadiff/worker/worker.js`;
 * with Vite or webpack 5:
 *
 * ```ts
 * import { createDiffWorker } from 'spreadiff/worker';
 * declare const oldFile: File, newFile: File;
 *
 * const differ = createDiffWorker(
 *   new Worker(new URL('spreadiff/worker/worker.js', import.meta.url), { type: 'module' }),
 * );
 * const { result, timings } = await differ.diffFiles(oldFile, newFile, { keys: 'auto' });
 * ```
 *
 * Results cross back in a compact encoding (see {@link packWorkbookDiff}).
 * If the worker crashes (for example, out of memory on a huge file), every
 * pending request rejects instead of hanging.
 */
export function createDiffWorker(worker: Worker): DiffWorker {
  let seq = 0;
  const pending = new Map<number, { resolve: (r: WorkerDiffResult) => void; reject: (e: Error) => void }>();
  worker.onmessage = (e) => {
    const p = pending.get(e.data.id);
    if (!p) return;
    pending.delete(e.data.id);
    if (e.data.error) p.reject(new Error(e.data.error));
    else p.resolve({ result: unpackWorkbookDiff(e.data.packed), timings: e.data.timings });
  };
  // A crashed worker (e.g. out of memory on a huge file) never replies;
  // fail every pending request instead of leaving them hanging forever.
  const failAll = (message: string) => {
    for (const p of pending.values()) p.reject(new Error(message));
    pending.clear();
  };
  worker.onerror = (e) => {
    e.preventDefault?.();
    failAll(e.message || 'Diff worker crashed (possibly out of memory)');
  };
  worker.onmessageerror = () => failAll('Diff worker sent an unreadable message');

  const diffBuffers: DiffWorker['diffBuffers'] = (oldData, newData, options, parse) => {
    const id = ++seq;
    const msg: WorkerRequest = { id, old: oldData, new: newData, options, parse };
    return new Promise((resolve, reject) => {
      pending.set(id, { resolve, reject });
      worker.postMessage(msg, [oldData.data, newData.data]);
    });
  };
  // Closures, not `this`: the methods keep working when destructured.
  return {
    diffBuffers,
    async diffFiles(oldFile, newFile, options, parse) {
      const [a, b] = await Promise.all([oldFile.arrayBuffer(), newFile.arrayBuffer()]);
      return diffBuffers({ name: oldFile.name, data: a }, { name: newFile.name, data: b }, options, parse);
    },
    terminate() {
      worker.terminate();
      failAll('Diff worker terminated');
    },
  };
}
