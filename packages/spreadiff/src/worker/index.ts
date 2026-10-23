import type { DiffOptions, WorkbookDiff } from '../core/types.js';
import type { TraceOptions, WorkbookTrace } from '../core/trace.js';
import type { WorkerRequest } from './worker.js';
import { unpackWorkbookDiff, unpackWorkbookTrace } from '../core/transfer.js';

/** What a worker diff resolves to. */
export interface WorkerDiffResult {
  /** The diff, unpacked and ready to render. */
  result: WorkbookDiff;
  /** The trace of `result`, when `settings.trace` asked for one and the files have formulas. */
  trace?: WorkbookTrace;
  /** Milliseconds spent inside the worker parsing both files, diffing them and tracing (0 when not asked). */
  timings: { parse: number; diff: number; trace: number };
}

/** Per-request settings for {@link DiffWorker} beyond the diff options. */
export interface WorkerSettings {
  /** Whether the first row holds column names. Default `'auto'`; see `ParseFileOptions.header`. */
  header?: boolean | 'auto';
  /**
   * Trace downstream changes in the worker too: `true` or `TraceOptions`.
   * Recalculating can take seconds on a large model; this keeps it off the
   * main thread. Default false.
   */
  trace?: boolean | TraceOptions;
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
  diffFiles(oldFile: File, newFile: File, options?: DiffOptions, settings?: WorkerSettings): Promise<WorkerDiffResult>;
  /** Parse and diff two files' bytes in the worker. The buffers are transferred, not copied. */
  diffBuffers(oldData: FileData, newData: FileData, options?: DiffOptions, settings?: WorkerSettings): Promise<WorkerDiffResult>;
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
    else {
      const result = unpackWorkbookDiff(e.data.packed);
      p.resolve({ result, trace: e.data.trace ? unpackWorkbookTrace(e.data.trace, result) : undefined, timings: e.data.timings });
    }
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

  const diffBuffers: DiffWorker['diffBuffers'] = (oldData, newData, options, settings) => {
    const id = ++seq;
    const msg: WorkerRequest = { id, old: oldData, new: newData, options, parse: { header: settings?.header }, trace: settings?.trace };
    return new Promise((resolve, reject) => {
      pending.set(id, { resolve, reject });
      worker.postMessage(msg, [oldData.data, newData.data]);
    });
  };
  // Closures, not `this`: the methods keep working when destructured.
  return {
    diffBuffers,
    async diffFiles(oldFile, newFile, options, settings) {
      const [a, b] = await Promise.all([oldFile.arrayBuffer(), newFile.arrayBuffer()]);
      return diffBuffers({ name: oldFile.name, data: a }, { name: newFile.name, data: b }, options, settings);
    },
    terminate() {
      worker.terminate();
      failAll('Diff worker terminated');
    },
  };
}
