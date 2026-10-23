import { describe, expect, it } from 'vitest';
import { createDiffWorker } from '../src/worker';
import { diffWorkbooks, packWorkbookDiff, parseFile } from '../src/core';

/** An in-process stand-in for a module Worker running src/worker/worker.ts. */
class FakeWorker {
  onmessage: ((e: MessageEvent) => void) | null = null;
  onerror: ((e: ErrorEvent) => void) | null = null;
  onmessageerror: (() => void) | null = null;
  crash = false;
  postMessage(msg: { id: number; old: { name: string; data: ArrayBuffer }; new: { name: string; data: ArrayBuffer } }) {
    queueMicrotask(() => {
      if (this.crash) return this.onerror?.({ message: 'boom', preventDefault() {} } as ErrorEvent);
      const result = diffWorkbooks(parseFile(msg.old.name, msg.old.data), parseFile(msg.new.name, msg.new.data));
      const { packed } = packWorkbookDiff(result);
      this.onmessage?.({ data: { id: msg.id, packed, timings: { parse: 1, diff: 1 } } } as MessageEvent);
    });
  }
  terminate() {}
}

const buf = (s: string) => new TextEncoder().encode(s).buffer as ArrayBuffer;

describe('createDiffWorker', () => {
  it('works with destructured methods', async () => {
    const { diffBuffers } = createDiffWorker(new FakeWorker() as unknown as Worker);
    const { result } = await diffBuffers({ name: 'a.csv', data: buf('id,v\n1,a') }, { name: 'b.csv', data: buf('id,v\n1,b') });
    expect(result.sheets[0].diff.stats.cellsChanged).toBe(1);
  });
  it('rejects pending requests when the worker crashes', async () => {
    const w = new FakeWorker();
    w.crash = true;
    const differ = createDiffWorker(w as unknown as Worker);
    await expect(differ.diffBuffers({ name: 'a.csv', data: buf('a\n1') }, { name: 'b.csv', data: buf('a\n2') })).rejects.toThrow('boom');
  });
});
