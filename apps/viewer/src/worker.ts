/// <reference lib="webworker" />
import { diffWorkbooks, packWorkbookDiff, parseCsv, parseFile, type DiffOptions, type Workbook } from 'spreadiff/core';
import { syntheticCsv } from './samples';

export type Req =
  | { id: number; type: 'files'; old: { name: string; data: ArrayBuffer }; new: { name: string; data: ArrayBuffer }; options: DiffOptions; header: boolean }
  | { id: number; type: 'books'; old: Workbook; new: Workbook; options: DiffOptions }
  | { id: number; type: 'synthetic'; rows: number; cols: number; options: DiffOptions };

self.onmessage = (e: MessageEvent<Req>) => {
  const req = e.data;
  try {
    const t0 = performance.now();
    let a: Workbook;
    let b: Workbook;
    if (req.type === 'files') {
      a = parseFile(req.old.name, req.old.data, { header: req.header });
      b = parseFile(req.new.name, req.new.data, { header: req.header });
    } else if (req.type === 'books') {
      a = req.old;
      b = req.new;
    } else {
      const s = syntheticCsv(req.rows, req.cols);
      a = { name: `synthetic-${req.rows}-a.csv`, sheets: [parseCsv(s.old, { name: 'synthetic' })] };
      b = { name: `synthetic-${req.rows}-b.csv`, sheets: [parseCsv(s.new, { name: 'synthetic' })] };
    }
    const t1 = performance.now();
    const result = diffWorkbooks(a, b, req.options);
    const t2 = performance.now();
    const { packed, transfer } = packWorkbookDiff(result);
    const t3 = performance.now();
    postMessage({ id: req.id, packed, names: [a.name, b.name], timings: { parse: t1 - t0, diff: t2 - t1, pack: t3 - t2 } }, transfer);
  } catch (err) {
    postMessage({ id: req.id, error: err instanceof Error ? err.message : String(err) });
  }
};
