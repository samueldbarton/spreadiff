import { describe, expect, it } from 'vitest';
import { diffWorkbooks, traceWorkbook, type Workbook } from 'spreadiff/core';
import { sampleModel } from '../../../apps/viewer/src/samples';
import { compare, reportOf, summarize, toMarkdown } from '../src';
import { regionalModel, type ModelInputs } from './fixtures';

const base: ModelInputs = { growth: 0.1, inr: 83, cny: 7.2, indiaRevenue2024: 8300, chinaRevenue2024: 720, indiaCosts2024: 4150, chinaCosts2024: 360 };
const files = async (next: Partial<ModelInputs>) => [
  { name: 'regions-v1.xlsx', data: await regionalModel(base) },
  { name: 'regions-v2.xlsx', data: await regionalModel({ ...base, ...next }) },
] as const;

describe('compare: a regional model saved by Excel', () => {
  it('names each cell by its section, row and year', async () => {
    const [a, b] = await files({ chinaRevenue2024: 800 });
    const r = compare(a, b);
    expect(r.edits.map((e) => [e.id, e.label, e.before, e.after])).toEqual([['Model!B8', 'China › Revenue · 2024', 720, 800]]);
    const usd = r.downstream.find((c) => c.id === 'Model!C11')!;
    expect(usd.label).toBe('China › Profit (USD) · 2025');
  });

  it('reports a refreshed Bloomberg value as an outside-data edit, and what it moved', async () => {
    const [a, b] = await files({ inr: 84 });
    const r = compare(a, b);
    expect(r.edits).toHaveLength(1);
    expect(r.edits[0]).toMatchObject({ id: 'Inputs!B2', kind: 'external', external: { provider: 'Bloomberg', function: 'BDP' }, before: 83, after: 84, impact: 4 }); // Model!B6, C6 and both Summary totals
    const india = r.downstream.find((c) => c.id === 'Model!C6')!;
    expect(india).toMatchObject({ method: 'recalculated', before: 55, causeCount: 1 });
    expect(india.after).toBeCloseTo(4565 / 84, 12);
    expect(india.causes![0].contribution).toBeCloseTo(4565 / 84 - 55, 12);
    expect(r.notes.join(' ')).toMatch(/outside data/);
  });

  it('attributes a result to each edit, by recalculating without it', async () => {
    const [a, b] = await files({ growth: 0.12, inr: 84 });
    const r = compare(a, b);
    const c6 = r.downstream.find((c) => c.id === 'Model!C6')!;
    const by = Object.fromEntries(c6.causes!.map((c) => [c.label, c.contribution]));
    expect(by['Growth']).toBeCloseTo(4648 / 84 - 4565 / 84, 10); // growth, at today's FX
    expect(by['USD/INR']).toBeCloseTo(4648 / 84 - 4648 / 83, 10); // FX, at today's growth
    expect(r.totals).toMatchObject({ inputs: 1, external: 1, formulas: 0, recalculated: r.totals.downstream, unexplained: 0 });
  });

  it('puts final outputs first', async () => {
    const [a, b] = await files({ growth: 0.12 });
    const r = compare(a, b);
    const outputs = r.downstream.filter((c) => c.output).map((c) => c.id);
    expect(outputs).toEqual(['Summary!C2']);
    expect(r.downstream[0].id).toBe('Summary!C2');
  });

  it('reports nothing for identical files, and the Markdown says so', async () => {
    const [a] = await files({});
    const r = compare(a, { ...a, name: 'copy.xlsx' });
    expect(r.edits.length + r.downstream.length + r.structure.length).toBe(0);
    expect(toMarkdown(r)).toBe('## regions-v1.xlsx → copy.xlsx\n\nNo changes.\n');
  });
});

describe('toMarkdown', () => {
  const m = sampleModel();
  const report = compare(m.old, m.new);
  const md = toMarkdown(report);

  it('summarises, then lists edits, outputs, intermediate results and shape changes', () => {
    expect(md.split('\n')[2]).toBe('**4 edits** (3 inputs, 1 formula) moved **21 results downstream** (5 final outputs); 1 change to rows, columns or sheets. Every cause was found by recalculating.');
    expect(md).toMatch(/^### Edits$/m);
    expect(md).toMatch(/^### Final outputs$/m);
    expect(md).toMatch(/^### Intermediate results$/m);
    expect(md).toContain('- **Revenue**: row 6 (LATAM) added; moves 10 results');
  });

  it('shows a formula edit as formulas, with the result it gave', () => {
    expect(md).toContain('| `Revenue!E4` | APAC · cogs | `=D4*Inputs!$B$3` | `=D4*0.42` | 16758 → 19656 | 6 |');
  });

  it('lists the largest causes first, and counts the rest', () => {
    const line = md.split('\n').find((l) => l.startsWith('| `Summary!B7`'))!;
    expect(line).toContain('Tax rate · value −0.02453; APAC · cogs −0.00774; EMEA · units +0.0001429; 2 more');
  });

  it('caps tables, and can leave out intermediate results', () => {
    const short = toMarkdown(report, { maxRows: 2, intermediate: false });
    expect(short).toContain('…and 2 more edits.');
    expect(short).not.toContain('Intermediate results');
  });

  it('escapes pipes and keeps a table row on one line', () => {
    const wb = (v: string): Workbook => ({ name: 'w', sheets: [{ name: 'S', columns: ['id', 'note'], rows: [['1', v]] }] });
    const text = summarize(wb('a'), wb('b | c\nd'));
    expect(text).toContain('| b \\| c d |');
  });
});

describe('reportOf', () => {
  it('builds the same report from a diff and trace you already have', () => {
    const m = sampleModel();
    const d = diffWorkbooks(m.old, m.new);
    const r = reportOf(d, traceWorkbook(d), 'a', 'b');
    expect(r.totals).toEqual(compare(m.old, m.new).totals);
    expect(JSON.parse(JSON.stringify(r)).diff).toBeUndefined(); // the diff and trace aren't serialised
    expect(r.diff).toBe(d);
  });
});

describe('names, charts and pivot tables', () => {
  it('reports a chart pointed at a new range, with no cell edits', async () => {
    const { readFileSync } = await import('node:fs');
    const { fileURLToPath } = await import('node:url');
    const { strFromU8, strToU8, unzipSync, zipSync } = await import('fflate');
    const bytes = new Uint8Array(readFileSync(fileURLToPath(new URL('../../spreadiff/test/fixtures/poi/WithTwoCharts.xlsx', import.meta.url))));
    const files = unzipSync(bytes);
    files['xl/charts/chart1.xml'] = strToU8(strFromU8(files['xl/charts/chart1.xml']).replace('Sheet1!$A$1:$A$6', 'Sheet1!$A$1:$A$7'));
    const report = compare({ name: 'a.xlsx', data: bytes }, { name: 'b.xlsx', data: zipSync(files) });
    expect(report.edits).toEqual([]);
    expect(report.objects).toHaveLength(1);
    const md = toMarkdown(report);
    expect(md).toContain('**No cell edits**, 1 change to names, charts or pivot tables.');
    expect(md).toContain('- Chart `Sheet2 › Chart 1` modified: series 1 (1st Column) values: Sheet1!$A$1:$A$6 → Sheet1!$A$1:$A$7');
    expect(md).not.toContain('No changes.');
  });
});
