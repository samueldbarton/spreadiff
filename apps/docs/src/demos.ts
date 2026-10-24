// Live demos, mounted on <div data-demo="name">. Each has a Preview and a
// Code tab; the code shown is what the preview runs.
import { marked } from 'marked';
import { DiffGrid, diffWorkbooks, traceWorkbook, computeMissingValues, type CellDisplay, type CellValue, type DiffLayout, type Table, type Workbook } from 'spreadiff';
import { HTML_CSS, renderHTML } from 'spreadiff/html';
import { compare, toMarkdown } from 'spreadiff-lite';
import { sampleModel, sampleWorkbooks } from '../../viewer/src/samples';

const h = <K extends keyof HTMLElementTagNameMap>(tag: K, props: Record<string, string> = {}, ...children: (Node | string)[]) => {
  const el = document.createElement(tag);
  for (const [k, v] of Object.entries(props)) el.setAttribute(k, v);
  el.append(...children);
  return el;
};

/** A demo frame: optional controls, Preview / Code tabs. */
function frame(el: HTMLElement, code: string, controls?: HTMLElement) {
  el.innerHTML = '';
  const id = Math.random().toString(36).slice(2, 8);
  const tabs = h('div', { class: 'demo-tabs', role: 'tablist', 'aria-label': 'Demo' });
  const preview = h('div', { class: 'demo-preview', role: 'tabpanel', id: `p-${id}` });
  const pre = h('pre', { class: 'shiki demo-code' }, h('code', {}, code));
  const codePanel = h('div', { class: 'demo-panel', role: 'tabpanel', id: `c-${id}`, hidden: '' }, pre);
  const tab = (label: string, panel: HTMLElement, selected: boolean) => {
    const b = h('button', { type: 'button', role: 'tab', 'aria-selected': String(selected), 'aria-controls': panel.id }, label);
    b.addEventListener('click', () => {
      tabs.querySelectorAll('[role=tab]').forEach((t) => t.setAttribute('aria-selected', String(t === b)));
      preview.hidden = panel !== preview;
      codePanel.hidden = panel !== codePanel;
    });
    return b;
  };
  tabs.append(tab('Preview', preview, true), tab('Code', codePanel, false));
  const bar = h('div', { class: 'demo-bar' }, tabs);
  if (controls) bar.append(controls);
  el.append(bar, preview, codePanel);
  return preview;
}

/** A segmented control. */
function segmented<T extends string>(label: string, options: T[], value: T, onChange: (v: T) => void) {
  const group = h('div', { class: 'segmented', role: 'radiogroup', 'aria-label': label });
  for (const o of options) {
    const b = h('button', { type: 'button', role: 'radio', 'aria-checked': String(o === value) }, o);
    b.addEventListener('click', () => {
      group.querySelectorAll('button').forEach((x) => x.setAttribute('aria-checked', String(x === b)));
      onChange(o);
    });
    group.append(b);
  }
  return group;
}

function gridDemo(el: HTMLElement) {
  const { old: a, new: b } = sampleWorkbooks();
  const diff = diffWorkbooks(a, b).sheets[0].diff;
  let layout: DiffLayout = 'split';
  let cellDisplay: CellDisplay = 'stacked';
  const code = () => `import { DiffGrid, diffTables } from 'spreadiff';

const diff = diffTables(before, after); // matched by the id column
new DiffGrid({ layout: '${layout}', cellDisplay: '${cellDisplay}' }).render({ diff, container });`;
  const controls = h('div', { class: 'demo-controls' });
  const grid = new DiffGrid({ layout, cellDisplay });
  const update = () => {
    grid.setOptions({ layout, cellDisplay });
    el.querySelector('.demo-code code')!.textContent = code();
  };
  controls.append(
    segmented('Layout', ['unified', 'split'] as DiffLayout[], layout, (v) => ((layout = v), update())),
    segmented('Cells', ['stacked', 'new', 'delta', 'inline'] as CellDisplay[], cellDisplay, (v) => ((cellDisplay = v), update())),
  );
  const preview = frame(el, code(), controls);
  const box = h('div', { class: 'demo-grid' });
  preview.append(box);
  grid.render({ diff, container: box });
}

function traceDemo(el: HTMLElement, hero = false) {
  const { old: a, new: b } = sampleModel();
  const d = diffWorkbooks(a, b);
  const trace = traceWorkbook(d);
  let sheet = d.sheets.findIndex((s) => s.name === 'Revenue');
  const code = `import { DiffGrid, diffWorkbooks, parseFile, traceWorkbook } from 'spreadiff';

const d = diffWorkbooks(parseFile('model-v1.xlsx', a), parseFile('model-v2.xlsx', b));
const trace = traceWorkbook(d); // undo each edit, recalculate: causes + contributions
const sheet = ${sheet};
new DiffGrid({ cellDisplay: 'delta', trace: { trace, workbook: d, sheet } })
  .render({ diff: d.sheets[sheet].diff, container });`;
  const grid = new DiffGrid({ cellDisplay: 'delta', trace: { trace, workbook: d, sheet } });
  const controls = segmented(
    'Sheet',
    d.sheets.map((s) => s.name),
    d.sheets[sheet].name,
    (name) => {
      sheet = d.sheets.findIndex((s) => s.name === name);
      grid.setOptions({ trace: { trace, workbook: d, sheet } });
      grid.render({ diff: d.sheets[sheet].diff, container: box });
    },
  );
  const preview = frame(el, code, controls);
  const box = h('div', { class: hero ? 'demo-grid hero-grid' : 'demo-grid short' });
  preview.append(box);
  grid.render({ diff: d.sheets[sheet].diff, container: box });
  const s = trace.stats;
  preview.append(h('p', { class: 'demo-note' }, `${s.inputs} input edits · ${s.formulas} formula edit · ${s.computed} downstream results, every cause recalculated. Hover an amber cell.`));
}

function htmlDemo(el: HTMLElement) {
  const { old: a, new: b } = sampleWorkbooks();
  const diff = diffWorkbooks(a, b).sheets[0].diff;
  let layout: 'split' | 'unified' = 'split';
  const code = () => `import { HTML_CSS, renderHTML } from 'spreadiff/html';

const html = \`<style>\${HTML_CSS}</style>\` + renderHTML(diff, { layout: '${layout}', maxRows: 12 });`;
  const out = h('div', { class: 'demo-html' });
  const shadow = out.attachShadow({ mode: 'open' });
  const draw = () => {
    shadow.innerHTML = `<style>${HTML_CSS}</style>` + renderHTML(diff, { layout, maxRows: 12 });
    el.querySelector('.demo-code code')!.textContent = code();
  };
  const preview = frame(el, code(), segmented('Layout', ['split', 'unified'] as const as ('split' | 'unified')[], layout, (v) => ((layout = v), draw())));
  preview.append(out);
  draw();
}

/** The regional model from the spreadiff-lite guide, built in memory. */
function regionModel(name: string, growth: number, inr: number): Workbook {
  const sheet = (n: string, rows: (CellValue | string)[][]): Table => ({
    name: n,
    columns: rows[0].map((_, i) => String.fromCharCode(65 + i)),
    origin: { row: 0, col: 0 },
    rows: rows.map((r) => r.map((v) => (typeof v === 'string' && v.startsWith('=') ? null : v))),
    formulas: rows.map((r) => r.map((v) => (typeof v === 'string' && v.startsWith('=') ? v : null))),
  });
  const wb: Workbook = {
    name,
    sheets: [
      sheet('Inputs', [['Growth', growth], ['USD/INR', '=BDP("USDINR Curncy","PX_LAST")']]),
      sheet('Model', [
        [null, 2024, '=B1+1'],
        ['India', null, null],
        ['  Revenue', 8300, '=B3*(1+Inputs!$B$1)'],
        ['  Costs', 4150, '=B4*(1+Inputs!$B$1)'],
        ['  Profit (USD)', '=(B3-B4)/Inputs!$B$2', '=(C3-C4)/Inputs!$B$2'],
      ]),
    ],
  };
  wb.sheets[0].rows[1][1] = inr;
  computeMissingValues(wb);
  return wb;
}

function liteDemo(el: HTMLElement) {
  let growth = 0.12;
  let inr = 84;
  let view: 'rendered' | 'markdown' = 'rendered';
  const code = () => `import { compare, toMarkdown } from 'spreadiff-lite';

// Before: growth 10%, USD/INR 83 (a Bloomberg BDP). After: what you set above.
const report = compare(before, after);
console.log(toMarkdown(report));`;
  const num = (label: string, value: number, step: string, onInput: (v: number) => void) => {
    const input = h('input', { type: 'number', value: String(value), step, 'aria-label': label });
    input.addEventListener('input', () => {
      const v = Number(input.value);
      if (Number.isFinite(v)) onInput(v);
    });
    return h('label', { class: 'num' }, `${label} `, input);
  };
  const controls = h(
    'div',
    { class: 'demo-controls' },
    num('Growth', growth, '0.01', (v) => ((growth = v), draw())),
    num('USD/INR', inr, '0.5', (v) => ((inr = v), draw())),
    segmented('View', ['rendered', 'markdown'] as ('rendered' | 'markdown')[], view, (v) => ((view = v), draw())),
  );
  const preview = frame(el, code(), controls);
  const out = h('div', { class: 'demo-md' });
  preview.append(out);
  const before = regionModel('regions-v1.xlsx', 0.1, 83);
  function draw() {
    const md = toMarkdown(compare(before, regionModel('regions-v2.xlsx', growth, inr)));
    if (view === 'markdown') {
      out.className = 'demo-md raw';
      out.replaceChildren(h('pre', { class: 'shiki' }, h('code', {}, md)));
    } else {
      out.className = 'demo-md prose';
      out.innerHTML = marked.parse(md.replace(/^## /m, '#### ').replace(/^### /gm, '##### '), { async: false }) as string;
      out.querySelectorAll('table').forEach((t) => t.replaceWith(h('div', { class: 'table' }, t.cloneNode(true))));
    }
  }
  draw();
}

function themeDemo(el: HTMLElement) {
  const { old: a, new: b } = sampleWorkbooks();
  const diff = diffWorkbooks(a, b).sheets[0].diff;
  const vars: Record<string, string> = { '--sd-added': '#1a7f37', '--sd-removed': '#cf222e', '--sd-modified': '#0969da', '--sd-row-height': '30px' };
  const code = () => `.review-panel {\n${Object.entries(vars).map(([k, v]) => `  ${k}: ${v};`).join('\n')}\n}`;
  const box = h('div', { class: 'demo-grid short review-panel' });
  const grid = new DiffGrid({ layout: 'unified' });
  const apply = () => {
    for (const [k, v] of Object.entries(vars)) box.style.setProperty(k, v);
    grid.refreshTheme();
    el.querySelector('.demo-code code')!.textContent = code();
  };
  const controls = h('div', { class: 'demo-controls' });
  for (const k of ['--sd-added', '--sd-removed', '--sd-modified']) {
    const input = h('input', { type: 'color', value: vars[k], 'aria-label': k });
    input.addEventListener('input', () => ((vars[k] = input.value), apply()));
    controls.append(h('label', { class: 'swatch' }, input, k.replace('--sd-', '')));
  }
  controls.append(
    segmented('Row height', ['26px', '30px', '36px'], vars['--sd-row-height'], (v) => ((vars['--sd-row-height'] = v), apply())),
  );
  const preview = frame(el, code(), controls);
  el.querySelector('.demo-code')!.classList.add('css');
  preview.append(box);
  grid.render({ diff, container: box });
  apply();
}

const DEMOS: Record<string, (el: HTMLElement) => void> = {
  grid: gridDemo,
  trace: (el) => traceDemo(el),
  hero: (el) => traceDemo(el, true),
  html: htmlDemo,
  lite: liteDemo,
  theme: themeDemo,
};

/** Mount the demo named by `el.dataset.demo`. */
export function mount(el: HTMLElement) {
  const demo = DEMOS[el.dataset.demo ?? ''];
  if (!demo) return;
  try {
    demo(el);
  } catch (e) {
    el.innerHTML = '';
    el.append(h('p', { class: 'demo-loading' }, `The demo failed to load: ${(e as Error).message}`));
  }
}
