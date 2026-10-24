// Scenes for scripts/screenshots.mjs: the library's components alone, no app chrome.
import { DiffGrid, diffWorkbooks, parseCsv, traceWorkbook, type CellDisplay, type DiffLayout } from 'spreadiff';
import { HTML_CSS, renderHTML, type HtmlDisplay } from 'spreadiff/html';
import { sampleModel, sampleWorkbooks } from './samples';

const stage = document.getElementById('stage')!;
let grid: DiffGrid<any> | null = null; // eslint-disable-line @typescript-eslint/no-explicit-any

const books = { model: sampleModel(), inventory: sampleWorkbooks() };

export interface Scene {
  sample: 'model' | 'inventory' | 'cities' | 'regions';
  /** Row and column headers: labels (the default) or numbers and letters. */
  headers?: 'labels' | 'numbers';
  sheet?: string;
  renderer?: 'grid' | 'html';
  layout?: DiffLayout;
  display?: CellDisplay;
  trace?: boolean;
  annotations?: boolean;
  height?: number;
  /** Stage width in CSS px (default 1120). */
  width?: number;
}

function workbook(sample: Scene['sample']) {
  if (sample === 'cities') {
    const t = (rows: (string | number)[][]) => ({ name: 'Cities', columns: ['id', 'city', 'population', 'mayor'], rows });
    return diffWorkbooks(
      { name: 'a', sheets: [t([[1, 'Lisbon', 545923, 'Carlos Moedas'], [2, 'Porto', 231800, 'Rui Moreira'], [3, 'Braga', 193333, 'Ricardo Rio'], [5, 'Faro', 64560, 'Rogério Bacalhau']])] },
      { name: 'b', sheets: [t([[1, 'Lisbon', 548703, 'Carlos Moedas'], [2, 'Porto', 232125, 'Pedro Duarte'], [3, 'Braga', 193333, 'João Rodrigues'], [7, 'Évora', 53591, 'Carlos Zorrinho']])] },
    );
  }
  if (sample === 'regions') {
    // A regional model as analysts lay one out: a year row, region headings, indented lines.
    const csv = (g: number, cn: number) => {
      const line = (label: string, v: number) => `  ${label},${v},${Math.round(v * (1 + g))}`;
      return [',FY2024,FY2025', 'India,,', line('Revenue', 8300), line('Costs', 4150), line('Profit', 4150), 'China,,', line('Revenue', cn), line('Costs', 360), line('Profit', cn - 360), `Total profit,${4150 + cn - 360},${Math.round((4150 + cn - 360) * (1 + g))}`].join('\n');
    };
    const sheet = (text: string) => ({ ...parseCsv(text, { header: false }), name: 'Model' });
    return diffWorkbooks({ name: 'a', sheets: [sheet(csv(0.1, 720))] }, { name: 'b', sheets: [sheet(csv(0.12, 800))] }, { keys: 'none', rowSimilarity: 0.3 });
  }
  const b = books[sample];
  return diffWorkbooks(b.old, b.new);
}

(window as unknown as { scene: (s: Scene) => void }).scene = (s) => {
  grid?.cleanUp();
  grid = null;
  stage.innerHTML = '';
  stage.className = s.renderer === 'html' ? 'auto' : '';
  stage.style.width = `${s.width ?? 1120}px`;
  stage.style.height = s.renderer === 'html' ? '' : `${s.height ?? 520}px`;
  const wb = workbook(s.sample);
  const sheet = Math.max(0, wb.sheets.findIndex((x) => x.name === s.sheet));
  const diff = wb.sheets[sheet].diff;
  const trace = s.trace ? traceWorkbook(wb) : undefined;
  if (s.renderer === 'html') {
    stage.innerHTML = `<style>${HTML_CSS}</style>` + renderHTML(diff, {
      layout: s.layout ?? 'split',
      display: (s.display ?? 'stacked') as HtmlDisplay,
      kindOf: trace ? (row, col) => trace.get({ sheet, row, col })?.kind : undefined,
    });
    return;
  }
  grid = new DiffGrid({
    layout: s.layout ?? 'unified',
    cellDisplay: s.display ?? 'stacked',
    collapseUnchanged: s.sample !== 'cities' && s.sample !== 'regions',
    rowHeaders: s.headers === 'numbers' ? 'numbers' : 'labels',
    columnHeaders: s.headers === 'numbers' ? 'letters' : 'labels',
    trace: trace ? { trace, workbook: wb, sheet } : undefined,
    annotations: s.annotations ? [{ row: 1, column: 'mayor', data: { author: 'ana', text: 'Confirmed with the 2025 local election results.' } }] : undefined,
    renderAnnotation: s.annotations
      ? (a) => {
          const el = document.createElement('div');
          el.className = 'note';
          const d = a.data as { author: string; text: string };
          el.innerHTML = `<b>@${d.author}</b> ${d.text}`;
          return el;
        }
      : undefined,
  });
  grid.render({ diff, container: stage });
  (window as unknown as { grid: unknown }).grid = grid;
};
(window as unknown as { ready: boolean }).ready = true;
