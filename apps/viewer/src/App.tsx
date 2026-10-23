import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { diffText, formatCell, numericDelta, toUnifiedCsv, traceWorkbook, unpackWorkbookDiff, type CellRef, type DiffOptions, type WorkbookDiff } from 'spreadiff';
import { DiffGrid, type CellDisplay, type CellInfo, type DiffGridHandle, type DiffLayout } from 'spreadiff/react';
import { CodeView } from './CodeView';
import { GlideDiff } from './GlideDiff';
import { sampleModel, sampleWorkbooks } from './samples';
import { TracePanel, type TraceView } from './TracePanel';
import type { Req } from './worker';
import DiffWorker from './worker?worker';

type Theme = 'system' | 'light' | 'dark';
type Renderer = 'spreadiff' | 'code' | 'glide';

interface Loaded {
  result: WorkbookDiff;
  names: [string, string];
  timings: { parse: number; diff: number; transfer: number };
}

// One long-lived worker; requests are matched by id. If it dies (e.g. out
// of memory on a huge file), fail what's pending and start a fresh one.
let worker = new DiffWorker();
let seq = 0;
const inflight = new Map<number, (err: Error) => void>();
function watchWorker(w: Worker) {
  w.addEventListener('error', (e) => {
    e.preventDefault();
    const err = new Error(e.message || 'The diff worker crashed (the files may be too large for this browser tab).');
    for (const fail of inflight.values()) fail(err);
    inflight.clear();
    w.terminate();
    worker = new DiffWorker();
    watchWorker(worker);
  });
}
watchWorker(worker);
type DistributiveOmit<T, K extends keyof T> = T extends unknown ? Omit<T, K> : never;
function run(req: DistributiveOmit<Req, 'id'>, transfer: Transferable[] = []): Promise<Loaded> {
  const id = ++seq;
  const t0 = performance.now();
  const w = worker;
  return new Promise((resolve, reject) => {
    inflight.set(id, (err) => {
      w.removeEventListener('message', onMessage);
      reject(err);
    });
    const onMessage = (e: MessageEvent) => {
      if (e.data.id !== id) return;
      w.removeEventListener('message', onMessage);
      inflight.delete(id);
      if (e.data.error) return reject(new Error(e.data.error));
      const result = unpackWorkbookDiff(e.data.packed);
      const total = performance.now() - t0;
      const { parse, diff } = e.data.timings;
      // "transfer" = pack in worker + postMessage + unpack on main thread.
      resolve({ result, names: e.data.names, timings: { parse, diff, transfer: Math.max(0, total - parse - diff) } });
    };
    w.addEventListener('message', onMessage);
    w.postMessage({ ...req, id }, transfer);
  });
}

const fmt = (n: number) => n.toLocaleString();
const ms = (n: number) => (n < 10 ? `${n.toFixed(1)} ms` : n < 1000 ? `${Math.round(n)} ms` : `${(n / 1000).toFixed(2)} s`);

function readTheme(): Theme {
  try {
    const t = localStorage.getItem('sd-theme');
    return t === 'light' || t === 'dark' ? t : 'system';
  } catch {
    return 'system'; // storage blocked (private mode, sandboxed iframe)
  }
}

/** What produced the current diff, so option changes can re-run it. */
type Source = { kind: 'files'; a: File; b: File } | { kind: 'sample' } | { kind: 'model' } | { kind: 'synthetic'; rows: number };

function usePrefersDark(theme: Theme) {
  const [sys, setSys] = useState(() => matchMedia('(prefers-color-scheme: dark)').matches);
  useEffect(() => {
    const mq = matchMedia('(prefers-color-scheme: dark)');
    const fn = () => setSys(mq.matches);
    mq.addEventListener('change', fn);
    return () => mq.removeEventListener('change', fn);
  }, []);
  return theme === 'system' ? sys : theme === 'dark';
}

export function App() {
  const [files, setFiles] = useState<[File | null, File | null]>([null, null]);
  const [loaded, setLoaded] = useState<Loaded | null>(null);
  const [busy, setBusy] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [sheet, setSheet] = useState(0);
  const sheetName = useRef<string | null>(null);
  const [layout, setLayout] = useState<DiffLayout>('unified');
  const [collapse, setCollapse] = useState(true);
  const [changedColsOnly, setChangedColsOnly] = useState(false);
  const [theme, setTheme] = useState<Theme>(readTheme);
  const [renderer, setRenderer] = useState<Renderer>('spreadiff');
  const [options, setOptions] = useState<DiffOptions>({ keys: 'auto', trimWhitespace: true, ignoreCase: false, numericTolerance: 0 });
  const [header, setHeader] = useState(true);
  const [selected, setSelected] = useState<CellInfo | null>(null);
  const [showOptions, setShowOptions] = useState(false);
  const [display, setDisplay] = useState<CellDisplay>('stacked');
  const [panel, setPanel] = useState(true);
  const [traceView, setTraceView] = useState<TraceView>('outputs');
  const [pending, setPending] = useState<CellRef | null>(null);
  const source = useRef<Source | null>(null);
  const loadToken = useRef(0);
  const gridRef = useRef<DiffGridHandle>(null);
  const dark = usePrefersDark(theme);

  useEffect(() => {
    document.documentElement.dataset.theme = theme;
    try {
      localStorage.setItem('sd-theme', theme);
    } catch {}
  }, [theme]);

  const run_ = useCallback(async (src: Source, opts: DiffOptions, hdr: boolean, label: string) => {
    // Only the latest request may update the UI; an older one finishing
    // later (or failing) must not overwrite it.
    const token = ++loadToken.current;
    source.current = src;
    setBusy(label);
    setError(null);
    try {
      let r: Loaded;
      if (src.kind === 'files') {
        const [x, y] = await Promise.all([src.a.arrayBuffer(), src.b.arrayBuffer()]);
        r = await run({ type: 'files', old: { name: src.a.name, data: x }, new: { name: src.b.name, data: y }, options: opts, header: hdr }, [x, y]);
      } else if (src.kind === 'sample' || src.kind === 'model') {
        const { old, new: next } = src.kind === 'model' ? sampleModel() : sampleWorkbooks();
        r = await run({ type: 'books', old, new: next, options: opts });
      } else {
        r = await run({ type: 'synthetic', rows: src.rows, cols: 12, options: opts });
      }
      if (token !== loadToken.current) return;
      setLoaded(r);
      setSelected(null);
      // Stay on the same sheet (by name) if it still exists.
      const byName = r.result.sheets.findIndex((s) => s.name === sheetName.current);
      const firstChanged = r.result.sheets.findIndex((s) => s.status !== 'equal');
      const idx = byName >= 0 ? byName : Math.max(0, firstChanged);
      sheetName.current = r.result.sheets[idx]?.name ?? null;
      setSheet(idx);
    } catch (e) {
      if (token === loadToken.current) setError(e instanceof Error ? e.message : String(e));
    } finally {
      if (token === loadToken.current) setBusy(null);
    }
  }, []);

  const diffFiles = (a: File, b: File) => {
    // A key column picked for the previous files may not exist in these.
    const opts = Array.isArray(options.keys) ? { ...options, keys: 'auto' as const } : options;
    if (opts !== options) setOptions(opts);
    void run_({ kind: 'files', a, b }, opts, header, `Diffing ${a.name} → ${b.name}`);
  };

  const loadSample = () => {
    setFiles([null, null]);
    void run_({ kind: 'sample' }, options, header, 'Diffing sample workbooks');
  };

  const loadModel = () => {
    setFiles([null, null]);
    void run_({ kind: 'model' }, options, header, 'Diffing financial model');
  };

  const loadSynthetic = (rows: number) => {
    setFiles([null, null]);
    void run_({ kind: 'synthetic', rows }, options, header, `Generating and diffing ${fmt(rows)} rows`);
  };

  // Re-run the last diff when options change.
  const rerun = (next: DiffOptions, hdr = header) => {
    setOptions(next);
    if (source.current) void run_(source.current, next, hdr, 'Re-diffing');
  };

  const pickSheet = (i: number) => {
    sheetName.current = sheets[i]?.name ?? null;
    setSheet(i);
  };

  // Formula tracing, when any sheet has formulas (cheap: no evaluation).
  const trace = useMemo(() => {
    if (!loaded || !loaded.result.sheets.some((s) => s.diff.newTable.formulas || s.diff.oldTable.formulas)) return null;
    return traceWorkbook(loaded.result);
  }, [loaded]);

  // Jump to a cell from the trace panel (possibly on another sheet).
  const pick = (ref: CellRef) => {
    if (ref.sheet !== sheet) pickSheet(ref.sheet);
    setRenderer((r) => (r === 'glide' ? 'spreadiff' : r));
    setPending(ref);
  };
  useEffect(() => {
    if (!pending || pending.sheet !== sheet) return;
    const id = requestAnimationFrame(() => {
      gridRef.current?.select(pending.row, pending.col);
      setPending(null);
    });
    return () => cancelAnimationFrame(id);
  }, [pending, sheet, loaded]);
  const selectedRef: CellRef | null = selected ? { sheet, row: selected.row, col: selected.columnIndex } : null;

  const setFile = (i: 0 | 1, f: File) => {
    const next: [File | null, File | null] = [...files] as [File | null, File | null];
    next[i] = f;
    setFiles(next);
    if (next[0] && next[1]) void diffFiles(next[0], next[1]);
  };

  const sheets = loaded?.result.sheets ?? [];
  const traceCtx = useMemo(() => (trace && loaded ? { trace, workbook: loaded.result, sheet } : undefined), [trace, loaded, sheet]);
  const current = sheets[sheet];
  const diff = current?.diff;
  const stats = diff?.stats;

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.metaKey || e.ctrlKey || e.altKey) return; // don't hijack ⌘S etc.
      if ((e.target as HTMLElement)?.closest?.('input,select,textarea')) return;
      if (e.key === 'u') setLayout('unified');
      if (e.key === 's') setLayout('split');
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, []);

  const exportCsv = () => {
    if (!diff) return;
    const blob = new Blob([toUnifiedCsv(diff)], { type: 'text/csv' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `${current.name}.diff.csv`;
    a.click();
    // Revoking synchronously can cancel the download in Firefox/Safari.
    setTimeout(() => URL.revokeObjectURL(a.href), 1000);
  };

  const keyChoices = useMemo(() => diff?.columns.filter((c) => c.oldIndex >= 0 && c.newIndex >= 0).map((c) => c.name) ?? [], [diff]);

  return (
    <div className="app">
      <header className="topbar">
        <div className="brand">
          <Logo />
          <span>spreadiff</span>
        </div>
        {loaded && (
          <div className="files">
            <FileChip side="old" name={loaded.names[0]} onPick={files[0] && files[1] ? (f) => setFile(0, f) : undefined} />
            <span className="arrow">→</span>
            <FileChip side="new" name={loaded.names[1]} onPick={files[0] && files[1] ? (f) => setFile(1, f) : undefined} />
          </div>
        )}
        <div className="spacer" />
        {stats && (
          <div className="stats" aria-label="Change summary">
            {stats.rowsAdded > 0 && <span className="stat added">+{fmt(stats.rowsAdded)} <small>rows</small></span>}
            {stats.rowsRemoved > 0 && <span className="stat removed">−{fmt(stats.rowsRemoved)} <small>rows</small></span>}
            {stats.cellsChanged > 0 && <span className="stat modified">~{fmt(stats.cellsChanged)} <small>cells</small></span>}
            {stats.columnsAdded + stats.columnsRemoved + stats.columnsRenamed > 0 && (
              <span className="stat neutral">
                {[stats.columnsAdded && `+${stats.columnsAdded}`, stats.columnsRemoved && `−${stats.columnsRemoved}`, stats.columnsRenamed && `${stats.columnsRenamed} renamed`].filter(Boolean).join(' ')} <small>cols</small>
              </span>
            )}
          </div>
        )}
        <Segmented value={theme} onChange={setTheme} options={[['system', 'Auto'], ['light', 'Light'], ['dark', 'Dark']]} />
      </header>

      {loaded && diff ? (
        <>
          <div className="toolbar">
            <Segmented value={layout} onChange={setLayout} options={[['unified', 'Unified'], ['split', 'Split']]} />
            <Toggle checked={collapse} onChange={setCollapse} label="Collapse unchanged" />
            <Toggle checked={changedColsOnly} onChange={setChangedColsOnly} label="Changed columns only" />
            <label className="select">
              <span>Cells</span>
              <select value={display} onChange={(e) => setDisplay(e.target.value as CellDisplay)}>
                <option value="stacked">Old over new</option>
                <option value="new">New only</option>
                <option value="delta">New + Δ</option>
                <option value="inline">Old → new</option>
              </select>
            </label>
            <label className="select">
              <span>Match rows by</span>
              <select
                value={Array.isArray(options.keys) ? options.keys[0] : options.keys}
                onChange={(e) => {
                  const v = e.target.value;
                  rerun({ ...options, keys: v === 'auto' || v === 'none' ? v : [v] });
                }}
              >
                <option value="auto">Auto{diff.keyColumns.length && options.keys === 'auto' ? ` (${diff.keyColumns.join(', ')})` : ' (order)'}</option>
                <option value="none">Row order</option>
                {keyChoices.map((k) => (
                  <option key={k} value={k}>{k}</option>
                ))}
              </select>
            </label>
            <div className="popover-wrap">
              <button className="btn ghost" onClick={() => setShowOptions((v) => !v)} aria-expanded={showOptions}>
                Options ▾
              </button>
              {showOptions && (
                <div className="popover" onMouseLeave={() => setShowOptions(false)}>
                  <Toggle checked={!!options.ignoreCase} onChange={(v) => rerun({ ...options, ignoreCase: v })} label="Ignore case" />
                  <Toggle checked={options.trimWhitespace !== false} onChange={(v) => rerun({ ...options, trimWhitespace: v })} label="Trim whitespace" />
                  <Toggle checked={header} onChange={(v) => { setHeader(v); rerun(options, v); }} label="First row is header" />
                  <label className="select">
                    <span>Numeric tolerance</span>
                    <select value={String(options.numericTolerance ?? 0)} onChange={(e) => rerun({ ...options, numericTolerance: Number(e.target.value) })}>
                      {['0', '0.000001', '0.001', '0.01', '0.5', '1'].map((v) => (
                        <option key={v} value={v}>{v === '0' ? 'Exact' : `± ${v}`}</option>
                      ))}
                    </select>
                  </label>
                </div>
              )}
            </div>
            <div className="spacer" />
            <div className="nav">
              <button className="btn icon" title="Previous change (p)" onClick={() => gridRef.current?.prevChange()} disabled={renderer !== 'spreadiff'}>↑</button>
              <button className="btn icon" title="Next change (n)" onClick={() => gridRef.current?.nextChange()} disabled={renderer !== 'spreadiff'}>↓</button>
            </div>
            <button className="btn ghost" onClick={exportCsv} title="daff-style highlighter CSV">Export</button>
            <Segmented value={renderer} onChange={setRenderer} options={[['spreadiff', 'Grid'], ['code', 'Code'], ['glide', 'Glide']]} title="Renderer: canvas grid, code-style HTML (spreadiff/lite), or Glide Data Grid" />
            {trace && <Toggle checked={panel} onChange={setPanel} label="Trace" />}
          </div>

          {sheets.length > 1 && (
            <nav className="tabs" role="tablist">
              {sheets.map((s, i) => {
                const st = s.diff.stats;
                const n = st.rowsAdded + st.rowsRemoved + st.rowsModified;
                return (
                  <button key={s.name} role="tab" aria-selected={i === sheet} className={`tab ${s.status}`} onClick={() => pickSheet(i)}>
                    <span className="dot" />
                    {s.name}
                    {s.status === 'added' ? <em>new</em> : s.status === 'removed' ? <em>deleted</em> : n > 0 ? <em>{fmt(n)}</em> : null}
                  </button>
                );
              })}
            </nav>
          )}

          <div className="work">
            <main className="grid-wrap">
              {renderer === 'spreadiff' ? (
                <DiffGrid
                  key={`${loaded.names.join()}:${sheet}`}
                  ref={gridRef}
                  diff={diff}
                  layout={layout}
                  collapseUnchanged={collapse}
                  hideUnchangedColumns={changedColsOnly}
                  cellDisplay={display}
                  trace={traceCtx}
                  theme={theme}
                  onSelectionChange={setSelected}
                  style={{ position: 'absolute', inset: 0 }}
                />
              ) : renderer === 'code' ? (
                <CodeView diff={diff} layout={layout} display={display} changedOnly={changedColsOnly} trace={trace ?? undefined} sheet={sheet} />
              ) : (
                <GlideDiff diff={diff} dark={dark} />
              )}
              {busy && <div className="busy">{busy}…</div>}
            </main>
            {trace && panel && (
              <TracePanel wb={loaded.result} trace={trace} sheet={sheet} selected={selectedRef} view={traceView} onView={setTraceView} onPick={pick} />
            )}
          </div>

          <footer className="statusbar">
            {selected ? <Inspector cell={selected} /> : <span className="hint">Click a cell to inspect · <kbd>n</kbd>/<kbd>p</kbd> next/prev change · <kbd>u</kbd>/<kbd>s</kbd> unified/split · arrows move</span>}
            <div className="spacer" />
            <span className="timing">
              {fmt(diff.oldTable.rows.length)} → {fmt(diff.newTable.rows.length)} rows · {diff.columns.length} cols · {source.current?.kind === 'synthetic' ? 'generate + parse' : 'parse'} {ms(loaded.timings.parse)} · diff {ms(loaded.timings.diff)} · transfer {ms(loaded.timings.transfer)}
            </span>
          </footer>
        </>
      ) : (
        <Landing busy={busy} error={error} files={files} setFile={setFile} loadSample={loadSample} loadModel={loadModel} loadSynthetic={loadSynthetic} />
      )}
      {error && loaded && <div className="toast" onClick={() => setError(null)}>{error}</div>}
    </div>
  );
}

function Landing(props: {
  busy: string | null;
  error: string | null;
  files: [File | null, File | null];
  setFile: (i: 0 | 1, f: File) => void;
  loadSample: () => void;
  loadModel: () => void;
  loadSynthetic: (rows: number) => void;
}) {
  return (
    <div className="landing">
      <h1>Diff spreadsheets, beautifully.</h1>
      <p className="lede">
        Drop two versions of a CSV, TSV or XLSX file. Rows are matched by key or by order, columns by name (renames
        detected), and every changed cell is shown old-over-new. Rendered on canvas; handles millions of cells.
      </p>
      <div className="drops">
        <DropZone side="old" file={props.files[0]} onFile={(f) => props.setFile(0, f)} />
        <div className="drops-arrow">→</div>
        <DropZone side="new" file={props.files[1]} onFile={(f) => props.setFile(1, f)} />
      </div>
      <div className="cta">
        <button className="btn primary" onClick={props.loadModel}>Try a financial model</button>
        <button className="btn" onClick={props.loadSample}>Inventory workbook</button>
        <button className="btn" onClick={() => props.loadSynthetic(100_000)}>Stress test · 100k rows</button>
        <button className="btn" onClick={() => props.loadSynthetic(1_000_000)}>1M rows</button>
      </div>
      {props.busy && <p className="busy-inline">{props.busy}…</p>}
      {props.error && <p className="error">{props.error}</p>}
      <p className="privacy">Files never leave your browser. Parsing and diffing run in a Web Worker.</p>
    </div>
  );
}

function DropZone({ side, file, onFile }: { side: 'old' | 'new'; file: File | null; onFile: (f: File) => void }) {
  const [over, setOver] = useState(false);
  const input = useRef<HTMLInputElement>(null);
  return (
    <div
      className={`drop ${side} ${over ? 'over' : ''} ${file ? 'has' : ''}`}
      onDragOver={(e) => {
        e.preventDefault();
        setOver(true);
      }}
      onDragLeave={() => setOver(false)}
      onDrop={(e) => {
        e.preventDefault();
        setOver(false);
        const fs = [...e.dataTransfer.files];
        if (fs[0]) onFile(fs[0]);
      }}
      onClick={() => input.current?.click()}
      onKeyDown={(e) => {
        if (e.key === 'Enter' || e.key === ' ') {
          e.preventDefault();
          input.current?.click();
        }
      }}
      role="button"
      tabIndex={0}
    >
      <span className="drop-label">{side === 'old' ? 'Before' : 'After'}</span>
      <span className="drop-name">{file ? file.name : 'Drop a file or click to choose'}</span>
      <span className="drop-hint">.csv · .tsv · .xlsx</span>
      <input ref={input} type="file" accept=".csv,.tsv,.txt,.xlsx,.xlsm" hidden onChange={(e) => {
        const f = e.target.files?.[0];
        e.target.value = ''; // so choosing the same file again still fires
        if (f) onFile(f);
      }} />
    </div>
  );
}

function FileChip({ side, name, onPick }: { side: 'old' | 'new'; name: string; onPick?: (f: File) => void }) {
  const input = useRef<HTMLInputElement>(null);
  // Replacing one side only makes sense when both sides are real files.
  if (!onPick) {
    return (
      <span className={`file-chip ${side}`}>
        <span className="dot" />
        {name}
      </span>
    );
  }
  return (
    <button className={`file-chip ${side}`} onClick={() => input.current?.click()} title="Replace file">
      <span className="dot" />
      {name}
      <input ref={input} type="file" accept=".csv,.tsv,.txt,.xlsx,.xlsm" hidden onChange={(e) => {
        const f = e.target.files?.[0];
        e.target.value = '';
        if (f) onPick(f);
      }} />
    </button>
  );
}

function Inspector({ cell }: { cell: CellInfo }) {
  const before = formatCell(cell.oldValue ?? null);
  const after = formatCell(cell.newValue ?? null);
  const delta = cell.changed ? numericDelta(cell.oldValue ?? null, cell.newValue ?? null) : null;
  // Character highlights help in text; in numbers they're noise.
  const numeric = typeof cell.oldValue === 'number' && typeof cell.newValue === 'number';
  const spans = cell.changed && !numeric && before.length + after.length < 4000 ? diffText(before, after) : null;
  return (
    <div className="inspector">
      <span className={`badge ${cell.status}`}>{cell.status}</span>
      {cell.kind && <span className={`kchip k-${cell.kind}`}>{cell.kind === 'computed' ? 'downstream' : cell.kind}</span>}
      <b>{cell.column.name}</b>
      {cell.changed ? (
        <span className="vals">
          <span className="old">{spans ? spans.filter((s) => s.kind !== 'added').map((s, i) => (s.kind === 'removed' ? <del key={i}>{s.text}</del> : <span key={i}>{s.text}</span>)) : before || '∅'}</span>
          <span className="arrow">→</span>
          <span className="new">{spans ? spans.filter((s) => s.kind !== 'removed').map((s, i) => (s.kind === 'added' ? <ins key={i}>{s.text}</ins> : <span key={i}>{s.text}</span>)) : after || '∅'}</span>
          {delta && delta.delta !== 0 && (
            <span className={`delta ${delta.delta > 0 ? 'up' : 'down'}`}>
              {delta.delta > 0 ? '+' : '−'}
              {formatCell(Math.abs(delta.delta))}
              {delta.ratio !== null && ` (${(Math.abs(delta.ratio) * 100).toFixed(1)}%)`}
            </span>
          )}
        </span>
      ) : (
        <span className="vals">{(cell.newValue !== undefined ? after : before) || '∅'}</span>
      )}
      {(cell.oldFormula || cell.newFormula) &&
        (cell.oldFormula === cell.newFormula || cell.kind === 'computed' ? (
          <code>ƒ {cell.newFormula ?? cell.oldFormula}</code>
        ) : (
          <code>{cell.oldFormula ?? '—'} → {cell.newFormula ?? '—'}</code>
        ))}
    </div>
  );
}

function Segmented<T extends string>({ value, onChange, options, title }: { value: T; onChange: (v: T) => void; options: [T, string][]; title?: string }) {
  return (
    <div className="segmented" role="radiogroup" title={title}>
      {options.map(([v, label]) => (
        <button key={v} role="radio" aria-checked={v === value} className={v === value ? 'on' : ''} onClick={() => onChange(v)}>
          {label}
        </button>
      ))}
    </div>
  );
}

function Toggle({ checked, onChange, label }: { checked: boolean; onChange: (v: boolean) => void; label: string }) {
  return (
    <label className="toggle">
      <input type="checkbox" checked={checked} onChange={(e) => onChange(e.target.checked)} />
      <span className="track"><span className="thumb" /></span>
      {label}
    </label>
  );
}

function Logo() {
  return (
    <svg width="22" height="22" viewBox="0 0 32 32" aria-hidden>
      <rect width="32" height="32" rx="8" fill="currentColor" opacity="0.08" />
      <rect x="6" y="8" width="9" height="5" rx="1.5" fill="var(--removed)" />
      <rect x="17" y="8" width="9" height="5" rx="1.5" fill="var(--added)" />
      <rect x="6" y="15" width="20" height="3.5" rx="1.2" fill="currentColor" opacity="0.35" />
      <rect x="6" y="20.5" width="20" height="3.5" rx="1.2" fill="currentColor" opacity="0.2" />
    </svg>
  );
}
