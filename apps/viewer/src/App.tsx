import { useCallback, useEffect, useMemo, useRef, useState, type ReactNode } from 'react';
import { applyRecalculation, diffText, formatCell, numericDelta, toUnifiedCsv, unpackWorkbookDiff, unpackWorkbookTrace, type CellRef, type DiffOptions, type WorkbookDiff, type WorkbookTrace } from 'spreadiff';
import { DiffGrid, type CellDisplay, type CellInfo, type DiffGridHandle, type DiffLayout } from 'spreadiff/react';
import { CodeView } from './CodeView';
import { Icon } from './Icon';
import { GlideDiff } from './GlideDiff';
import { sampleModel, sampleWorkbooks } from './samples';
import { TracePanel, type TraceView } from './TracePanel';
import type { Req } from './worker';
import DiffWorker from './worker?worker';

type Theme = 'system' | 'light' | 'dark';
type Renderer = 'spreadiff' | 'code' | 'glide';

interface Loaded {
  result: WorkbookDiff;
  /** Traced in the worker, when the files have formulas. */
  trace: WorkbookTrace | null;
  names: [string, string];
  timings: { parse: number; diff: number; trace: number; transfer: number };
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
      const trace = e.data.trace ? unpackWorkbookTrace(e.data.trace, result) : null;
      const total = performance.now() - t0;
      const { parse, diff, trace: traced } = e.data.timings;
      // "transfer" = pack in worker + postMessage + unpack on main thread.
      resolve({ result, trace, names: e.data.names, timings: { parse, diff, trace: traced, transfer: Math.max(0, total - parse - diff - traced) } });
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
  const [header, setHeader] = useState<boolean | 'auto'>('auto');
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

  const run_ = useCallback(async (src: Source, opts: DiffOptions, hdr: boolean | 'auto', label: string) => {
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

  // Formula tracing, done in the worker when any sheet has formulas.
  const trace = loaded?.trace ?? null;
  // Stale results in the new file: show what recalculating gives, unless asked for the saved values.
  const [savedValues, setSavedValues] = useState(false);
  const shown = useMemo(() => (loaded && trace && !savedValues ? applyRecalculation(loaded.result, trace) : loaded?.result ?? null), [loaded, trace, savedValues]);

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

  const sheets = shown?.sheets ?? [];
  const traceCtx = useMemo(() => (trace && shown ? { trace, workbook: shown, sheet } : undefined), [trace, shown, sheet]);
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

  const keyLabel = diff ? (diff.keyColumns.length ? diff.keyColumns.join(' + ') : 'row order') : '';
  const columnSummary = stats
    ? [
        stats.columnsAdded && plural(stats.columnsAdded, 'column') + ' added',
        stats.columnsRemoved && plural(stats.columnsRemoved, 'column') + ' removed',
        stats.columnsRenamed && plural(stats.columnsRenamed, 'column') + ' renamed',
      ].filter(Boolean)
    : [];

  return (
    <div className="app">
      <header className="topbar">
        <div className="brand">
          <Logo />
          <h1>spreadiff</h1>
        </div>
        {loaded && (
          <div className="files">
            <FileChip side="old" name={loaded.names[0]} onPick={files[0] && files[1] ? (f) => setFile(0, f) : undefined} />
            <Icon name="arrow-right" className="arrow" />
            <FileChip side="new" name={loaded.names[1]} onPick={files[0] && files[1] ? (f) => setFile(1, f) : undefined} />
          </div>
        )}
        <div className="spacer" />
        {loaded && (
          <button className="btn ghost small" onClick={() => { setLoaded(null); setFiles([null, null]); source.current = null; }} title="Compare other files">
            New diff
          </button>
        )}
        <Segmented
          value={theme}
          onChange={setTheme}
          label="Theme"
          options={[['system', <Icon name="monitor" />, 'Match system'], ['light', <Icon name="sun" />, 'Light'], ['dark', <Icon name="moon" />, 'Dark']]}
        />
      </header>

      {loaded && diff ? (
        <main className="main">
          <div className="toolbar" role="toolbar" aria-label="View">
            <Segmented
              value={layout}
              onChange={setLayout}
              label="Layout"
              options={[['unified', <><Icon name="unified" /><span className="label">Unified</span></>, 'Unified: one grid, old and new together (u)'], ['split', <><Icon name="split" /><span className="label">Split</span></>, 'Split: old and new side by side (s)']]}
            />
            <Segmented
              value={display}
              onChange={setDisplay}
              label="Changed cells"
              options={[
                ['stacked', 'Stacked', 'Old value over new'],
                ['new', 'New', 'New value only'],
                ['delta', 'Δ %', 'New value with absolute and percent change'],
                ['inline', 'Inline', 'Old → new on one line'],
              ]}
            />
            <span className="divider" />
            <Toggle checked={collapse} onChange={setCollapse} label="Collapse unchanged" />
            <div className="spacer" />
            <div className="joined" role="group" aria-label="Changes">
              <button className="btn icon" title="Previous change (p)" aria-label="Previous change" onClick={() => gridRef.current?.prevChange()} disabled={renderer !== 'spreadiff'}><Icon name="up" /></button>
              <button className="btn icon" title="Next change (n)" aria-label="Next change" onClick={() => gridRef.current?.nextChange()} disabled={renderer !== 'spreadiff'}><Icon name="down" /></button>
            </div>
            <div className="popover-wrap">
              <button className={`btn ${showOptions ? 'on' : ''}`} onClick={() => setShowOptions((v) => !v)} aria-expanded={showOptions}>
                <Icon name="sliders" />
                Options
                {(changedColsOnly || options.ignoreCase || options.trimWhitespace === false || header !== 'auto' || options.numericTolerance || options.keys !== 'auto') && <span className="pip" title="Some options differ from the defaults" />}
              </button>
              {showOptions && (
                <>
                  <div className="popover-scrim" onClick={() => setShowOptions(false)} />
                  <div className="popover" role="dialog" aria-label="Options">
                    <span className="popover-title">View</span>
                    <Toggle checked={changedColsOnly} onChange={setChangedColsOnly} label="Changed columns only" />
                    <hr />
                    <span className="popover-title">Matching</span>
                    <label className="field">
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
                    <label className="field">
                      <span>First row</span>
                      <select value={String(header)} title="Auto: column names, unless the row holds numbers (years aside), dates or formulas, or formulas read it" onChange={(e) => { const v = e.target.value === 'auto' ? 'auto' : e.target.value === 'true'; setHeader(v); rerun(options, v); }}>
                        <option value="auto">Auto</option>
                        <option value="true">Column names</option>
                        <option value="false">Data</option>
                      </select>
                    </label>
                    <label className="field">
                      <span>Numbers equal within</span>
                      <select value={String(options.numericTolerance ?? 0)} onChange={(e) => rerun({ ...options, numericTolerance: Number(e.target.value) })}>
                        {['0', '0.000001', '0.001', '0.01', '0.5', '1'].map((v) => (
                          <option key={v} value={v}>{v === '0' ? 'Exact' : `± ${v}`}</option>
                        ))}
                      </select>
                    </label>
                    <hr />
                    <Toggle checked={!!options.ignoreCase} onChange={(v) => rerun({ ...options, ignoreCase: v })} label="Ignore case" />
                    <Toggle checked={options.trimWhitespace !== false} onChange={(v) => rerun({ ...options, trimWhitespace: v })} label="Trim whitespace" />
                    <hr className="tight" />
                    <button className="btn ghost small left" onClick={exportCsv} title="daff-style highlighter CSV">
                      <Icon name="download" />
                      Export diff as CSV
                    </button>
                  </div>
                </>
              )}
            </div>
            <span className="divider" />
            <Segmented
              value={renderer}
              onChange={setRenderer}
              label="Renderer"
              options={[['spreadiff', 'Grid', 'Canvas grid'], ['code', 'Code', 'Code-review HTML (spreadiff/lite)'], ['glide', 'Glide', 'Glide Data Grid, for comparison']]}
            />
            {trace && (
              <button className={`btn icon ${panel ? 'on' : ''}`} onClick={() => setPanel((v) => !v)} aria-pressed={panel} title={panel ? 'Hide trace panel' : 'Show trace panel'} aria-label="Trace panel">
                <Icon name="panel" />
              </button>
            )}
          </div>

          <div className="sheetbar">
            <nav className="tabs" role="tablist" aria-label="Sheets">
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
            {stats && (
              <div className="summary" aria-label="Change summary">
                {stats.rowsAdded > 0 && <span className="sum added"><b>+{fmt(stats.rowsAdded)}</b> {stats.rowsAdded === 1 ? 'row' : 'rows'}</span>}
                {stats.rowsRemoved > 0 && <span className="sum removed"><b>−{fmt(stats.rowsRemoved)}</b> {stats.rowsRemoved === 1 ? 'row' : 'rows'}</span>}
                {stats.cellsChanged > 0 && <span className="sum modified"><b>{fmt(stats.cellsChanged)}</b> {stats.cellsChanged === 1 ? 'cell' : 'cells'} changed</span>}
                {columnSummary.length > 0 && <span className="sum">{columnSummary.join(' · ')}</span>}
                {stats.rowsAdded + stats.rowsRemoved + stats.cellsChanged + columnSummary.length === 0 && <span className="sum">No changes</span>}
                <span className="sum faint" title="How rows were matched (change it under Options)">matched by {keyLabel}</span>
              </div>
            )}
          </div>

          {(trace?.stats.stale || 0) > 0 && (
            <div className="banner">
              <b>{plural(trace!.stats.stale, 'formula result')} in the new file {trace!.stats.stale === 1 ? 'was' : 'were'} saved without recalculating.</b>
              <span>Their inputs changed but their saved values didn't (manual calculation mode, or a script edited the file). {savedValues ? 'Showing the saved values.' : 'Showing what recalculating gives.'}</span>
              <button className="btn ghost small" onClick={() => setSavedValues((v) => !v)}>{savedValues ? 'Show recalculated' : 'Show saved values'}</button>
            </div>
          )}
          {missingNote(loaded.result) && <div className="banner info"><span>{missingNote(loaded.result)}</span></div>}
          <div className="work">
            <section className="grid-wrap" aria-label="Diff">
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
              {busy && <div className="busy"><span className="spinner" />{busy}…</div>}
            </section>
            {trace && panel && (
              <TracePanel wb={shown!} trace={trace} sheet={sheet} selected={selectedRef} view={traceView} onView={setTraceView} onPick={pick} />
            )}
          </div>

          <footer className="statusbar">
            {selected ? <Inspector cell={selected} /> : <span className="hint">Click a cell to inspect · <kbd>n</kbd> <kbd>p</kbd> next / previous change · <kbd>u</kbd> <kbd>s</kbd> unified / split · arrows move</span>}
            <div className="spacer" />
            <span className="timing">
              {fmt(diff.oldTable.rows.length)} → {fmt(diff.newTable.rows.length)} rows · {diff.columns.length} cols · {source.current?.kind === 'synthetic' ? 'generate + parse' : 'parse'} {ms(loaded.timings.parse)} · diff {ms(loaded.timings.diff)}{loaded.trace ? ` · trace ${ms(loaded.timings.trace)}` : ''} · transfer {ms(loaded.timings.transfer)}
            </span>
          </footer>
        </main>
      ) : (
        <Landing busy={busy} error={error} files={files} setFile={setFile} loadSample={loadSample} loadModel={loadModel} loadSynthetic={loadSynthetic} />
      )}
      {error && loaded && <div className="toast" onClick={() => setError(null)}>{error}</div>}
    </div>
  );
}

const plural = (n: number, word: string) => `${fmt(n)} ${word}${n === 1 ? '' : 's'}`;

/** A note when a file had formulas without saved values (written by a tool that doesn't calculate). */
function missingNote(d: WorkbookDiff): string | null {
  const parts = (['old', 'new'] as const).flatMap((side) => {
    const m = d.computedValues?.[side];
    if (!m) return [];
    const which = side === 'old' ? 'The old file' : 'The new file';
    return [`${which} had no saved results for ${plural(m.computed + m.failed, 'formula')}; spreadiff computed ${m.failed ? `${fmt(m.computed)} of them (${fmt(m.failed)} use functions it can't evaluate and show empty)` : 'them'}.`];
  });
  return parts.length ? parts.join(' ') : null;
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
    <main className="landing">
      <div className="hero">
        <h1>Diff spreadsheets, beautifully.</h1>
        <p className="lede">
          Compare two versions of a CSV, TSV or XLSX file. Rows match by key or by order, columns by name, and in
          workbooks with formulas every changed result is traced back to the edits that caused it.
        </p>
      </div>
      <div className="drops">
        <DropZone side="old" file={props.files[0]} onFile={(f) => props.setFile(0, f)} />
        <div className="drops-arrow"><Icon name="arrow-right" /></div>
        <DropZone side="new" file={props.files[1]} onFile={(f) => props.setFile(1, f)} />
      </div>
      <div className="samples">
        <span className="samples-label">Or try a sample</span>
        <div className="cta">
          <button className="sample primary" onClick={props.loadModel}>
            <b>Financial model</b>
            <span>Formulas across three sheets, with downstream changes traced</span>
          </button>
          <button className="sample" onClick={props.loadSample}>
            <b>Inventory workbook</b>
            <span>Keyed rows, a renamed column, added and removed rows</span>
          </button>
          <button className="sample" onClick={() => props.loadSynthetic(100_000)}>
            <b>100k rows</b>
            <span>Generated, 12 columns</span>
          </button>
          <button className="sample" onClick={() => props.loadSynthetic(1_000_000)}>
            <b>1M rows</b>
            <span>Takes a few seconds to generate</span>
          </button>
        </div>
      </div>
      {props.busy && <p className="busy-inline"><span className="spinner" />{props.busy}…</p>}
      {props.error && <p className="error">{props.error}</p>}
      <p className="privacy"><Icon name="lock" />Files never leave your browser. Parsing and diffing run in a Web Worker.</p>
    </main>
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
      <span className="drop-icon"><Icon name={file ? 'file' : 'upload'} /></span>
      <span className="drop-text">
        <span className="drop-label">{side === 'old' ? 'Before' : 'After'}</span>
        <span className="drop-name">{file ? file.name : 'Drop a file, or click to choose'}</span>
        <span className="drop-hint">CSV, TSV or XLSX</span>
      </span>
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

/** Options are `[value, content, description]`; the description is the tooltip and, for icon-only content, the label. */
function Segmented<T extends string>({ value, onChange, options, label }: { value: T; onChange: (v: T) => void; options: [T, ReactNode, string?][]; label: string }) {
  return (
    <div className="segmented" role="radiogroup" aria-label={label}>
      {options.map(([v, content, description]) => (
        <button
          key={v}
          role="radio"
          aria-checked={v === value}
          aria-label={typeof content === 'string' ? undefined : description}
          title={description}
          className={v === value ? 'on' : ''}
          onClick={() => onChange(v)}
        >
          {content}
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
