'use client';

import {
  forwardRef,
  useEffect,
  useImperativeHandle,
  useLayoutEffect,
  useRef,
  useState,
  type CSSProperties,
  type ReactNode,
} from 'react';
import { createPortal } from 'react-dom';
import { DiffGrid as DiffGridView, type CellAnnotation, type DiffGridOptions } from '../render/DiffGrid.js';
import type { TableDiff } from '../core/types.js';
import type { Comparison, SheetView, SpreadsheetInput } from '../core/compare.js';
import { SpreadsheetDiff as SpreadsheetDiffView, type SpreadsheetDiffOptions } from '../render/SpreadsheetDiff.js';

/** Props of the React {@link DiffGrid}: every {@link DiffGridOptions} field, `diff` or `sheet`, plus these. */
export type DiffGridProps<A = unknown> = DiffGridBaseProps<A> &
  (
    | {
        /** The diff to show. Changing it resets scroll and selection. */
        diff: TableDiff;
        sheet?: undefined;
      }
    | {
        /** A sheet from `diffFiles` / `compareWorkbooks`: its diff, with its trace attached. */
        sheet: SheetView;
        diff?: undefined;
      }
  );

/** Props of the React {@link DiffGrid} other than what to show. */
export interface DiffGridBaseProps<A = unknown> extends Omit<DiffGridOptions<A>, 'renderAnnotation'> {
  /** Keep this array referentially stable (useMemo); a new array re-renders every annotation. */
  annotations?: CellAnnotation<A>[];
  /** Class for the container element. */
  className?: string;
  /** Style for the container element. The grid fills it: give it a height. */
  style?: CSSProperties;
  /** Render annotation content with React; it is portaled into place below the annotated row. */
  renderAnnotation?: (annotation: CellAnnotation<A>) => ReactNode;
}

/**
 * The `ref` of the React {@link DiffGrid}: the imperative methods of the
 * vanilla grid, and the instance itself for anything else.
 */
export type DiffGridHandle = Pick<
  DiffGridView,
  'nextChange' | 'prevChange' | 'scrollToRow' | 'expandAll' | 'collapseAll' | 'select' | 'refreshTheme' | 'cellInfo'
> & { readonly instance: DiffGridView | null };

/**
 * React wrapper around the canvas `DiffGrid`. Props are the grid's options;
 * changing them updates the grid in place (only layout-affecting changes
 * re-layout). Give the element a height.
 *
 * ```tsx
 * import { DiffGrid, type DiffGridHandle } from 'spreadiff/react';
 * import { useRef } from 'react';
 * import type { TableDiff } from 'spreadiff';
 *
 * function Review({ diff }: { diff: TableDiff }) {
 *   const grid = useRef<DiffGridHandle>(null);
 *   return (
 *     <>
 *       <button onClick={() => grid.current?.nextChange()}>Next change</button>
 *       <DiffGrid ref={grid} diff={diff} layout="split" cellDisplay="delta" style={{ height: 600 }} />
 *     </>
 *   );
 * }
 * ```
 */
export const DiffGrid = forwardRef(function DiffGrid<A>(
  { diff: diffProp, sheet, className, style, renderAnnotation, ...props }: DiffGridProps<A>,
  ref: React.ForwardedRef<DiffGridHandle>,
) {
  const containerRef = useRef<HTMLDivElement>(null);
  const viewRef = useRef<DiffGridView<A> | null>(null);
  const [portals, setPortals] = useState<Map<CellAnnotation<A>, HTMLElement>>(new Map());
  // Stable React keys per annotation object (row:column can repeat).
  const keys = useRef(new WeakMap<CellAnnotation<A>, number>());
  const nextKey = useRef(0);
  const keyOf = (a: CellAnnotation<A>) => {
    let k = keys.current.get(a);
    if (k === undefined) keys.current.set(a, (k = nextKey.current++));
    return k;
  };

  const diff = sheet ? sheet.diff : diffProp!;
  // A sheet brings its trace; an explicit `trace` prop wins.
  const options = sheet && props.trace === undefined ? { ...props, trace: sheet.trace } : props;
  const vanillaRender = renderAnnotation
    ? (a: CellAnnotation<A>) => {
        const host = document.createElement('div');
        setPortals((prev) => new Map(prev).set(a, host));
        return host;
      }
    : undefined;

  useLayoutEffect(() => {
    const view = new DiffGridView<A>({ ...options, renderAnnotation: vanillaRender });
    viewRef.current = view;
    view.render({ diff, container: containerRef.current! });
    return () => {
      view.cleanUp();
      viewRef.current = null;
      setPortals(new Map());
    };
    // Created once; later prop changes go through setOptions/setDiff.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const { annotations, ...rest } = options;
  // Options first: a new diff should meet its own trace, not the previous one.
  useEffect(() => {
    viewRef.current?.setOptions(rest);
  });
  useEffect(() => {
    viewRef.current?.setDiff(diff);
  }, [diff]);
  useEffect(() => {
    if (annotations === undefined) return;
    setPortals((prev) => {
      const keep = new Set(annotations);
      const next = new Map([...prev].filter(([a]) => keep.has(a)));
      return next.size === prev.size ? prev : next;
    });
    viewRef.current?.setOptions({ annotations, renderAnnotation: vanillaRender });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [annotations]);

  useImperativeHandle(
    ref,
    () => ({
      get instance() {
        return viewRef.current as DiffGridView | null;
      },
      nextChange: () => viewRef.current?.nextChange() ?? false,
      prevChange: () => viewRef.current?.prevChange() ?? false,
      scrollToRow: (row, align) => viewRef.current?.scrollToRow(row, align),
      expandAll: () => viewRef.current?.expandAll(),
      collapseAll: () => viewRef.current?.collapseAll(),
      select: (row, col) => viewRef.current?.select(row, col) ?? false,
      refreshTheme: () => viewRef.current?.refreshTheme(),
      cellInfo: (row, col) => viewRef.current!.cellInfo(row, col),
    }),
    [],
  );

  return (
    <div ref={containerRef} className={className} style={{ position: 'relative', ...style }}>
      {renderAnnotation && [...portals].map(([a, host]) => createPortal(renderAnnotation(a), host, keyOf(a)))}
    </div>
  );
}) as <A = unknown>(props: DiffGridProps<A> & { ref?: React.Ref<DiffGridHandle> }) => ReactNode;

/** Props of the React {@link SpreadsheetDiff}: every {@link SpreadsheetDiffOptions} field, what to compare, plus these. */
export interface SpreadsheetDiffProps<A = unknown> extends Omit<SpreadsheetDiffOptions<A>, 'renderSheetHeader' | 'renderAnnotation'> {
  /** The old version (File, Blob, URL, bytes, `{ name, data }`). With `newFile`; compared when either changes. */
  oldFile?: SpreadsheetInput;
  /** The new version. */
  newFile?: SpreadsheetInput;
  /** Or a comparison you already have (from `diffFiles`). */
  comparison?: Comparison;
  /** Class for the container element. */
  className?: string;
  /** Style for the container element. The component fills it: give it a height. */
  style?: CSSProperties;
  /** Content for each sheet's header, rendered with React. */
  renderSheetHeader?: (sheet: SheetView, comparison: Comparison) => ReactNode;
}

/** The `ref` of the React {@link SpreadsheetDiff}: the vanilla component's methods. */
export type SpreadsheetDiffHandle = Pick<SpreadsheetDiffView, 'showSheet' | 'reveal' | 'getDecisions' | 'setDecisions' | 'decideSelection' | 'exportXlsx' | 'exportCsv' | 'mergedTable'> & {
  readonly instance: SpreadsheetDiffView | null;
};

/** What identifies an input across renders: a URL by its address, anything else by identity. */
const inputKey = (input: SpreadsheetInput | undefined): unknown => (input instanceof URL ? `url:${input.href}` : input);

/**
 * Two spreadsheets in, the whole review out: sheet tabs, the traced grid,
 * the outputs panel and, with `review`, accept/reject and the merged
 * download. Give the element a height.
 *
 * ```tsx
 * import { SpreadsheetDiff } from 'spreadiff/react';
 *
 * export const Review = () => <SpreadsheetDiff oldFile={new URL('/v1.xlsx', location.href)} newFile={new URL('/v2.xlsx', location.href)} style={{ height: 600 }} />;
 * ```
 */
export const SpreadsheetDiff = forwardRef(function SpreadsheetDiff<A>(
  { oldFile, newFile, comparison, className, style, renderSheetHeader, ...options }: SpreadsheetDiffProps<A>,
  ref: React.ForwardedRef<SpreadsheetDiffHandle>,
) {
  const containerRef = useRef<HTMLDivElement>(null);
  const viewRef = useRef<SpreadsheetDiffView<A> | null>(null);
  const [header, setHeader] = useState<{ host: HTMLElement; sheet: SheetView; cmp: Comparison } | null>(null);
  const headerFn = renderSheetHeader
    ? (sheet: SheetView, cmp: Comparison) => {
        const host = document.createElement('div');
        setHeader({ host, sheet, cmp });
        return host;
      }
    : undefined;

  useLayoutEffect(() => {
    const view = new SpreadsheetDiffView<A>({ ...options, renderSheetHeader: headerFn });
    viewRef.current = view;
    return () => {
      view.cleanUp();
      viewRef.current = null;
    };
    // Created once; later prop changes go through setOptions and render.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);
  useEffect(() => {
    viewRef.current?.setOptions({ ...options, renderSheetHeader: headerFn });
  });
  useEffect(() => {
    const view = viewRef.current;
    const container = containerRef.current;
    if (!view || !container) return;
    if (comparison) void view.render({ comparison, container });
    else if (oldFile && newFile) view.render({ oldFile, newFile, container }).catch(() => {}); // shown, and passed to onError
    // A URL written inline (`oldFile={new URL(…)}`) is a new object each render: compare by address, or every render would compare again.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [inputKey(oldFile), inputKey(newFile), comparison]);

  useImperativeHandle(
    ref,
    () => ({
      get instance() {
        return viewRef.current as SpreadsheetDiffView | null;
      },
      showSheet: (i) => viewRef.current?.showSheet(i),
      reveal: (r) => viewRef.current?.reveal(r),
      getDecisions: () => viewRef.current?.getDecisions() ?? [],
      setDecisions: (d) => viewRef.current?.setDecisions(d),
      decideSelection: (d) => viewRef.current?.decideSelection(d),
      exportXlsx: () => viewRef.current!.exportXlsx(),
      exportCsv: (sheet) => viewRef.current!.exportCsv(sheet),
      mergedTable: (sheet) => viewRef.current!.mergedTable(sheet),
    }),
    [],
  );

  return (
    <div ref={containerRef} className={className} style={{ position: 'relative', ...style }}>
      {renderSheetHeader && header && createPortal(renderSheetHeader(header.sheet, header.cmp), header.host)}
    </div>
  );
}) as <A = unknown>(props: SpreadsheetDiffProps<A> & { ref?: React.Ref<SpreadsheetDiffHandle> }) => ReactNode;

export type { CellAnnotation, CellDisplay, CellInfo, DiffGridOptions, DiffLayout, GridRange, TraceContext } from '../render/DiffGrid.js';
export type { SpreadsheetDiffOptions } from '../render/SpreadsheetDiff.js';
export type { DiffStrings } from '../render/strings.js';
export type { Comparison, SheetView, SpreadsheetInput } from '../core/compare.js';
