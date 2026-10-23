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

export interface DiffGridProps<A = unknown> extends Omit<DiffGridOptions<A>, 'renderAnnotation'> {
  diff: TableDiff;
  /** Keep this array referentially stable (useMemo); a new array re-renders every annotation. */
  annotations?: CellAnnotation<A>[];
  className?: string;
  style?: CSSProperties;
  /** Render annotation content with React; it is portaled below the annotated row. */
  renderAnnotation?: (annotation: CellAnnotation<A>) => ReactNode;
}

export type DiffGridHandle = Pick<
  DiffGridView,
  'nextChange' | 'prevChange' | 'scrollToRow' | 'expandAll' | 'collapseAll' | 'select' | 'refreshTheme' | 'cellInfo'
> & { readonly instance: DiffGridView | null };

/**
 * React wrapper around the canvas `DiffGrid`. Give the element a height.
 *
 * ```tsx
 * <DiffGrid diff={diff} layout="split" style={{ height: 600 }} />
 * ```
 */
export const DiffGrid = forwardRef(function DiffGrid<A>(
  { diff, className, style, renderAnnotation, ...options }: DiffGridProps<A>,
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

  useEffect(() => {
    viewRef.current?.setDiff(diff);
  }, [diff]);

  const { annotations, ...rest } = options;
  useEffect(() => {
    viewRef.current?.setOptions(rest);
  });
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

export type { CellAnnotation, CellInfo, DiffGridOptions, DiffLayout } from '../render/DiffGrid.js';
