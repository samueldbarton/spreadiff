import { useMemo } from 'react';
import { HTML_CSS, renderHTML, type HtmlDisplay, type HtmlLayout } from 'spreadiff/html';
import type { TableDiff, WorkbookTrace } from 'spreadiff';

/** The diff rendered by `spreadiff/html`: plain HTML, code-review style. */
export function CodeView(props: { diff: TableDiff; layout: HtmlLayout; display: HtmlDisplay; changedOnly: boolean; trace?: WorkbookTrace; sheet: number }) {
  const html = useMemo(
    () =>
      renderHTML(props.diff, {
        layout: props.layout,
        display: props.display,
        changedColumnsOnly: props.changedOnly,
        kindOf: props.trace ? (row, col) => props.trace!.get({ sheet: props.sheet, row, col })?.kind : undefined,
      }),
    [props.diff, props.layout, props.display, props.changedOnly, props.trace, props.sheet],
  );
  return (
    <div className="code-view">
      <style>{HTML_CSS}</style>
      <div dangerouslySetInnerHTML={{ __html: html }} />
    </div>
  );
}
