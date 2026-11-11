import { MAX_COL, MAX_ROW } from '../engine/address.js';
import { frozenColEnd, frozenColStart, frozenRowEnd, frozenRowStart } from './freeze-band.js';
import { isRowHidden } from './row-visibility.js';
import type { State } from './types.js';

const firstShown = (
  from: number,
  to: number,
  hidden: (index: number) => boolean,
): number | null => {
  for (let i = from; i < to; i += 1) if (!hidden(i)) return i;
  return null;
};

/** The leftmost column on screen: the frozen band's first shown column, else the scrolling
 *  pane's. A row header click makes it the active column. */
export function firstVisibleCol(state: State): number {
  const { layout, viewport } = state;
  const hidden = (col: number): boolean => layout.hiddenCols.has(col);
  const frozen = firstShown(frozenColStart(layout), frozenColEnd(layout), hidden);
  if (frozen !== null) return frozen;
  const start = Math.max(
    frozenColEnd(layout),
    viewport.colStart,
    viewport.navigationRange?.c0 ?? 0,
  );
  return firstShown(start, MAX_COL + 1, hidden) ?? start;
}

/** The top row on screen: the frozen band's first shown row, else the scrolling pane's. A
 *  column header click makes it the active row. */
export function firstVisibleRow(state: State): number {
  const { layout, viewport } = state;
  const hidden = (row: number): boolean => isRowHidden(layout, row);
  const frozen = firstShown(frozenRowStart(layout), frozenRowEnd(layout), hidden);
  if (frozen !== null) return frozen;
  const start = Math.max(
    frozenRowEnd(layout),
    viewport.rowStart,
    viewport.navigationRange?.r0 ?? 0,
  );
  return firstShown(start, MAX_ROW + 1, hidden) ?? start;
}
