import type { WorkbookHandle } from '../../engine/workbook-handle.js';
import type { State } from '../../store/store.js';
import { type CopyResult, copy } from './copy.js';

/**
 * Copy the selection to the clipboard, then blank the source range. Excel
 * normally defers the blank until the next paste — we apply it eagerly
 * here for v1.0 to keep the wb state visibly in sync. The dotted-marquee
 * UX is a v1.x concern.
 */
export function cut(state: State, wb: WorkbookHandle): CopyResult | null {
  const result = copy(state);
  if (!result) return null;
  const { range } = result;
  const sheet = state.data.sheetIndex;
  for (let row = range.r0; row <= range.r1; row += 1) {
    for (let col = range.c0; col <= range.c1; col += 1) {
      wb.setBlank({ sheet, row, col });
    }
  }
  return result;
}
