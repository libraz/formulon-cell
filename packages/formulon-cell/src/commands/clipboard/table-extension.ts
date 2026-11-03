import type { Range } from '../../engine/types.js';
import type { SpreadsheetStore } from '../../store/store.js';
import { updateTableOverlay } from '../format-as-table.js';
import type { History } from '../history.js';
import { recordTablesChange } from '../slice-history.js';

/**
 * Grow every table a paste lands against. A paste that starts in the row
 * directly below a table, within its columns, adds the rows it covers; one
 * that starts in the column directly right of it, within its rows, adds the
 * columns. A table with a total row does not grow downward. Table A1:B3 grows
 * to A1:B4 for a paste at A4:B4, and A1:B5 grows to A1:C5 for one at C1:C3.
 */
export function extendTablesForPaste(
  store: SpreadsheetStore,
  history: History | null,
  written: Range,
): void {
  const growth = store.getState().tables.tables.flatMap((table) => {
    const t = table.range;
    if (t.sheet !== written.sheet) return [];
    const below =
      !table.showTotal && written.r0 === t.r1 + 1 && written.c0 >= t.c0 && written.c0 <= t.c1;
    const right = written.c0 === t.c1 + 1 && written.r0 >= t.r0 && written.r0 <= t.r1;
    if (below) return [{ id: table.id, range: { ...t, r1: Math.max(t.r1, written.r1) } }];
    if (right) return [{ id: table.id, range: { ...t, c1: Math.max(t.c1, written.c1) } }];
    return [];
  });
  if (growth.length === 0) return;
  recordTablesChange(history, store, () => {
    for (const { id, range } of growth) updateTableOverlay(store, id, { range });
  });
}
