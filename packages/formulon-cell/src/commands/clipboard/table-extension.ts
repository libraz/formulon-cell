import { addrKey } from '../../engine/address.js';
import type { Addr, Range } from '../../engine/types.js';
import type { WorkbookHandle } from '../../engine/workbook-handle.js';
import { mutators, type SpreadsheetStore } from '../../store/store.js';
import { defaultObjectNames, type TableOverlay, updateTableOverlay } from '../format-as-table.js';
import type { History } from '../history.js';
import { recordTablesChange } from '../slice-history.js';
import { applyPatchesAtomic } from './apply-patches-atomic.js';

/** The range `table` grows to for a write over `written`, or null when it does not grow. */
const grownRange = (table: TableOverlay, written: Range): Range | null => {
  const t = table.range;
  if (t.sheet !== written.sheet) return null;
  if (!table.showTotal && written.r0 === t.r1 + 1 && written.c0 >= t.c0 && written.c0 <= t.c1) {
    return { ...t, r1: Math.max(t.r1, written.r1) };
  }
  if (written.c0 === t.c1 + 1 && written.r0 >= t.r0 && written.r0 <= t.r1) {
    return { ...t, c1: Math.max(t.c1, written.c1) };
  }
  return null;
};

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
    const range = grownRange(table, written);
    return range ? [{ id: table.id, range }] : [];
  });
  if (growth.length === 0) return;
  recordTablesChange(history, store, () => {
    for (const { id, range } of growth) updateTableOverlay(store, id, { range });
  });
}

const isBlank = (wb: WorkbookHandle, addr: Addr): boolean => wb.getValue(addr).kind === 'blank';

/** Header for a column added to the right: the last header's trailing number
 *  continued (`k2` -> `k3`), else the first free default column name. */
function nextHeaderName(wb: WorkbookHandle, t: Range): string {
  const headers: string[] = [];
  for (let col = t.c0; col <= t.c1; col += 1) {
    const value = wb.getValue({ sheet: t.sheet, row: t.r0, col });
    if (value.kind === 'text') headers.push(value.value.trim());
  }
  const used = new Set(headers.map((h) => h.toLocaleLowerCase()));
  const pattern = /^(.*?)(\d+)$/.exec(headers.at(-1) ?? '');
  const prefix = pattern ? (pattern[1] ?? '') : defaultObjectNames(wb).column;
  let n = pattern ? Number(pattern[2]) + 1 : 1;
  while (used.has(`${prefix}${n}`.toLocaleLowerCase())) n += 1;
  return `${prefix}${n}`;
}

/**
 * Grow a table toward one typed entry at `addr`, as one undo step after the
 * entry itself. An entry in the row directly below a table adds that row
 * unless the row already holds a value beside the table; one in the column
 * directly right of it adds that column and names its blank header. Diagonal
 * corners and cells past a gap leave the table alone.
 */
export function extendTablesForEntry(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  history: History | null,
  addr: Addr,
): void {
  const written: Range = {
    sheet: addr.sheet,
    r0: addr.row,
    c0: addr.col,
    r1: addr.row,
    c1: addr.col,
  };
  const table = store.getState().tables.tables.find((t) => grownRange(t, written));
  const range = table && grownRange(table, written);
  if (!table || !range) return;
  const t = table.range;
  const header: Addr | null =
    range.c1 > t.c1 && table.showHeader ? { sheet: t.sheet, row: t.r0, col: range.c1 } : null;
  if (range.r1 > t.r1) {
    for (let col = Math.max(0, t.c0 - 1); col <= t.c1 + 1; col += 1) {
      const cell = { sheet: t.sheet, row: range.r1, col };
      if (addrKey(cell) !== addrKey(addr) && !isBlank(wb, cell)) return;
    }
  }
  const transaction = history?.begin();
  try {
    if (header && isBlank(wb, header)) {
      const name = nextHeaderName(wb, t);
      applyPatchesAtomic(wb, [{ addr: header, value: { kind: 'text', value: name } }], history);
      if (store.getState().data.sheetIndex === t.sheet)
        mutators.replaceCells(store, wb.cells(t.sheet));
    }
    recordTablesChange(history, store, () => {
      updateTableOverlay(store, table.id, { range });
    });
  } finally {
    if (transaction !== undefined) history?.end(transaction);
  }
}
