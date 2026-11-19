import { isPopulated } from '../commands/current-region.js';
import {
  defaultObjectNames,
  type TableOverlay,
  updateTableOverlay,
} from '../commands/format-as-table.js';
import { transactIn } from '../commands/history.js';
import { recordTablesChange } from '../commands/slice-history.js';
import {
  type TableActionContext,
  tableChangeAllowed,
  tableWrite,
} from '../commands/table-actions.js';
import type { Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { layoutForView, rangeRects } from '../render/geometry.js';
import { tableGripRect } from '../render/painters/handles.js';
import { mutators, type SpreadsheetStore, type State } from '../store/store.js';

/** The command a grip drag runs as. */
export const TABLE_RESIZE_COMMAND = 'tableResize';

/** Pixels the table grip hit zone extends past the painted glyph. */
const TABLE_GRIP_HIT_PAD = 3;

/** The table whose resize grip lies under the pointer on the current sheet, when `store` lets a
 *  grip drag resize it. */
export function resizableTableGripAt(
  store: SpreadsheetStore,
  x: number,
  y: number,
): TableOverlay | null {
  const table = tableGripAt(store.getState(), x, y);
  const ctx = { store, history: null, origin: 'editor', commandId: TABLE_RESIZE_COMMAND } as const;
  return table && tableChangeAllowed(ctx, table.range) ? table : null;
}

/** The table whose resize grip lies under the pointer on the current sheet. */
function tableGripAt(s: State, x: number, y: number): TableOverlay | null {
  const layout = layoutForView(s);
  for (const table of s.tables.tables) {
    if (table.range.sheet !== s.data.sheetIndex) continue;
    const { r1, c1 } = table.range;
    const [corner] = rangeRects(layout, s.viewport, { r0: r1, c0: c1, r1, c1 });
    if (!corner) continue;
    const grip = tableGripRect(corner, layout.rtl);
    if (
      x >= grip.x - TABLE_GRIP_HIT_PAD &&
      x <= grip.x + grip.w + TABLE_GRIP_HIT_PAD &&
      y >= grip.y - TABLE_GRIP_HIT_PAD &&
      y <= grip.y + grip.h + TABLE_GRIP_HIT_PAD
    ) {
      return table;
    }
  }
  return null;
}

const rowBlank = (s: State, sheet: number, row: number, c0: number, c1: number): boolean => {
  for (let col = c0; col <= c1; col += 1) if (isPopulated(s, sheet, row, col)) return false;
  return true;
};

/**
 * The range a grip drag to `cell` resizes `table` to. Columns follow the
 * pointer; rows only grow, and only into rows that are blank across the new
 * width, so a table never swallows data below it.
 */
export function tableResizeRange(
  s: State,
  table: TableOverlay,
  cell: { row: number; col: number },
): Range {
  const { sheet, r0, c0, r1 } = table.range;
  const c1 = Math.max(c0, cell.col);
  let nextR1 = r1;
  if (!table.showTotal) {
    for (let row = r1 + 1; row <= cell.row; row += 1) {
      if (!rowBlank(s, sheet, row, c0, c1)) break;
      nextR1 = row;
    }
  }
  return { sheet, r0, c0, r1: nextR1, c1 };
}

/** `h3` -> `h4`; a name without a trailing number falls back to the default column label. */
const nextHeaderName = (previous: string, fallback: string): string => {
  const match = /^(.*?)(\d+)$/.exec(previous.trim());
  return match ? `${match[1]}${Number(match[2]) + 1}` : fallback;
};

/**
 * Resize `table` to `range` from the entry point `ctx` names, as one authorized step. New columns
 * get header names that continue the last header's numbering; columns that leave the table keep
 * their contents. Returns whether the table changed.
 */
export function applyTableResize(
  ctx: TableActionContext & { readonly workbook: WorkbookHandle },
  table: TableOverlay,
  range: Range,
): boolean {
  const old = table.range;
  if (range.r1 === old.r1 && range.c1 === old.c1) return false;
  return tableWrite(ctx, old, false, () => resizeTable(ctx, table, range));
}

const resizeTable = (
  ctx: TableActionContext & { readonly workbook: WorkbookHandle },
  table: TableOverlay,
  range: Range,
): boolean => {
  const { store, history, workbook: wb } = ctx;
  const old = table.range;
  let ok = false;
  transactIn(
    history && !history.isReplaying() ? history : null,
    () => {
      if (range.c1 > old.c1 && table.showHeader) {
        const { column } = defaultObjectNames(wb);
        let previous = wb.getValue({ sheet: old.sheet, row: old.r0, col: old.c1 });
        for (let col = old.c1 + 1; col <= range.c1; col += 1) {
          const prevText = previous.kind === 'text' ? previous.value : '';
          const name = nextHeaderName(prevText, `${column}${col - old.c0 + 1}`);
          wb.setText({ sheet: old.sheet, row: old.r0, col }, name);
          previous = { kind: 'text', value: name };
        }
      }
      recordTablesChange(history, store, () => {
        ok = updateTableOverlay(store, table.id, { range }) !== null;
      });
      if (ok) mutators.replaceCells(store, wb.cells(old.sheet));
    },
    { keep: () => ok },
  );
  return ok;
};
