/**
 * @file Row Height and Column Width dialogs. The dialogs speak the desktop
 * spreadsheet's units: row heights in points and column widths in characters
 * of the Normal font, both snapped to whole screen pixels of the workbook's
 * platform. The field opens with the selection's shared size, or empty when
 * the selected tracks differ; zero hides the tracks.
 */

import type { Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { SpreadsheetStore } from '../store/store.js';
import type { History } from './history.js';
import { hideCols, hideRows, setColsWidth, setRowsHeight } from './row-col-layout.js';

/** Largest row height the dialog accepts, in points. */
export const ROW_HEIGHT_MAX_PT = 409;

/** Widest column the dialog accepts, in characters. */
export const COLUMN_WIDTH_MAX_CHARS = 255;

/** Shows the dialog with `initial` (undefined leaves the field empty) and
 *  resolves to the typed value, or null on cancel. */
export type DimensionPrompt = (initial: number | undefined, max: number) => Promise<number | null>;

export interface DimensionDialogContext {
  store: SpreadsheetStore;
  history: History | null;
  wb: WorkbookHandle;
  range: Range;
  prompt: DimensionPrompt;
}

/** The size every track in `[first, last]` shares, or undefined when they differ. */
const sharedSize = (
  sizes: ReadonlyMap<number, number>,
  fallback: number,
  first: number,
  last: number,
): number | undefined => {
  let shared: number | undefined;
  let overridden = 0;
  for (const [index, size] of sizes) {
    if (index < first || index > last) continue;
    if (shared !== undefined && size !== shared) return undefined;
    shared = size;
    overridden += 1;
  }
  if (overridden === last - first + 1) return shared;
  return shared === undefined || shared === fallback ? fallback : undefined;
};

/** Prompt for the selection's row height in points and apply it. */
export async function runRowHeightDialog(ctx: DimensionDialogContext): Promise<void> {
  const { store, history, wb, range } = ctx;
  const layout = store.getState().layout;
  const px = sharedSize(layout.rowHeights, layout.defaultRowHeight, range.r0, range.r1);
  const pt = await ctx.prompt(
    px === undefined ? undefined : wb.rowDisplayPt(px),
    ROW_HEIGHT_MAX_PT,
  );
  if (pt === null) return;
  if (pt === 0) {
    hideRows(store, history, range.r0, range.r1, wb);
    return;
  }
  setRowsHeight(store, history, range.r0, range.r1, wb.rowPxForDisplayPt(pt), wb);
}

/** Prompt for the selection's column width in characters and apply it. */
export async function runColumnWidthDialog(ctx: DimensionDialogContext): Promise<void> {
  const { store, history, wb, range } = ctx;
  const layout = store.getState().layout;
  const px = sharedSize(layout.colWidths, layout.defaultColWidth, range.c0, range.c1);
  const chars = await ctx.prompt(
    px === undefined ? undefined : wb.columnDisplayChars(range.sheet, px),
    COLUMN_WIDTH_MAX_CHARS,
  );
  if (chars === null) return;
  if (chars === 0) {
    hideCols(store, history, range.c0, range.c1, wb);
    return;
  }
  setColsWidth(
    store,
    history,
    range.c0,
    range.c1,
    wb.columnPxForDisplayChars(range.sheet, chars),
    wb,
  );
}
