import { coerceInput, inputLocaleForWorkbook, workbookDate1904 } from '../commands/coerce-input.js';
import { effectiveCellFormatWithPending } from '../engine/cell-format-projection.js';
import type { Addr } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { bodyBandOrigin, cellRectUnclamped, layoutForView, type Rect } from '../render/geometry.js';
import {
  cellFontFamily,
  cellTextVerticalOffset,
  cssCellFontSize,
  fontSizeWithVerticalAlign,
} from '../render/painters/text.js';
import { frozenColEnd, frozenRowEnd } from '../store/freeze-band.js';
import { renderScale } from '../store/render-scale.js';
import type { SpreadsheetStore } from '../store/store.js';

/** In-cell editor DOM surface: placement, sizing, alignment and cell-format painting. */

/** The editor dependencies the surface helpers read. */
export interface EditorSurfaceDeps {
  store: SpreadsheetStore;
  wb: WorkbookHandle;
  grid: Pick<HTMLElement, 'clientWidth'>;
}

/** The editor's frame for a cell: its 2px border lies exactly where the
 *  selection outline is painted, over the cell's gridlines and one pixel
 *  outside them, so editing keeps the same frame. */
const editorFrame = (cell: Rect): Rect => {
  const x = Math.round(cell.x);
  const y = Math.round(cell.y);
  return {
    x: x - 1,
    y: y - 1,
    w: Math.round(cell.x + cell.w) - x + 3,
    h: Math.round(cell.y + cell.h) - y + 3,
  };
};

export const fitEditorToContent = (
  input: HTMLTextAreaElement | null,
  addr: Addr | null,
  s: EditorSurfaceDeps,
): void => {
  if (!input) return;
  const lines = Math.max(1, (input.value.match(/\n/g)?.length ?? 0) + 1);
  if (lines === 1) {
    // Hide the per-line growth on a fresh single-line edit so the editor
    //  visually matches the cell rect exactly.
    input.style.minHeight = '';
  } else {
    // Spreadsheets grow the editor downward; mirror that with a min-height bump.
    const { layout, viewport } = s.store.getState();
    input.style.minHeight = `${layout.defaultRowHeight * renderScale(layout, viewport) * lines}px`;
  }
  refreshWidth(input, addr, s);
};

/** Grow the editor rightward to fit content wider than the cell — desktop
 *  spreadsheet behavior, so a long formula shows from its leading `=`
 *  instead of scrolling to its tail. Capped at the grid's right edge so the
 *  editor never escapes the viewport. */
const refreshWidth = (
  input: HTMLTextAreaElement | null,
  addr: Addr | null,
  s: EditorSurfaceDeps,
): void => {
  if (!input || !addr) return;
  const state = s.store.getState();
  const layout = layoutForView(state);
  const r = editorFrame(cellRectUnclamped(layout, state.viewport, addr.row, addr.col));
  // Reset to the cell width first so the editor can shrink back as content
  //  is deleted, then measure the natural content width.
  input.style.width = `${r.w}px`;
  const content = input.scrollWidth;
  if (content <= r.w) {
    input.classList.remove('fc-host__editor--overflow');
    input.classList.remove('fc-host__editor--overflow-rtl');
    return;
  }
  // The editor grows away from the cell's leading edge, which is the right
  // edge on a right-to-left sheet — so there it grows leftwards and its
  // left offset moves with the width.
  const room = layout.rtl ? r.x + r.w - 2 : s.grid.clientWidth - r.x - 2;
  const maxWidth = Math.max(r.w, room);
  const want = content + 2;
  const width = Math.min(want, maxWidth);
  input.style.width = `${width}px`;
  if (layout.rtl) input.style.left = `${r.x + r.w - width}px`;
  // When the content is still wider than the editor can grow, the text
  //  scrolls; flag the overflow so the editor drops its trailing border.
  input.classList.toggle('fc-host__editor--overflow', want > maxWidth && !layout.rtl);
  input.classList.toggle('fc-host__editor--overflow-rtl', want > maxWidth && layout.rtl);
};

export const alignEditorText = (
  input: HTMLTextAreaElement | null,
  addr: Addr | null,
  s: EditorSurfaceDeps,
  raw: string,
): void => {
  if (!input || !addr) return;
  // Explicit cell alignment wins over the type-based defaults below — once
  // a user sets center/right/justify on a cell, the editor honors it even
  // for formula edits so the visual position stays anchored.
  const fmt = effectiveCellFormatWithPending(s.wb, s.store.getState(), addr);
  if (fmt?.align) {
    input.style.textAlign = fmt.align;
    return;
  }
  // Formula edits default to left so the leading `=` stays in view rather
  // than scrolling off to the right under content overflow.
  if (raw.startsWith('=')) {
    input.style.textAlign = 'left';
    return;
  }
  const coerced = coerceInput(raw, {
    locale: inputLocaleForWorkbook(s.wb),
    date1904: workbookDate1904(s.wb),
  });
  if (coerced.kind === 'number') input.style.textAlign = 'right';
  else if (coerced.kind === 'bool') input.style.textAlign = 'center';
  else input.style.textAlign = 'left';
};

/** Give the editor the cell's own fill and text color. Without this the
 *  editor reads as a panel floating above the sheet whenever the cell
 *  carries a fill; with it, the edit happens visually inside the cell. */
export const paintEditorCellFormat = (
  input: HTMLTextAreaElement | null,
  addr: Addr | null,
  s: EditorSurfaceDeps,
): void => {
  if (!input || !addr) return;
  const state = s.store.getState();
  const fmt = effectiveCellFormatWithPending(s.wb, state, addr);
  const zoom = renderScale(state.layout, state.viewport);
  // Cell colors stay explicit when authored; a missing family falls through
  // to the host UI stack so an unavailable workbook font cannot switch the
  // editor to the browser's default serif face.
  input.style.background = fmt?.fill ?? '';
  input.style.color = fmt?.color ?? '';
  input.style.fontFamily = cellFontFamily(
    fmt?.fontFamily,
    'var(--fc-font-ui, system-ui, sans-serif)',
  );
  input.style.fontSize = cssCellFontSize(fmt?.fontSize, fmt?.fontVertAlign, zoom);
  const visualSize =
    fmt?.fontSize !== undefined && Number.isFinite(fmt.fontSize)
      ? fontSizeWithVerticalAlign(fmt.fontSize, fmt.fontVertAlign) * zoom
      : undefined;
  input.style.lineHeight = visualSize ? `${visualSize * 1.25}px` : '';
  if (visualSize && fmt?.fontVertAlign) {
    // A textarea ignores `vertical-align` for its text run. Keep the input
    // box welded to the cell and move only the run by balancing its top and
    // bottom padding around the same baseline offset as the canvas painter.
    const offset = cellTextVerticalOffset(visualSize, fmt.fontVertAlign);
    input.style.paddingTop = `${Math.max(0, 4 + offset)}px`;
    input.style.paddingBottom = `${Math.max(0, 4 - offset)}px`;
  } else {
    input.style.paddingTop = '';
    input.style.paddingBottom = '';
  }
  input.style.fontWeight = fmt?.bold ? 'bold' : '';
  input.style.fontStyle = fmt?.italic ? 'italic' : '';
  input.style.textDecoration = [
    fmt?.underline ? 'underline' : '',
    fmt?.strike ? 'line-through' : '',
  ]
    .filter(Boolean)
    .join(' ');
};

export const positionEditorOverCell = (
  input: HTMLTextAreaElement | null,
  addr: Addr | null,
  s: EditorSurfaceDeps,
): void => {
  if (!input || !addr) return;
  const state = s.store.getState();
  const layout = layoutForView(state);
  const cell = cellRectUnclamped(layout, state.viewport, addr.row, addr.col);
  const r = editorFrame(cell);
  if (!input) return;
  // A cell scrolled off the leading edge sits behind the frozen band and the
  // headers. Park the editor outside the grid's clip rect instead of letting
  // it hang over that chrome — it keeps DOM focus (and the caret, and any IME
  // composition) while its cell is out of view, the way a desktop
  // spreadsheet does.
  const band = bodyBandOrigin(layout, state.viewport);
  const scrolledOut =
    (addr.row >= frozenRowEnd(state.layout) && cell.y < band.y) ||
    (addr.col >= frozenColEnd(state.layout) &&
      (layout.rtl ? cell.x + cell.w > band.x : cell.x < band.x));
  input.style.left = scrolledOut ? `${-r.w - 4}px` : `${r.x}px`;
  input.style.top = scrolledOut ? `${-r.h - 4}px` : `${r.y}px`;
  input.style.width = `${r.w}px`;
  input.style.height = `${r.h}px`;
};
