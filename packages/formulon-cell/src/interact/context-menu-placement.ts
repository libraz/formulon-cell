import { mergeAt } from '../commands/merge.js';
import { bodyBandOrigin, layoutForView, rangeRects } from '../render/geometry.js';
import type { State } from '../store/types.js';
import { panelSize, placePanel, VIEWPORT_PANEL_PAD, viewportSize } from './overlay-position.js';

/** Gap the Mac menu keeps above the bottom of the viewport. */
const MAC_MENU_BOTTOM_PAD = 6;
/** Distance from a Mac submenu's top edge to the centre of its first row. */
const MAC_SUB_ROW_CENTER = 18;
/** How far a non-Mac submenu overlaps its parent row horizontally and sits above it. */
const SUB_OVERLAP_PX = 2;
const SUB_TOP_OFFSET_PX = 4;

/**
 * Places a context menu where the Mac desktop opens one for a pointer at (x, y):
 * one pixel right of and above the pointer, flipped to the pointer's left at the
 * right edge and slid up (not flipped) at the bottom edge.
 */
export function placeMacMenu(panel: HTMLElement, x: number, y: number): void {
  const { width, height } = panelSize(panel);
  const viewport = viewportSize(panel.ownerDocument);
  const left = x + 1 + width > viewport.width ? x - width : x + 1;
  const top = Math.min(y - 1, viewport.height - height - MAC_MENU_BOTTOM_PAD);
  placePanel(panel, left, top);
}

/**
 * Places a submenu beside its parent row: non-Mac overlaps the parent by 2 px and sits
 * 4 px above it; Mac is flush and centres its first row; flips left at the right viewport edge.
 */
export function placeSubmenu(panel: HTMLElement, parentRect: DOMRect, mac: boolean): void {
  const { width: sw } = panelSize(panel);
  const viewport = viewportSize(panel.ownerDocument);
  // The desktop submenu overlaps its parent by the panel padding and centres its
  // first row on the parent row.
  let x = mac ? parentRect.right : parentRect.right - SUB_OVERLAP_PX;
  if (x + sw > viewport.width - VIEWPORT_PANEL_PAD) {
    x = mac ? parentRect.left - sw : parentRect.left - sw + SUB_OVERLAP_PX;
  }
  placePanel(
    panel,
    x,
    mac
      ? parentRect.top + parentRect.height / 2 - MAC_SUB_ROW_CENTER
      : parentRect.top - SUB_TOP_OFFSET_PX,
  );
}

/** Position keyboard-opened menus from the active cell's visible geometry.
 *  The cell layout is local to the grid; the grid rect converts it to client
 *  coordinates before the shared overlay positioning handles clamping and
 *  any transformed portal. */
export function keyboardMenuPoint(state: State, grid: HTMLElement): { x: number; y: number } {
  const layout = layoutForView(state);
  const viewport = state.viewport;
  const gridRect = grid.getBoundingClientRect();
  const fallback = (): { x: number; y: number } => {
    const origin = bodyBandOrigin(layout, viewport);
    return { x: gridRect.left + origin.x, y: gridRect.top + origin.y };
  };

  if (
    !Number.isFinite(gridRect.left) ||
    !Number.isFinite(gridRect.top) ||
    gridRect.width <= 0 ||
    gridRect.height <= 0
  ) {
    return fallback();
  }

  const active = state.selection.active;
  if (active.sheet !== state.data.sheetIndex) return fallback();

  const activeRange = {
    r0: active.row,
    r1: active.row,
    c0: active.col,
    c1: active.col,
  };
  const merge = mergeAt(state, active);
  const fragments = rangeRects(layout, viewport, merge ?? activeRange).filter(
    (rect) => rect.w > 0 && rect.h > 0,
  );
  if (fragments.length === 0) return fallback();

  // `rangeRects` clips to the viewport, so this stays bounded for a far-away row.
  const activeCell = merge
    ? rangeRects(layout, viewport, activeRange).find((rect) => rect.w > 0 && rect.h > 0)
    : fragments[0];

  const fragment = activeCell
    ? fragments.find(
        (rect) =>
          activeCell.x + activeCell.w / 2 >= rect.x &&
          activeCell.x + activeCell.w / 2 <= rect.x + rect.w &&
          activeCell.y + activeCell.h / 2 >= rect.y &&
          activeCell.y + activeCell.h / 2 <= rect.y + rect.h,
      )
    : undefined;
  const visibleFragment = fragment ?? fragments[0];
  if (!visibleFragment) return fallback();

  return {
    x: gridRect.left + visibleFragment.x + (layout.rtl ? visibleFragment.w : 0),
    y: gridRect.top + visibleFragment.y + visibleFragment.h,
  };
}
