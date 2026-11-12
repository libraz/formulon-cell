import { panelSize, placePanel, viewportSize } from './overlay-position.js';

/** Gap the Mac menu keeps above the bottom of the viewport. */
const MAC_MENU_BOTTOM_PAD = 6;

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
