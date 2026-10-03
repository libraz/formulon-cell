import type { ErrorMenuHandle } from '../interact/error-menu.js';
import { getErrorTriangleHits } from '../render/grid.js';

export interface ErrorIndicatorClickDeps {
  canvas: HTMLElement;
  getErrorMenu: () => ErrorMenuHandle | null;
}

// Error / validation triangle clicks. We use `click` (not `pointerdown`)
// so the existing pointer handler gets to set the active cell first —
// that way the menu and the cell select agree on the addr the user
// just clicked.
export function createErrorIndicatorClickHandler(
  deps: ErrorIndicatorClickDeps,
): (e: MouseEvent) => void {
  const { canvas, getErrorMenu } = deps;
  return (e) => {
    const errorMenu = getErrorMenu();
    if (!errorMenu) return;
    if (e.button !== 0) return;
    const rect = canvas.getBoundingClientRect();
    const lx = e.clientX - rect.left;
    const ly = e.clientY - rect.top;
    // Pad by 2px on each side so the 6px corner triangle is comfortable to
    // hit on touch / coarse-pointer devices.
    const pad = 2;
    for (const hit of getErrorTriangleHits()) {
      const r = hit.rect;
      if (lx < r.x - pad || lx > r.x + r.w + pad || ly < r.y - pad || ly > r.y + r.h + pad) {
        continue;
      }
      e.stopPropagation();
      e.preventDefault();
      errorMenu.open(hit.addr, e.clientX, e.clientY, hit.kind);
      return;
    }
  };
}
