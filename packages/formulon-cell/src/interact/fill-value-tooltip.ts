import { overlayPortalFor } from './overlay-portal.js';

/** The fill handle's drag tooltip: the value the furthest target cell will
 *  receive, in a small dark box just right of the drag range, level with the pointer. */
export interface FillValueTooltip {
  /** `edgeX` is the drag range's right edge, `pointerY` the pointer row; both in client px. */
  show(text: string, edgeX: number, pointerY: number): void;
  hide(): void;
  dispose(): void;
}

/** Offset from the range's right edge and the pointer row to the box's top-left, in CSS px. */
const ANCHOR_OFFSET = { x: 5, y: -2.5 };

export function createFillValueTooltip(host: HTMLElement): FillValueTooltip {
  let tip: HTMLElement | null = null;
  return {
    show(text, edgeX, pointerY) {
      if (!tip) {
        tip = document.createElement('div');
        tip.className = 'fc-fill-tooltip';
        tip.setAttribute('role', 'tooltip');
        overlayPortalFor(host).appendChild(tip);
      }
      tip.textContent = text;
      tip.style.left = `${edgeX + ANCHOR_OFFSET.x}px`;
      tip.style.top = `${pointerY + ANCHOR_OFFSET.y}px`;
      tip.hidden = false;
    },
    hide() {
      if (tip) tip.hidden = true;
    },
    dispose() {
      tip?.remove();
      tip = null;
    },
  };
}
