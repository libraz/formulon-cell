import { overlayPortalFor } from './overlay-portal.js';

/** The fill handle's drag tooltip: the value the furthest target cell will
 *  receive, in a small dark box just below-right of the pointer. */
export interface FillValueTooltip {
  show(text: string, clientX: number, clientY: number): void;
  hide(): void;
  dispose(): void;
}

/** Offset from the pointer to the tooltip's top-left corner, in CSS px. */
const POINTER_OFFSET = { x: 10, y: 14 };

export function createFillValueTooltip(host: HTMLElement): FillValueTooltip {
  let tip: HTMLElement | null = null;
  return {
    show(text, clientX, clientY) {
      if (!tip) {
        tip = document.createElement('div');
        tip.className = 'fc-fill-tooltip';
        tip.setAttribute('role', 'tooltip');
        overlayPortalFor(host).appendChild(tip);
      }
      tip.textContent = text;
      tip.style.left = `${clientX + POINTER_OFFSET.x}px`;
      tip.style.top = `${clientY + POINTER_OFFSET.y}px`;
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
