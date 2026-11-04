import { overlayPortalFor } from './overlay-portal.js';

/** The size readout shown beside the pointer while a row / column border is
 *  dragged: a small dark box whose bottom edge floats above the pointer. */
export interface ResizeSizeTooltip {
  show(text: string, clientX: number, clientY: number): void;
  hide(): void;
  dispose(): void;
}

/** Gap between the pointer and the tooltip's bottom edge, in CSS px. */
const POINTER_GAP = 14;

export function createResizeSizeTooltip(host: HTMLElement): ResizeSizeTooltip {
  let tip: HTMLElement | null = null;
  return {
    show(text, clientX, clientY) {
      if (!tip) {
        tip = document.createElement('div');
        tip.className = 'fc-resize-tooltip';
        tip.setAttribute('role', 'tooltip');
        overlayPortalFor(host).appendChild(tip);
      }
      tip.textContent = text;
      tip.style.left = `${clientX}px`;
      tip.style.top = `${clientY - POINTER_GAP}px`;
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
