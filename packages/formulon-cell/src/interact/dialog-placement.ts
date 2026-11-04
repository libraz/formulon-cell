/**
 * Where a dialog opens. On the Mac profile a dialog sits centred on a point and keeps its top-left
 * corner when it later changes size: the point is the clicked context-menu row when the dialog
 * comes from one, the window centre otherwise. Other profiles keep the stylesheet's centring.
 */

import { clamp, viewportSize } from './overlay-position.js';

export interface DialogAnchor {
  x: number;
  y: number;
}

const PLACED_CLASS = 'fc-fmtdlg--placed';

/** Window-relative limits measured on the desktop app: the title bar stops just below the menu bar,
 *  the dialog stops short of the bottom edge and 2pt short of the right edge. */
const DRAG_TOP_MIN = 3;
const DRAG_BOTTOM_MARGIN = 104;
const DRAG_RIGHT_MARGIN = 2;

const draggableHeaders = new WeakSet<HTMLElement>();

let active: DialogAnchor | null = null;

/** Run `open` with `anchor` as the dialogs' centre; the anchor never outlives the call. */
export const withDialogAnchor = <T>(anchor: DialogAnchor | null, open: () => T): T => {
  const previous = active;
  active = anchor;
  try {
    return open();
  } finally {
    active = previous;
  }
};

/** The anchor of the dispatch in progress, for code that must pass it across an await. */
export const currentDialogAnchor = (): DialogAnchor | null => active;

const onMac = (overlay: HTMLElement): boolean =>
  overlay.closest<HTMLElement>('[data-fc-platform]')?.dataset.fcPlatform === 'mac';

/** Let the dialog title bar move the pinned panel by the pointer delta. */
const enableTitleDrag = (overlay: HTMLElement, panel: HTMLElement): void => {
  const header = panel.querySelector<HTMLElement>(':scope > .fc-fmtdlg__header');
  if (!header || draggableHeaders.has(header)) return;
  draggableHeaders.add(header);
  header.addEventListener('pointerdown', (event) => {
    if (event.button !== 0 || !overlay.classList.contains(PLACED_CLASS)) return;
    if ((event.target as HTMLElement).closest('button, input, select, textarea, a')) return;
    const doc = overlay.ownerDocument;
    const startX = event.clientX;
    const startY = event.clientY;
    const startLeft = Number.parseFloat(panel.style.left) || 0;
    const startTop = Number.parseFloat(panel.style.top) || 0;
    const move = (e: MouseEvent): void => {
      const { width, height } = viewportSize(doc);
      const rect = panel.getBoundingClientRect();
      panel.style.left = `${clamp(startLeft + e.clientX - startX, 0, width - rect.width - DRAG_RIGHT_MARGIN)}px`;
      panel.style.top = `${clamp(startTop + e.clientY - startY, DRAG_TOP_MIN, height - rect.height - DRAG_BOTTOM_MARGIN)}px`;
    };
    const stop = (): void => {
      doc.removeEventListener('pointermove', move);
      doc.removeEventListener('pointerup', stop);
      doc.removeEventListener('pointercancel', stop);
    };
    doc.addEventListener('pointermove', move);
    doc.addEventListener('pointerup', stop);
    doc.addEventListener('pointercancel', stop);
    event.preventDefault();
  });
};

/** Pin `panel` so that it is centred on `anchor` (the window centre when absent). */
export const placeDialogPanel = (
  overlay: HTMLElement,
  panel: HTMLElement,
  anchor: DialogAnchor | null = null,
): void => {
  if (!onMac(overlay)) return;
  const { width, height } = viewportSize(overlay.ownerDocument);
  const rect = panel.getBoundingClientRect();
  if (rect.width === 0 && rect.height === 0) return;
  const left = clamp(Math.round((anchor?.x ?? width / 2) - rect.width / 2), 0, width - rect.width);
  const top = clamp(
    Math.round((anchor?.y ?? height / 2) - rect.height / 2),
    0,
    height - rect.height,
  );
  overlay.classList.add(PLACED_CLASS);
  panel.style.position = 'absolute';
  panel.style.margin = '0';
  panel.style.left = `${left}px`;
  panel.style.top = `${top}px`;
  enableTitleDrag(overlay, panel);
};

/** Undo `placeDialogPanel` so the next opening is placed afresh. */
export const clearDialogPlacement = (overlay: HTMLElement, panel: HTMLElement): void => {
  overlay.classList.remove(PLACED_CLASS);
  for (const property of ['position', 'margin', 'left', 'top'])
    panel.style.removeProperty(property);
};
