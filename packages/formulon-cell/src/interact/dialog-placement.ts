/**
 * Where a dialog opens. On the Mac profile a dialog sits centred on a point and keeps its top-left
 * corner when it later changes size: the point is the clicked context-menu row when the dialog
 * comes from one, the window centre otherwise. Other profiles keep the stylesheet's centring.
 */

export interface DialogAnchor {
  x: number;
  y: number;
}

const PLACED_CLASS = 'fc-fmtdlg--placed';

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

const clamp = (value: number, min: number, max: number): number =>
  Math.min(Math.max(value, min), Math.max(min, max));

/** Pin `panel` so that it is centred on `anchor` (the window centre when absent). */
export const placeDialogPanel = (
  overlay: HTMLElement,
  panel: HTMLElement,
  anchor: DialogAnchor | null = null,
): void => {
  if (!onMac(overlay)) return;
  const view = overlay.ownerDocument.defaultView;
  const width = view?.innerWidth || 1024;
  const height = view?.innerHeight || 768;
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
};

/** Undo `placeDialogPanel` so the next opening is placed afresh. */
export const clearDialogPlacement = (overlay: HTMLElement, panel: HTMLElement): void => {
  overlay.classList.remove(PLACED_CLASS);
  for (const property of ['position', 'margin', 'left', 'top'])
    panel.style.removeProperty(property);
};
