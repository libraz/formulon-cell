import type { PivotSpill } from '../engine/workbook-handle-pivot.js';
import type { Strings } from '../i18n/strings.js';
import { cellRect, layoutForView } from '../render/geometry.js';
import type { SpreadsheetStore } from '../store/store.js';
import { createInteractionButton } from './chip-button.js';
import { deferInitialFocus } from './dialog-shell.js';
import { overlayPortalFor } from './overlay-portal.js';
import { clampPanelToViewport } from './overlay-position.js';

const open = new WeakMap<HTMLElement, () => void>();

/** Tell the user a report edit left a report at #SPILL!: a callout hanging
 *  from the report's anchor cell, dismissed with OK, a click elsewhere or
 *  Escape. No callout for an anchor on another sheet. */
export function showPivotSpillCallout(
  host: HTMLElement,
  store: SpreadsheetStore,
  strings: Strings,
  spill: PivotSpill,
): void {
  open.get(host)?.();
  const state = store.getState();
  if (state.data.sheetIndex !== spill.anchor.sheet) return;
  const t = strings.workbookObjects;
  const doc = host.ownerDocument;
  const callout = doc.createElement('div');
  callout.className = 'fc-pivot-spill-callout';
  callout.setAttribute('role', 'alertdialog');
  callout.setAttribute('aria-label', t.pivotSpillTitle);
  callout.style.position = 'fixed';
  const title = doc.createElement('div');
  title.className = 'fc-pivot-spill-callout__title';
  title.textContent = t.pivotSpillTitle;
  const body = doc.createElement('div');
  body.className = 'fc-pivot-spill-callout__body';
  body.textContent = t.pivotSpillBody;
  const ok = createInteractionButton({
    className: 'fc-pivot-spill-callout__ok',
    text: t.pivotSpillOk,
  });
  callout.append(title, body, ok);

  const close = (): void => {
    callout.remove();
    doc.removeEventListener('mousedown', onDocMouseDown, true);
    doc.removeEventListener('keydown', onDocKey, true);
    if (open.get(host) === close) open.delete(host);
  };
  const onDocMouseDown = (e: MouseEvent): void => {
    if (!callout.contains(e.target as Node)) close();
  };
  const onDocKey = (e: KeyboardEvent): void => {
    if (e.key === 'Escape') close();
  };
  ok.addEventListener('click', close);

  overlayPortalFor(host).appendChild(callout);
  const cell = cellRect(layoutForView(state), state.viewport, spill.anchor.row, spill.anchor.col);
  const hostRect = host.getBoundingClientRect();
  // The callout points down at the anchor cell from above it.
  const height = callout.getBoundingClientRect().height || 241;
  const position = clampPanelToViewport(
    callout,
    hostRect.left + cell.x + cell.w / 2 - 192,
    hostRect.top + cell.y - height - 8,
    { pad: 4, fallbackWidth: 384, fallbackHeight: 241 },
  );
  callout.style.left = `${position.x}px`;
  callout.style.top = `${position.y}px`;
  deferInitialFocus(callout, () => ok.focus({ preventScroll: true }));
  doc.addEventListener('mousedown', onDocMouseDown, true);
  doc.addEventListener('keydown', onDocKey, true);
  open.set(host, close);
}
