import type { PivotSpill } from '../engine/workbook-handle-pivot.js';
import type { Strings } from '../i18n/strings.js';
import { cellRect, layoutForView } from '../render/geometry.js';
import type { SpreadsheetStore } from '../store/store.js';
import { createInteractionButton } from './chip-button.js';
import { onOutsidePress } from './outside-press.js';
import { overlayPortalFor } from './overlay-portal.js';
import { placePanel } from './overlay-position.js';

const open = new WeakMap<HTMLElement, () => void>();

/** Callout size used to place it before it has been measured. */
const CALLOUT_W = 384;
const CALLOUT_H = 241;

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
  // The desktop's help link; the host has no help page to open, so it is shown inert.
  const more = doc.createElement('span');
  more.className = 'fc-pivot-spill-callout__link';
  more.textContent = t.pivotSpillLearnMore;
  callout.append(title, body, more, ok);

  const close = (): void => {
    callout.remove();
    detachOutside();
    doc.removeEventListener('keydown', onDocKey, true);
    if (open.get(host) === close) open.delete(host);
  };
  const onDocKey = (e: KeyboardEvent): void => {
    if (e.key === 'Escape') close();
  };
  ok.addEventListener('click', close);

  overlayPortalFor(host).appendChild(callout);
  const cell = cellRect(layoutForView(state), state.viewport, spill.anchor.row, spill.anchor.col);
  // Cell rectangles are relative to the grid, not to the whole host.
  const grid = host.querySelector<HTMLElement>('.fc-host__grid') ?? host;
  const hostRect = grid.getBoundingClientRect();
  const height = callout.getBoundingClientRect().height || CALLOUT_H;
  // Above the cell when there is room, else below it, so the cell stays in view.
  const above = hostRect.top + cell.y - height - 8;
  placePanel(
    callout,
    hostRect.left + cell.x + cell.w / 2 - CALLOUT_W / 2,
    above >= 4 ? above : hostRect.top + cell.y + cell.h + 8,
    { fallbackWidth: CALLOUT_W, fallbackHeight: CALLOUT_H },
  );
  ok.focus({ preventScroll: true });
  const detachOutside = onOutsidePress(
    doc,
    () => [callout],
    () => close(),
    { event: 'mousedown' },
  );
  doc.addEventListener('keydown', onDocKey, true);
  open.set(host, close);
}
