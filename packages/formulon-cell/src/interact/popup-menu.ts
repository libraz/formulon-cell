import { focusMenuItem, handleMenuKeydown, prepareMenu } from '../toolbar/menu-a11y.js';
import { createInteractionButton } from './chip-button.js';
import { onOutsidePress } from './outside-press.js';
import { overlayPortalFor } from './overlay-portal.js';

export interface PopupMenuItem {
  label: string;
  /** Set for a choice among several: the item shows a check while true. */
  checked?: boolean;
  onPick: () => void;
}

export interface PopupMenuOptions {
  anchor: HTMLElement;
  items: readonly PopupMenuItem[];
  label?: string;
}

/** A small menu dropped under `anchor` in the overlay portal. Arrow keys move, Return picks,
 *  Escape or a press outside closes it and hands focus back to the anchor. Returns `close`. */
export const openPopupMenu = (opts: PopupMenuOptions): (() => void) => {
  const { anchor } = opts;
  const menu = document.createElement('div');
  menu.className = 'fc-popup-menu';
  let closed = false;
  const close = (): void => {
    if (closed) return;
    closed = true;
    detachOutside();
    menu.remove();
    anchor.setAttribute('aria-expanded', 'false');
  };
  for (const entry of opts.items) {
    const item = createInteractionButton({ className: 'fc-popup-menu__item', text: entry.label });
    if (entry.checked !== undefined) {
      item.setAttribute('role', 'menuitemradio');
      item.setAttribute('aria-checked', entry.checked ? 'true' : 'false');
    }
    item.addEventListener('click', () => {
      close();
      entry.onPick();
    });
    menu.append(item);
  }
  prepareMenu(menu, opts.label);
  menu.addEventListener('keydown', (e) => {
    handleMenuKeydown(e, menu, { close, restoreFocusTo: anchor });
  });
  overlayPortalFor(anchor).append(menu);
  const r = anchor.getBoundingClientRect();
  menu.style.left = `${Math.max(4, r.left)}px`;
  menu.style.top = `${r.bottom + 2}px`;
  anchor.setAttribute('aria-expanded', 'true');
  const detachOutside = onOutsidePress(
    document,
    () => [menu, anchor],
    () => close(),
    {
      event: 'pointerdown',
    },
  );
  focusMenuItem(menu, 'first');
  return close;
};
