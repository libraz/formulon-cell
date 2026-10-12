// Open/close handling for ribbon menus that are plain sibling DOM (the
// fallback when no dynamic-dropdown delegation is attached).

import { clampPanelToViewport, panelSize, viewportSize } from '../../interact/overlay-position.js';

const STATIC_MENU_GAP = 3;
const STATIC_MENU_PAD = 4;
const STATIC_MENU_FALLBACK_WIDTH = 216;
const STATIC_MENU_FALLBACK_HEIGHT = 260;

const clearStaticMenuPosition = (menu: HTMLElement): void => {
  menu.style.removeProperty('position');
  menu.style.removeProperty('left');
  menu.style.removeProperty('top');
  menu.style.removeProperty('right');
  menu.style.removeProperty('bottom');
  menu.style.removeProperty('width');
  menu.style.removeProperty('min-width');
  menu.style.removeProperty('max-width');
  menu.style.removeProperty('max-height');
  menu.style.removeProperty('overflow-y');
  menu.style.removeProperty('overscroll-behavior');
  menu.style.removeProperty('box-sizing');
};

/** Places a static sibling in the viewport coordinate space. The ribbon panel
 * is a horizontal scroll owner, so an absolutely positioned sibling would be
 * clipped by the panel's overflow. A fixed menu retains the trigger's visual
 * position while `clampPanelToViewport` also handles transformed hosts. */
const positionStaticRibbonMenu = (button: HTMLElement, menu: HTMLElement): void => {
  clearStaticMenuPosition(menu);
  menu.style.position = 'fixed';
  menu.style.left = '-10000px';
  menu.style.top = '-10000px';
  menu.style.right = 'auto';
  menu.style.bottom = 'auto';
  menu.style.boxSizing = 'border-box';

  const viewport = viewportSize(menu.ownerDocument);
  const availableWidth = Math.max(0, viewport.width - STATIC_MENU_PAD * 2);
  const availableHeight = Math.max(0, viewport.height - STATIC_MENU_PAD * 2);
  const natural = panelSize(menu, STATIC_MENU_FALLBACK_WIDTH, STATIC_MENU_FALLBACK_HEIGHT);
  if (natural.width > availableWidth) {
    menu.style.minWidth = '0';
    menu.style.width = `${availableWidth}px`;
    menu.style.maxWidth = `${availableWidth}px`;
  }
  if (natural.height > availableHeight) {
    menu.style.maxHeight = `${availableHeight}px`;
    menu.style.overflowY = 'auto';
    menu.style.overscrollBehavior = 'contain';
  }

  const buttonRect = button.getBoundingClientRect();
  const size = panelSize(menu, STATIC_MENU_FALLBACK_WIDTH, STATIC_MENU_FALLBACK_HEIGHT);
  const belowSpace = Math.max(
    0,
    viewport.height - STATIC_MENU_PAD - buttonRect.bottom - STATIC_MENU_GAP,
  );
  const aboveSpace = Math.max(0, buttonRect.top - STATIC_MENU_PAD - STATIC_MENU_GAP);
  const opensBelow = size.height <= belowSpace || belowSpace >= aboveSpace;
  const desiredX = buttonRect.left;
  const desiredY = opensBelow
    ? buttonRect.bottom + STATIC_MENU_GAP
    : buttonRect.top - size.height - STATIC_MENU_GAP;
  const position = clampPanelToViewport(menu, desiredX, desiredY, {
    pad: STATIC_MENU_PAD,
    fallbackWidth: STATIC_MENU_FALLBACK_WIDTH,
    fallbackHeight: STATIC_MENU_FALLBACK_HEIGHT,
  });
  menu.style.left = `${Math.round(position.x)}px`;
  menu.style.top = `${Math.round(position.y)}px`;
};

/** Hides every open static menu except `except`, collapses the function
 *  category panels, and optionally returns focus to the first closed trigger. */
export const closeStaticRibbonMenus = (
  host: HTMLElement,
  except?: HTMLElement,
  restoreFocus = false,
): void => {
  let restoreTarget: HTMLButtonElement | null = null;
  for (const menu of host.querySelectorAll<HTMLDivElement>('.fc-tb__menu')) {
    if (menu === except || menu.hidden) continue;
    menu.hidden = true;
    clearStaticMenuPosition(menu);
    const button = host.querySelector<HTMLButtonElement>(`[data-ribbon-menu-id="${menu.id}"]`);
    button?.setAttribute('aria-expanded', 'false');
    restoreTarget ??= button;
  }
  for (const panel of host.querySelectorAll<HTMLElement>('[data-function-category-panel]')) {
    panel.hidden = true;
  }
  for (const trigger of host.querySelectorAll<HTMLElement>('[data-function-category-submenu]')) {
    trigger.classList.remove('fc-tb__menu-item--active');
    trigger.setAttribute('aria-expanded', 'false');
  }
  if (restoreFocus) restoreTarget?.focus();
};

/** True when any static ribbon menu under `host` is visible. */
export const hasOpenStaticRibbonMenu = (host: HTMLElement): boolean =>
  Array.from(host.querySelectorAll<HTMLDivElement>('.fc-tb__menu')).some((menu) => !menu.hidden);

/** Toggles the menu attached as `button`'s next sibling, closing the others.
 *  Returns false when the button has no such menu. */
export const toggleStaticRibbonSubmenu = (host: HTMLElement, button: HTMLElement): boolean => {
  const submenu = button.nextElementSibling;
  if (!(submenu instanceof HTMLDivElement) || !submenu.classList.contains('fc-tb__menu'))
    return false;
  const wasOpen = !submenu.hidden;
  closeStaticRibbonMenus(host, submenu);
  submenu.hidden = wasOpen;
  button.setAttribute('aria-expanded', wasOpen ? 'false' : 'true');
  if (wasOpen) {
    clearStaticMenuPosition(submenu);
  } else {
    positionStaticRibbonMenu(button, submenu);
  }
  return true;
};
