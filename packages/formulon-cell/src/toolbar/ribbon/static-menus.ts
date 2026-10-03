// Open/close handling for ribbon menus that are plain sibling DOM (the
// fallback when no dynamic-dropdown delegation is attached).

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
  return true;
};
