type MenuOptions = {
  close: (restoreFocus: boolean) => void;
  restoreFocusTo?: HTMLElement | null;
};

const menuItems = (menu: HTMLElement): HTMLButtonElement[] =>
  Array.from(menu.querySelectorAll<HTMLButtonElement>('button')).filter((item) => !item.disabled);

export function prepareMenu(menu: HTMLElement, label?: string): void {
  menu.setAttribute('role', 'menu');
  if (label) menu.setAttribute('aria-label', label);
  for (const button of menu.querySelectorAll<HTMLButtonElement>('button')) {
    if (!button.hasAttribute('role')) button.setAttribute('role', 'menuitem');
    button.tabIndex = -1;
  }
}

export function focusMenuItem(menu: HTMLElement, index = 0): void {
  const items = menuItems(menu);
  if (items.length === 0) {
    menu.tabIndex = -1;
    menu.focus();
    return;
  }
  const next = Math.max(0, Math.min(index, items.length - 1));
  for (const [idx, item] of items.entries()) item.tabIndex = idx === next ? 0 : -1;
  items[next]?.focus();
}

export function handleMenuKeydown(
  event: KeyboardEvent,
  menu: HTMLElement,
  options: MenuOptions,
): void {
  const items = menuItems(menu);
  const active =
    document.activeElement instanceof HTMLButtonElement ? document.activeElement : null;
  const idx = active ? items.indexOf(active) : -1;
  const move = (next: number): void => {
    event.preventDefault();
    event.stopPropagation();
    focusMenuItem(menu, (next + items.length) % items.length);
  };

  if (event.key === 'Escape') {
    event.preventDefault();
    event.stopPropagation();
    options.close(true);
    options.restoreFocusTo?.focus();
    return;
  }
  if (items.length === 0) return;
  if (event.key === 'ArrowDown') {
    move(idx < 0 ? 0 : idx + 1);
  } else if (event.key === 'ArrowUp') {
    move(idx < 0 ? items.length - 1 : idx - 1);
  } else if (event.key === 'Home') {
    move(0);
  } else if (event.key === 'End') {
    move(items.length - 1);
  } else if (event.key === 'Enter' || event.key === ' ') {
    event.preventDefault();
    event.stopPropagation();
    (idx >= 0 ? items[idx] : items[0])?.click();
  }
}
