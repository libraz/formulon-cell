type MenuOptions = {
  close: (restoreFocus: boolean) => void;
  restoreFocusTo?: HTMLElement | null;
  /** Selector for the menu's items; every button by default. */
  items?: string;
};

/** Dataset key holding an element's own title while a reason replaces it. */
const BASE_TITLE_KEY = 'fcBaseTitle';

export interface DisabledReasonProjectionOptions {
  ariaDescription?: boolean;
  describedById?: string | null;
  datasetKey?: string;
  title?: boolean;
  titlePrefix?: string;
}

/** Shows why a control is unavailable: as its tooltip, as its accessible description, or both.
 *  A host's search results or menus use it the way the ribbon does; `null` clears it. */
export function projectDisabledReason(
  el: HTMLElement,
  reason: string | null,
  options: DisabledReasonProjectionOptions = {},
): void {
  const useTitle = options.title ?? true;
  const useAriaDescription = options.ariaDescription ?? !options.describedById;
  if (reason) {
    if (useTitle) {
      // Remember the title the element had before its first reason so re-enabling restores it.
      if (!options.titlePrefix && el.dataset[BASE_TITLE_KEY] === undefined)
        el.dataset[BASE_TITLE_KEY] = el.title;
      el.title = options.titlePrefix ? `${options.titlePrefix}\n${reason}` : reason;
    }
    if (options.describedById) el.setAttribute('aria-describedby', options.describedById);
    if (useAriaDescription) el.setAttribute('aria-description', reason);
    if (options.datasetKey) el.dataset[options.datasetKey] = reason;
    return;
  }
  if (useTitle) {
    const base = el.dataset[BASE_TITLE_KEY];
    if (options.titlePrefix !== undefined) el.title = options.titlePrefix;
    else if (base !== undefined) el.title = base;
    delete el.dataset[BASE_TITLE_KEY];
  }
  if (options.describedById !== undefined) el.removeAttribute('aria-describedby');
  if (useAriaDescription) el.removeAttribute('aria-description');
  if (options.datasetKey) delete el.dataset[options.datasetKey];
}

export function projectDisabledState<T extends HTMLElement & { disabled?: boolean }>(
  el: T,
  disabled: boolean,
  reason: string | null,
  options: DisabledReasonProjectionOptions = {},
): void {
  if ('disabled' in el) el.disabled = disabled;
  el.setAttribute('aria-disabled', disabled ? 'true' : 'false');
  projectDisabledReason(el, disabled ? reason : null, options);
}

/** Dataset key holding a value-showing control's own name, before its value is appended. */
const CONTROL_NAME_KEY = 'fcRibbonControlLabel';

/** Names a control after its visible value: `name: value`, so the accessible name contains the
 *  text on screen. The name is the one given, else the one remembered, else the current label; a
 *  repeated call replaces the value instead of appending. Returns the resulting name. */
export function projectValueName(el: HTMLElement, value: string | null, name?: string): string {
  const base =
    name ??
    el.dataset[CONTROL_NAME_KEY] ??
    el.getAttribute('aria-label') ??
    el.getAttribute('title') ??
    '';
  el.dataset[CONTROL_NAME_KEY] = base;
  const shown = value?.trim() ?? '';
  // A name that already ends in a colon ("Width:", "Largeur :") takes the value after a space.
  const joiner = /[:：]\s*$/.test(base) ? ' ' : ': ';
  const label = shown ? (base ? `${base.trimEnd()}${joiner}${shown}` : shown) : base;
  if (label) el.setAttribute('aria-label', label);
  return label;
}

const foldForName = (text: string): string =>
  text
    .normalize('NFKC')
    .toLowerCase()
    .replace(/[^\p{L}\p{N}]+/gu, ' ')
    .trim();

/** Accessible name for a control that shows `label` and has the tooltip `title`: the tooltip
 *  when it contains the visible label, else the label itself, so the name always carries the
 *  visible text. A tooltip that is not the name stays the control's description. */
export function visibleLabelName(label: string, title: string): string {
  return foldForName(title).includes(foldForName(label)) ? title : label;
}

const menuItems = (menu: HTMLElement, selector = 'button'): HTMLButtonElement[] =>
  Array.from(menu.querySelectorAll<HTMLButtonElement>(selector)).filter(
    (item) =>
      !item.disabled &&
      item.getAttribute('aria-disabled') !== 'true' &&
      item.closest<HTMLElement>('[role="menu"]') === menu,
  );

/** Runs `handler` once the pointer moves onto `item`, and again after it leaves and returns. An
 *  item that appears under a resting pointer (a menu opened from the keyboard) fires
 *  `mouseenter` without a move, so hover never takes keyboard focus or opens a submenu then. */
export function onPointerArrive(item: HTMLElement, handler: () => void): void {
  let arrived = false;
  item.addEventListener('mousemove', () => {
    if (arrived) return;
    arrived = true;
    handler();
  });
  item.addEventListener('mouseleave', () => {
    arrived = false;
  });
}

/** The delegated form of `onPointerArrive`: runs `handler` when the pointer moves onto another
 *  element under `root`, never for an element that appeared under a resting pointer. Returns the
 *  detacher. */
export function delegatePointerArrive(
  root: Document | HTMLElement,
  handler: (event: MouseEvent) => void,
): () => void {
  let last: EventTarget | null = null;
  const onMove = (event: Event): void => {
    if (event.target === last) return;
    last = event.target;
    handler(event as MouseEvent);
  };
  root.addEventListener('mousemove', onMove);
  return () => root.removeEventListener('mousemove', onMove);
}

export function prepareMenu(menu: HTMLElement, label?: string): void {
  menu.setAttribute('role', 'menu');
  if (label) menu.setAttribute('aria-label', label);
  for (const button of menu.querySelectorAll<HTMLButtonElement>('button')) {
    if (!button.hasAttribute('role')) button.setAttribute('role', 'menuitem');
    button.tabIndex = -1;
  }
}

export function focusMenuItem(
  menu: HTMLElement,
  index: number | 'first' | 'last' = 0,
  selector?: string,
): void {
  const items = menuItems(menu, selector);
  if (items.length === 0) {
    menu.tabIndex = -1;
    menu.focus({ preventScroll: true });
    return;
  }
  const numericIndex = index === 'first' ? 0 : index === 'last' ? items.length - 1 : index;
  const next = Math.max(0, Math.min(numericIndex, items.length - 1));
  for (const [idx, item] of items.entries()) item.tabIndex = idx === next ? 0 : -1;
  const item = items[next];
  if (!item) return;
  item.focus({ preventScroll: true });
  revealInMenu(menu, item);
}

/** Scrolls only the menu's own overflow box so `item` is in view; the page stays put. */
function revealInMenu(menu: HTMLElement, item: HTMLElement): void {
  if (menu.scrollHeight <= menu.clientHeight) return;
  const top = menu.getBoundingClientRect().top + menu.clientTop;
  const bottom = top + menu.clientHeight;
  const rect = item.getBoundingClientRect();
  if (rect.top < top) menu.scrollTop -= top - rect.top;
  else if (rect.bottom > bottom) menu.scrollTop += rect.bottom - bottom;
}

export function handleMenuKeydown(
  event: KeyboardEvent,
  menu: HTMLElement,
  options: MenuOptions,
): void {
  const items = menuItems(menu, options.items);
  const active =
    document.activeElement instanceof HTMLButtonElement ? document.activeElement : null;
  const idx = active ? items.indexOf(active) : -1;
  const move = (next: number): void => {
    event.preventDefault();
    event.stopPropagation();
    focusMenuItem(menu, (next + items.length) % items.length, options.items);
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
    // As in a native menu, Return with no highlighted item does nothing.
    event.preventDefault();
    event.stopPropagation();
    if (idx >= 0) items[idx]?.click();
  }
}
