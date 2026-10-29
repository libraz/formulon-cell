const FOCUSABLE_DIALOG_SELECTOR = [
  'button',
  'input',
  'select',
  'textarea',
  'a[href]',
  '[tabindex]:not([tabindex="-1"])',
].join(',');

/**
 * Collect sequential dialog tab stops, accounting for the single-stop rule
 * browsers apply to each named radio group. The checked member is the default
 * representative; a currently focused member wins while focus is inside the
 * group. Unnamed radios remain independent tab stops.
 */
export function focusableDialogItems(root: HTMLElement): HTMLElement[] {
  const ownerWindow = root.ownerDocument.defaultView;
  const eligible = Array.from(root.querySelectorAll<HTMLElement>(FOCUSABLE_DIALOG_SELECTOR)).filter(
    (el) => {
      if (el.closest('[hidden],[aria-hidden="true"]')) return false;
      if (el.matches(':disabled')) return false;
      if (el.tagName === 'INPUT' && (el as HTMLInputElement).type === 'hidden') return false;
      const visibility = ownerWindow?.getComputedStyle(el).visibility;
      if (visibility === 'hidden' || visibility === 'collapse') return false;
      for (let ancestor: HTMLElement | null = el; ancestor; ancestor = ancestor.parentElement) {
        if (ownerWindow?.getComputedStyle(ancestor).display === 'none') return false;
      }
      return el.tabIndex >= 0;
    },
  );

  const groups = new Map<HTMLFormElement | null, Map<string, HTMLInputElement[]>>();
  for (const item of eligible) {
    if (item.tagName !== 'INPUT') continue;
    const radio = item as HTMLInputElement;
    if (radio.type !== 'radio' || radio.name === '') continue;
    let namedGroups = groups.get(radio.form);
    if (!namedGroups) {
      namedGroups = new Map();
      groups.set(radio.form, namedGroups);
    }
    const group = namedGroups.get(radio.name);
    if (group) group.push(radio);
    else namedGroups.set(radio.name, [radio]);
  }

  const activeElement = root.ownerDocument.activeElement;
  const excluded = new Set<HTMLElement>();
  for (const namedGroups of groups.values()) {
    for (const radios of namedGroups.values()) {
      if (radios.length < 2) continue;
      const representative =
        radios.find((radio) => radio === activeElement) ??
        radios.find((radio) => radio.checked) ??
        radios[0];
      for (const radio of radios) {
        if (radio !== representative) excluded.add(radio);
      }
    }
  }

  return eligible.filter((item) => !excluded.has(item));
}

/** Move one sequential Tab stop within the dialog and prevent browser escape. */
export function advanceDialogTab(root: HTMLElement, event: KeyboardEvent): boolean {
  if (
    event.key !== 'Tab' ||
    event.defaultPrevented ||
    event.altKey ||
    event.ctrlKey ||
    event.metaKey
  ) {
    return false;
  }

  const items = focusableDialogItems(root);
  if (items.length === 0) {
    event.preventDefault();
    root.focus({ preventScroll: true });
    return true;
  }

  const activeIndex = items.indexOf(root.ownerDocument.activeElement as HTMLElement);
  const delta = event.shiftKey ? -1 : 1;
  const nextIndex =
    activeIndex === -1
      ? event.shiftKey
        ? items.length - 1
        : 0
      : (activeIndex + delta + items.length) % items.length;
  event.preventDefault();
  items[nextIndex]?.focus({ preventScroll: true });
  return true;
}
