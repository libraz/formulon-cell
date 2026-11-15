// jsdom has no sequential focus navigation, so a Tab the page leaves alone is
// carried out here the way a browser does: to the next element in tab order.

/** Elements in the document's tab order, in document order. */
export const tabStops = (): HTMLElement[] =>
  Array.from(
    document.querySelectorAll<HTMLElement>('a[href], button, input, select, textarea, [tabindex]'),
  ).filter(
    (el) => el.tabIndex >= 0 && !el.closest('[hidden]') && !(el as HTMLButtonElement).disabled,
  );

/** Press Tab (or Shift+Tab) on the focused element; false when the page swallowed it. */
export const pressTab = (back: boolean): boolean => {
  const from = document.activeElement as HTMLElement;
  const event = new KeyboardEvent('keydown', {
    key: 'Tab',
    shiftKey: back,
    bubbles: true,
    cancelable: true,
  });
  from.dispatchEvent(event);
  if (event.defaultPrevented) return false;
  // The browser moves on from whatever holds focus once the key's handlers have run.
  const origin = document.activeElement as HTMLElement;
  const stops = tabStops();
  stops[stops.indexOf(origin) + (back ? -1 : 1)]?.focus();
  return true;
};

/**
 * Walk Tab through `region` from its first stop (Shift+Tab from its last) until focus leaves it.
 * Returns the presses taken, or null when a press is swallowed or `limit` presses stay inside.
 */
export const tabOutOf = (region: HTMLElement, back: boolean, limit: number): number | null => {
  const inside = tabStops().filter((el) => region.contains(el));
  const start = back ? inside.at(-1) : inside[0];
  if (!start) throw new Error('The region has no tab stop.');
  start.focus();
  for (let presses = 1; presses <= limit; presses++) {
    if (!pressTab(back)) return null;
    if (!region.contains(document.activeElement)) return presses;
  }
  return null;
};
