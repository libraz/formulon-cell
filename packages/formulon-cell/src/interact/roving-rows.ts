/**
 * Rows of a list or table that take one tab stop between them, as a listbox does.
 *
 * `bindRovingRows` is attached once to the element holding the rows; it moves focus
 * with ArrowUp/ArrowDown/Home/End and hands the tab stop to whichever row takes focus.
 * `setRovingStop` picks the tab stop again after the rows are rendered.
 */

const ROW_KEYS: Record<string, (at: number, last: number) => number> = {
  ArrowDown: (at, last) => Math.min(at + 1, last),
  ArrowUp: (at) => Math.max(at - 1, 0),
  Home: () => 0,
  End: (_at, last) => last,
};

/** Give `stop` (else the first row) the tab stop and take it from the other rows. */
export function setRovingStop(rows: readonly HTMLElement[], stop?: HTMLElement | null): void {
  const target = stop && rows.includes(stop) ? stop : rows[0];
  for (const row of rows) row.tabIndex = row === target ? 0 : -1;
}

/** Move focus among `rows()` from the keyboard; returns the detach function. */
export function bindRovingRows(container: HTMLElement, rows: () => HTMLElement[]): () => void {
  const onFocusIn = (event: FocusEvent): void => {
    const current = rows();
    const row = current.find((candidate) => candidate === event.target);
    if (row) setRovingStop(current, row);
  };
  const onKeyDown = (event: KeyboardEvent): void => {
    if (event.altKey || event.ctrlKey || event.metaKey || event.shiftKey) return;
    const move = ROW_KEYS[event.key];
    if (!move) return;
    const current = rows();
    const at = current.indexOf(event.target as HTMLElement);
    if (at < 0) return;
    event.preventDefault();
    current[move(at, current.length - 1)]?.focus();
  };
  container.addEventListener('focusin', onFocusIn);
  container.addEventListener('keydown', onKeyDown);
  return () => {
    container.removeEventListener('focusin', onFocusIn);
    container.removeEventListener('keydown', onKeyDown);
  };
}
