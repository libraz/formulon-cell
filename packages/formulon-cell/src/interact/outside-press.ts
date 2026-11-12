/** Which press event a dismissal listens for, and in which phase. */
export interface OutsidePressOptions {
  /** `mousedown` or `pointerdown`; each popover keeps the event it was built on. */
  readonly event: 'mousedown' | 'pointerdown';
  /** Listen in the capture phase, before the target's own handlers. Defaults to true. */
  readonly capture?: boolean;
}

/**
 * Calls `onOutside` for a press anywhere in `doc` whose target is not inside one of the
 * elements `inside` returns, read at each press so a popover built later still counts.
 * Returns the function that removes the listener.
 */
export function onOutsidePress<E extends MouseEvent = MouseEvent>(
  doc: Document,
  inside: () => readonly (Node | null | undefined)[],
  onOutside: (event: E) => void,
  options: OutsidePressOptions,
): () => void {
  const capture = options.capture ?? true;
  const listener = (event: Event): void => {
    // A press on the document targets one of its nodes, which may come from another window.
    const target = event.target as Node;
    if (inside().some((node) => node?.contains(target))) return;
    onOutside(event as E);
  };
  doc.addEventListener(options.event, listener, capture);
  return () => doc.removeEventListener(options.event, listener, capture);
}
