/**
 * Take pointer capture on `el`. `setPointerCapture` throws on synthetic events and on
 * pointer-id mismatches; the gesture then still works while the pointer stays inside `el`.
 */
export function capturePointerSafely(el: Element, pointerId: number): void {
  try {
    el.setPointerCapture(pointerId);
  } catch {
    /* synthetic events / id mismatch: the drag still works inside the element */
  }
}

/** Release pointer capture on `el` when it holds it. */
export function releaseHeldPointer(el: Element, pointerId: number): void {
  if (el.hasPointerCapture(pointerId)) el.releasePointerCapture(pointerId);
}
