const NON_SUBMIT_TARGETS = [
  'button',
  'select',
  'textarea',
  'a[href]',
  '[role="combobox"]',
  '[role="listbox"]',
  '[role="option"]',
  '.fc-select',
  '.fc-custom-select',
].join(',');

/**
 * True when an Enter keydown reaching a dialog overlay should run its primary
 * action. Enter on a button, inside a select or custom select, in a textarea,
 * or during IME composition keeps its native meaning instead.
 */
export function isSubmitEnter(event: KeyboardEvent): boolean {
  if (event.key !== 'Enter' || event.isComposing) return false;
  const target = event.target;
  return !(target instanceof Element && target.closest(NON_SUBMIT_TARGETS));
}
