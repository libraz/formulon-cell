// Safari delivers the IME confirmation Enter with isComposing false.
const IME_PROCESS_KEY_CODE = 229;

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
 * or during or confirming IME composition keeps its native meaning instead.
 */
export function isSubmitEnter(event: KeyboardEvent): boolean {
  if (event.key !== 'Enter' || event.isComposing || event.keyCode === IME_PROCESS_KEY_CODE) {
    return false;
  }
  const target = event.target as Element | null;
  return !(target?.nodeType === 1 && target.closest(NON_SUBMIT_TARGETS));
}
