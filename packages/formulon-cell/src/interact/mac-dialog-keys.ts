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

/** What else an Enter may act on besides a dialog's primary action. */
export interface SubmitEnterOptions {
  /** A control, such as a combobox, whose own Enter runs the action instead of keeping its
   *  native meaning. */
  readonly actsOn?: Element;
}

/**
 * True when an Enter keydown reaching a dialog overlay should run its primary
 * action. Enter on a button, inside a select or custom select, in a textarea,
 * or during or confirming IME composition keeps its native meaning instead.
 */
export function isSubmitEnter(event: KeyboardEvent, options: SubmitEnterOptions = {}): boolean {
  if (event.key !== 'Enter' || event.isComposing || event.keyCode === IME_PROCESS_KEY_CODE) {
    return false;
  }
  const target = event.target as Element | null;
  if (options.actsOn && target === options.actsOn) return true;
  return !(target?.nodeType === 1 && target.closest(NON_SUBMIT_TARGETS));
}
