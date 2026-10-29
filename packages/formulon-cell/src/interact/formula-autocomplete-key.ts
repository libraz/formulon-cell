export type FormulaAutocompleteNavigation = 'first' | 'last' | 'pageUp' | 'pageDown';

/** The stable keyboard surface shared by the inline editor and formula bar. */
export interface FormulaAutocompleteKeyTarget {
  isOpen(): boolean;
  move(delta: 1 | -1): void;
  acceptHighlighted(): boolean;
  close(): void;
  refresh(): void;
  navigate?(action: FormulaAutocompleteNavigation): void;
  dismiss?(): void;
  toggle?(): void;
}

export interface FormulaAutocompleteKeyOptions {
  formulaEdit: boolean;
}

const unmodified = (event: KeyboardEvent): boolean =>
  !event.altKey && !event.ctrlKey && !event.metaKey && !event.shiftKey;

/**
 * Apply desktop spreadsheet formula-list keys consistently on both editor
 * surfaces. Returning true means the event was consumed by the helper.
 */
export function handleFormulaAutocompleteKey(
  event: KeyboardEvent,
  autocomplete: FormulaAutocompleteKeyTarget,
  options: FormulaAutocompleteKeyOptions,
): boolean {
  const exactAltDown =
    event.key === 'ArrowDown' &&
    event.altKey &&
    !event.ctrlKey &&
    !event.metaKey &&
    !event.shiftKey;
  if (exactAltDown) {
    if (autocomplete.toggle) autocomplete.toggle();
    else if (autocomplete.isOpen()) autocomplete.close();
    else autocomplete.refresh();
    event.preventDefault();
    return true;
  }

  if (!autocomplete.isOpen()) return false;

  if (unmodified(event)) {
    if (event.key === 'ArrowDown' || event.key === 'ArrowUp') {
      autocomplete.move(event.key === 'ArrowDown' ? 1 : -1);
      event.preventDefault();
      return true;
    }

    const navigation: Partial<Record<string, FormulaAutocompleteNavigation>> = {
      Home: 'first',
      End: 'last',
      PageUp: 'pageUp',
      PageDown: 'pageDown',
    };
    const action = navigation[event.key];
    if (action) {
      if (autocomplete.navigate) autocomplete.navigate(action);
      else autocomplete.move(action === 'first' || action === 'pageUp' ? -1 : 1);
      event.preventDefault();
      return true;
    }

    if (event.key === 'Tab' || (event.key === 'Enter' && !options.formulaEdit)) {
      if (!autocomplete.acceptHighlighted()) return false;
      event.preventDefault();
      return true;
    }

    if (event.key === 'Escape') {
      if (autocomplete.dismiss) autocomplete.dismiss();
      else autocomplete.close();
      event.preventDefault();
      return true;
    }
  }
  return false;
}
