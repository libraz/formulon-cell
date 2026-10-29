import { describe, expect, it, vi } from 'vitest';
import { handleFormulaAutocompleteKey } from '../../../src/interact/formula-autocomplete-key.js';

const openAutocomplete = () => {
  const autocomplete = {
    isOpen: () => true,
    move: vi.fn(),
    acceptHighlighted: vi.fn(() => true),
    close: vi.fn(),
    refresh: vi.fn(),
    navigate: vi.fn(),
    dismiss: vi.fn(),
    toggle: vi.fn(),
  };
  return autocomplete;
};

const key = (name: string, init: KeyboardEventInit = {}) =>
  new KeyboardEvent('keydown', { key: name, cancelable: true, ...init });

describe('formula autocomplete keyboard policy', () => {
  it('inserts with plain Tab and leaves formula Enter for typed-formula commit', () => {
    const autocomplete = openAutocomplete();
    const tab = key('Tab');
    expect(handleFormulaAutocompleteKey(tab, autocomplete, { formulaEdit: true })).toBe(true);
    expect(autocomplete.acceptHighlighted).toHaveBeenCalledOnce();
    expect(tab.defaultPrevented).toBe(true);

    const enter = key('Enter');
    expect(handleFormulaAutocompleteKey(enter, autocomplete, { formulaEdit: true })).toBe(false);
    expect(autocomplete.acceptHighlighted).toHaveBeenCalledOnce();
  });

  it('preserves plain column-history Enter acceptance and modified Enter or Tab behavior', () => {
    const autocomplete = openAutocomplete();
    expect(handleFormulaAutocompleteKey(key('Enter'), autocomplete, { formulaEdit: false })).toBe(
      true,
    );
    expect(autocomplete.acceptHighlighted).toHaveBeenCalledOnce();

    for (const event of [
      key('Enter', { altKey: true }),
      key('Enter', { ctrlKey: true }),
      key('Enter', { metaKey: true }),
      key('Enter', { shiftKey: true }),
      key('Tab', { shiftKey: true }),
      key('Tab', { ctrlKey: true }),
    ]) {
      expect(handleFormulaAutocompleteKey(event, autocomplete, { formulaEdit: false })).toBe(false);
    }
    expect(autocomplete.acceptHighlighted).toHaveBeenCalledOnce();
  });

  it('routes navigation, exact Alt+Down, and first Escape dismissal', () => {
    const autocomplete = openAutocomplete();
    for (const [name, action] of [
      ['ArrowDown', null],
      ['ArrowUp', null],
      ['Home', 'first'],
      ['End', 'last'],
      ['PageUp', 'pageUp'],
      ['PageDown', 'pageDown'],
    ] as const) {
      expect(handleFormulaAutocompleteKey(key(name), autocomplete, { formulaEdit: true })).toBe(
        true,
      );
      if (action) expect(autocomplete.navigate).toHaveBeenLastCalledWith(action);
    }

    expect(
      handleFormulaAutocompleteKey(key('ArrowDown', { altKey: true }), autocomplete, {
        formulaEdit: true,
      }),
    ).toBe(true);
    expect(autocomplete.toggle).toHaveBeenCalledOnce();

    expect(handleFormulaAutocompleteKey(key('Escape'), autocomplete, { formulaEdit: true })).toBe(
      true,
    );
    expect(autocomplete.dismiss).toHaveBeenCalledOnce();
    expect(autocomplete.close).not.toHaveBeenCalled();
  });

  it('uses old handle methods when the optional methods are absent', () => {
    const autocomplete = {
      isOpen: () => true,
      move: vi.fn(),
      acceptHighlighted: vi.fn(() => false),
      close: vi.fn(),
      refresh: vi.fn(),
    };
    expect(
      handleFormulaAutocompleteKey(key('ArrowDown'), autocomplete, { formulaEdit: true }),
    ).toBe(true);
    expect(autocomplete.move).toHaveBeenCalledWith(1);
    expect(handleFormulaAutocompleteKey(key('Escape'), autocomplete, { formulaEdit: true })).toBe(
      true,
    );
    expect(autocomplete.close).toHaveBeenCalledOnce();
  });
});
