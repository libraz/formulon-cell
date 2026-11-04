import { afterEach, describe, expect, it, vi } from 'vitest';
import { attachAutocomplete } from '../../../src/interact/autocomplete.js';

describe('autocomplete column history guards', () => {
  afterEach(() => {
    document.body.innerHTML = '';
  });

  const run = (value: string, caret: number): number => {
    const input = document.createElement('textarea');
    input.value = value;
    document.body.append(input);
    input.setSelectionRange(caret, caret);
    const getColumnValues = vi.fn(() => ['Apple']);
    const handle = attachAutocomplete({
      input,
      editingAddr: { sheet: 0, row: 4, col: 2 },
      getColumnValues,
    });
    handle.refresh();
    handle.detach();
    return getColumnValues.mock.calls.length;
  };

  it('does not read column values for empty text or a mid-text caret', () => {
    expect(run('', 0)).toBe(0);
    expect(run('App', 1)).toBe(0);
  });

  it('reads column values when the caret is at the end of non-empty text', () => {
    expect(run('App', 3)).toBe(1);
  });
});
