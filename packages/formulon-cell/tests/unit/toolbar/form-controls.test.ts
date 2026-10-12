import { describe, expect, it } from 'vitest';
import {
  appendDialogDatalistOptions,
  appendDialogSelectOptions,
  appendSelectRow,
  createDialogSelect,
} from '../../../src/toolbar/dialogs/form-controls.js';

describe('toolbar/dialogs/form-controls', () => {
  it('creates dialog selects with shared option and accessibility contracts', () => {
    const select = createDialogSelect(
      [
        { value: 'a', label: 'Column A' },
        { value: 'b', label: 'Column B' },
      ],
      'b',
      {
        ariaLabel: 'Sort by',
        className: 'fc-tb__dlg__select',
        fieldName: 'sortColumn',
        role: 'cell',
      },
    );

    expect(select.className).toBe('fc-tb__dlg__select');
    expect(select.dataset.dialogField).toBe('sortColumn');
    expect(select.getAttribute('aria-label')).toBe('Sort by');
    expect(select.getAttribute('role')).toBe('cell');
    expect(select.value).toBe('b');
    expect(Array.from(select.options).map((option) => [option.value, option.textContent])).toEqual([
      ['a', 'Column A'],
      ['b', 'Column B'],
    ]);
  });

  it('reuses the same select creation path for labeled rows and appended options', () => {
    const body = document.createElement('div');
    const select = appendSelectRow(
      body,
      'Order',
      [
        { value: 'asc', label: 'A to Z' },
        { value: 'desc', label: 'Z to A' },
      ],
      'asc',
      'sortOrder',
    );
    appendDialogSelectOptions(select, [{ value: 'custom', label: 'Custom List...' }]);

    expect(body.querySelector('.fc-fmtdlg__row--block')).not.toBeNull();
    expect(body.querySelector('.fc-tb__dlg__label')?.textContent).toContain('Order');
    expect(select.dataset.dialogField).toBe('sortOrder');
    expect(Array.from(select.options).map((option) => option.value)).toEqual([
      'asc',
      'desc',
      'custom',
    ]);
  });

  it('keeps datalist values and display labels in the shared helper', () => {
    const datalist = document.createElement('datalist');
    appendDialogDatalistOptions(datalist, [
      'Arial',
      { value: '游ゴシック', label: '游ゴシック Regular' },
    ]);

    expect(Array.from(datalist.options, (option) => [option.value, option.label])).toEqual([
      ['Arial', undefined],
      ['游ゴシック', '游ゴシック Regular'],
    ]);
  });
});
