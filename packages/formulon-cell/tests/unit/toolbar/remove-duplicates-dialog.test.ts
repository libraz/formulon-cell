import { afterEach, describe, expect, it } from 'vitest';
import { showRemoveDuplicatesDialog } from '../../../src/toolbar/dialogs/remove-duplicates.js';

describe('showRemoveDuplicatesDialog', () => {
  afterEach(() => {
    document.body.innerHTML = '';
  });

  it('renders an Excel-style column checklist and validates selection', async () => {
    const pending = showRemoveDuplicatesDialog({
      title: 'Remove Duplicates',
      columnsLabel: 'Columns',
      headerLabel: 'My data has headers',
      selectAllLabel: 'Select All',
      unselectAllLabel: 'Unselect All',
      noColumnsLabel: 'Select at least one column.',
      columns: [
        { value: '0', label: 'Column A' },
        { value: '1', label: 'Column B' },
      ],
      initialColumns: ['0', '1'],
      initialHasHeader: true,
      okLabel: 'OK',
      cancelLabel: 'Cancel',
    });

    const dialog = document.querySelector<HTMLElement>('.app__dlg');
    const list = dialog?.querySelector<HTMLElement>('.fc-dedupedlg__column-list');
    expect(list?.getAttribute('role')).toBe('group');
    expect(list?.getAttribute('aria-label')).toBe('Columns');
    expect(dialog?.querySelectorAll('.fc-dedupedlg__column')).toHaveLength(2);

    dialog?.querySelector<HTMLButtonElement>('.fc-dedupedlg__actions button:nth-child(2)')?.click();
    for (const checkbox of Array.from(
      dialog?.querySelectorAll<HTMLInputElement>('fieldset input') ?? [],
    )) {
      expect(checkbox.checked).toBe(false);
    }
    dialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    expect(dialog?.querySelector<HTMLElement>('.app__dlg__error')?.textContent).toBe(
      'Select at least one column.',
    );

    dialog?.querySelector<HTMLButtonElement>('.fc-dedupedlg__actions button')?.click();
    dialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await expect(pending).resolves.toEqual({
      columns: ['0', '1'],
      hasHeader: true,
    });
  });
});
