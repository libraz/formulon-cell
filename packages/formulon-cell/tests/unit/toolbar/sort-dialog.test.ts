import { afterEach, describe, expect, it } from 'vitest';
import { showSortDialog } from '../../../src/toolbar/dialogs/sort.js';

const labels = {
  title: 'Sort',
  columnLabel: 'Sort by',
  thenByLabel: 'Then by',
  noThenByLabel: '(none)',
  orderLabel: 'Order',
  headerLabel: 'My data has headers',
  addLevelLabel: 'Add Level',
  deleteLevelLabel: 'Delete Level',
  copyLevelLabel: 'Copy Level',
  levelUnavailableLabel: 'At least one sort level is required.',
  ascendingLabel: 'A to Z',
  descendingLabel: 'Z to A',
  okLabel: 'OK',
  cancelLabel: 'Cancel',
};

describe('showSortDialog', () => {
  afterEach(() => {
    document.body.innerHTML = '';
  });

  it('renders an Excel-style level grid and returns edited levels', async () => {
    const resultPromise = showSortDialog({
      ...labels,
      columns: [
        { value: '0', label: 'Column A' },
        { value: '1', label: 'Column B' },
      ],
      initialColumn: '0',
      initialDirection: 'asc',
      initialHasHeader: true,
    });

    const dialog = document.querySelector<HTMLElement>('.app__dlg');
    expect(dialog?.textContent).toContain('Add Level');
    expect(dialog?.textContent).toContain('Delete Level');
    expect(dialog?.textContent).toContain('Copy Level');
    expect(dialog?.querySelectorAll('.fc-sortdlg__level')).toHaveLength(1);

    const deleteLevel = dialog?.querySelector<HTMLButtonElement>('.fc-sortdlg__delete-level');
    expect(deleteLevel?.disabled).toBe(true);
    expect(deleteLevel?.dataset.disabledReason).toBe(labels.levelUnavailableLabel);
    expect(deleteLevel?.getAttribute('aria-description')).toBe(labels.levelUnavailableLabel);

    dialog?.querySelector<HTMLButtonElement>('.fc-sortdlg__add-level')?.click();
    expect(dialog?.querySelectorAll('.fc-sortdlg__level')).toHaveLength(2);
    expect(deleteLevel?.disabled).toBe(false);
    expect(deleteLevel?.dataset.disabledReason).toBeUndefined();

    const rows = Array.from(dialog?.querySelectorAll<HTMLElement>('.fc-sortdlg__level') ?? []);
    const secondColumn = rows[1]?.querySelector<HTMLSelectElement>('select[aria-label="Then by"]');
    const secondDirection = rows[1]?.querySelectorAll<HTMLSelectElement>('select')[1];
    if (secondColumn) secondColumn.value = '1';
    if (secondDirection) secondDirection.value = 'desc';

    dialog?.querySelector<HTMLButtonElement>('.fc-sortdlg__copy-level')?.click();
    expect(dialog?.querySelectorAll('.fc-sortdlg__level')).toHaveLength(3);

    dialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await expect(resultPromise).resolves.toEqual({
      column: '0',
      direction: 'asc',
      levels: [
        { column: '0', direction: 'asc' },
        { column: '1', direction: 'desc' },
        { column: '1', direction: 'desc' },
      ],
      hasHeader: true,
    });
  });
});
