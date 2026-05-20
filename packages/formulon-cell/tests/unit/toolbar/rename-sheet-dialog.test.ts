import { afterEach, describe, expect, it } from 'vitest';

import { showRenameSheetDialog } from '../../../src/toolbar/dialogs/rename-sheet.js';

describe('showRenameSheetDialog', () => {
  afterEach(() => {
    document.body.replaceChildren();
  });

  it('returns the trimmed sheet name from the dedicated rename dialog', async () => {
    const pending = showRenameSheetDialog({
      title: 'Rename',
      label: 'Sheet name',
      initial: 'Sheet1',
      requiredMessage: 'Enter a sheet name.',
      okLabel: 'OK',
      cancelLabel: 'Cancel',
    });

    const dialog = document.body.querySelector<HTMLElement>('.app__dlg');
    expect(dialog?.textContent).toContain('Sheet name');
    const input = dialog?.querySelector<HTMLInputElement>('input');
    if (!input) throw new Error('Expected sheet name input.');
    input.value = '  Summary  ';
    dialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();

    await expect(pending).resolves.toBe('Summary');
  });

  it('shows a localized validation message for an empty sheet name', async () => {
    const pending = showRenameSheetDialog({
      title: 'Rename',
      label: 'Sheet name',
      initial: 'Sheet1',
      requiredMessage: 'Enter a sheet name.',
      okLabel: 'OK',
      cancelLabel: 'Cancel',
    });

    const dialog = document.body.querySelector<HTMLElement>('.app__dlg');
    const input = dialog?.querySelector<HTMLInputElement>('input');
    if (!input) throw new Error('Expected sheet name input.');
    input.value = '   ';
    dialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    expect(dialog?.textContent).toContain('Enter a sheet name.');
    dialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn:not(.fc-fmtdlg__btn--primary)')?.click();

    await expect(pending).resolves.toBeNull();
  });
});
