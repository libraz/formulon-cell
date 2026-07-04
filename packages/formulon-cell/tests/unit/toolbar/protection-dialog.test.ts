import { afterEach, describe, expect, it } from 'vitest';
import {
  showAllowEditRangeDialog,
  showUnprotectSheetDialog,
} from '../../../src/toolbar/dialogs/protection.js';

describe('showAllowEditRangeDialog', () => {
  afterEach(() => {
    document.body.innerHTML = '';
  });

  it('uses the shared range picker for the editable range input', async () => {
    let picked = 'B2:D4';
    const listeners: Array<() => void> = [];
    const promise = showAllowEditRangeDialog({
      strings: {
        title: 'Allow Users to Edit Ranges',
        range: 'Range',
        invalid: 'Enter a range such as A1:B10.',
        rangePickerLabel: 'Select range',
        ok: 'OK',
        cancel: 'Cancel',
      },
      initialRange: 'A1',
      pickRange: () => picked,
      validateRange: (value) => value.length > 0,
      subscribeToRangeChanges: (listener) => {
        listeners.push(listener);
        return () => undefined;
      },
    });

    const input = document.querySelector<HTMLInputElement>('.fc-tb__dlg__input');
    const picker = document.querySelector<HTMLButtonElement>(
      '[data-range-picker="allow-edit-ranges-range"]',
    );
    expect(input?.value).toBe('A1');
    expect(picker?.getAttribute('aria-label')).toBe('Select range');

    picker?.click();
    expect(input?.value).toBe('B2:D4');
    expect(picker?.getAttribute('aria-pressed')).toBe('true');
    expect(document.querySelector('.fc-fmtdlg--range-picking')).toBeTruthy();
    picked = 'C3:C8';
    listeners.at(-1)?.();
    expect(input?.value).toBe('C3:C8');

    document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await expect(promise).resolves.toBe('C3:C8');
    expect(document.querySelector('.fc-fmtdlg--range-picking')).toBeNull();
  });
});

describe('showUnprotectSheetDialog', () => {
  afterEach(() => {
    document.body.innerHTML = '';
  });

  it('renders a dedicated password dialog and returns the entered password', async () => {
    const promise = showUnprotectSheetDialog({
      title: 'Unprotect Sheet',
      password: 'Password',
      ok: 'OK',
      cancel: 'Cancel',
    });

    const dialog = document.querySelector<HTMLElement>('.fc-tb__dlg');
    expect(dialog?.textContent).toContain('Unprotect Sheet');
    const input = dialog?.querySelector<HTMLInputElement>('input[type="password"]');
    expect(input).toBeTruthy();
    if (!input) throw new Error('Expected Unprotect Sheet password input.');
    input.value = 'pw';
    dialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();

    await expect(promise).resolves.toBe('pw');
  });
});
