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
        titleLabel: 'Title',
        password: 'Password',
        passwordPlaceholder: 'Optional',
        confirmPassword: 'Re-enter password',
        passwordMismatch: 'The password confirmation does not match.',
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
    await expect(promise).resolves.toEqual({ range: 'C3:C8', title: '' });
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

describe('showAllowEditRangeDialog password field', () => {
  afterEach(() => {
    document.body.innerHTML = '';
  });

  const open = (hasPassword: boolean) =>
    showAllowEditRangeDialog({
      strings: {
        title: 'Allow Users to Edit Ranges',
        range: 'Range',
        invalid: 'Enter a range such as A1:B10.',
        rangePickerLabel: 'Select range',
        titleLabel: 'Title',
        password: 'Password',
        passwordPlaceholder: 'Optional',
        confirmPassword: 'Re-enter password',
        passwordMismatch: 'The password confirmation does not match.',
        ok: 'OK',
        cancel: 'Cancel',
      },
      initialRange: 'A1',
      hasPassword,
      pickRange: () => 'A1',
      validateRange: () => true,
      subscribeToRangeChanges: () => () => undefined,
    });
  const passwordInputs = (): HTMLInputElement[] =>
    Array.from(document.querySelectorAll<HTMLInputElement>('input[type="password"]'));
  const type = (input: HTMLInputElement, value: string): void => {
    input.value = value;
    input.dispatchEvent(new Event('input', { bubbles: true }));
  };
  const confirm = (): void =>
    document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();

  it('leaves the password unchanged when the field was never touched', async () => {
    const promise = open(true);
    expect(passwordInputs().every((input) => input.value !== '')).toBe(true);
    confirm();
    await expect(promise).resolves.toEqual({ range: 'A1', title: '' });
  });

  it('removes the password when the field is cleared', async () => {
    const promise = open(true);
    type(passwordInputs()[0] as HTMLInputElement, '');
    confirm();
    await expect(promise).resolves.toEqual({ range: 'A1', title: '', password: null });
  });

  it('sets a new password only when both fields match', async () => {
    const promise = open(true);
    const [password, again] = passwordInputs() as [HTMLInputElement, HTMLInputElement];
    type(password, 'fresh');
    confirm();
    expect(document.querySelector('.fc-tb__dlg')).toBeTruthy();
    type(again, 'fresh');
    confirm();
    await expect(promise).resolves.toEqual({ range: 'A1', title: '', password: 'fresh' });
  });
});
