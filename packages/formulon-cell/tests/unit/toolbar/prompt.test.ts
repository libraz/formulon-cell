import { afterEach, describe, expect, it } from 'vitest';
import { showNumberPrompt } from '../../../src/toolbar/dialogs/prompt.js';

describe('showNumberPrompt', () => {
  afterEach(() => {
    document.body.replaceChildren();
  });

  it('uses the active decimal grammar for non-integer prompts', async () => {
    const pending = showNumberPrompt({
      title: 'Tolerance',
      label: 'Maximum change',
      initial: 1.25,
      min: 0,
      step: 0.01,
      okLabel: 'OK',
      cancelLabel: 'Cancel',
      invalidMessage: 'Enter a valid number.',
      locale: () => 'de-DE',
    });

    const input = document.querySelector<HTMLInputElement>('input[type="text"]');
    expect(input?.value).toBe('1,25');
    expect(input?.inputMode).toBe('decimal');
    if (!input) throw new Error('Expected decimal prompt input.');
    input.value = '2,5';
    document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();

    await expect(pending).resolves.toBe(2.5);
  });

  it('rejects malformed localized numbers instead of accepting a prefix', async () => {
    const pending = showNumberPrompt({
      title: 'Tolerance',
      label: 'Maximum change',
      initial: 1.25,
      step: 0.01,
      okLabel: 'OK',
      cancelLabel: 'Cancel',
      invalidMessage: 'Enter a valid number.',
      locale: 'de-DE',
    });

    const input = document.querySelector<HTMLInputElement>('input[type="text"]');
    if (!input) throw new Error('Expected decimal prompt input.');
    input.value = '1,2,3';
    document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    expect(document.querySelector('.fc-tb__dlg__error')?.textContent).toBe('Enter a valid number.');

    input.value = '1,5';
    document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await expect(pending).resolves.toBe(1.5);
  });

  it('keeps fractional values when step is only a spinner increment', async () => {
    const pending = showNumberPrompt({
      title: 'Row height',
      label: 'Height',
      initial: 12.5,
      min: 1,
      step: 1,
      okLabel: 'OK',
      cancelLabel: 'Cancel',
      invalidMessage: 'Enter a valid number.',
      locale: 'de-DE',
    });

    const input = document.querySelector<HTMLInputElement>('input[type="text"]');
    expect(input?.value).toBe('12,5');
    if (!input) throw new Error('Expected decimal prompt input.');
    input.value = '13,25';
    document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();

    await expect(pending).resolves.toBe(13.25);
  });

  it('applies whole-number validation only when explicitly requested', async () => {
    const pending = showNumberPrompt({
      title: 'Zoom',
      label: 'Percent',
      initial: 100,
      min: 50,
      max: 400,
      step: 1,
      integer: true,
      okLabel: 'OK',
      cancelLabel: 'Cancel',
      invalidMessage: 'Enter a whole number.',
      locale: 'de-DE',
    });

    const input = document.querySelector<HTMLInputElement>('input[type="number"]');
    if (!input) throw new Error('Expected integer prompt input.');
    input.value = '100,5';
    document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    expect(document.querySelector('.fc-tb__dlg__error')?.textContent).toBe('Enter a whole number.');

    input.value = '101';
    document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await expect(pending).resolves.toBe(101);
  });
});
