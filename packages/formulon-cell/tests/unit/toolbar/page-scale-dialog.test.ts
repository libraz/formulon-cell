import { afterEach, describe, expect, it } from 'vitest';

import { showPageScaleDialog } from '../../../src/toolbar/dialogs/page-scale.js';

describe('showPageScaleDialog', () => {
  afterEach(() => {
    document.body.replaceChildren();
  });

  it('bounds custom scale percentages to Excel-like limits', async () => {
    const pending = showPageScaleDialog({
      title: 'Scale',
      label: 'Adjust to',
      initial: 100,
      kind: 'scale',
      okLabel: 'OK',
      cancelLabel: 'Cancel',
      invalidMessage: 'Enter a scale from 10 to 400.',
    });

    const input = document.body.querySelector<HTMLInputElement>('input[type="number"]');
    expect(input?.min).toBe('10');
    expect(input?.max).toBe('400');
    expect(input?.step).toBe('1');
    if (!input) throw new Error('Expected scale input.');
    input.value = '80';
    document.body.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();

    await expect(pending).resolves.toBe(80);
  });

  it('bounds custom fit-to pages to page-count limits', async () => {
    const pending = showPageScaleDialog({
      title: 'Width',
      label: 'Pages',
      initial: 1,
      kind: 'pages',
      okLabel: 'OK',
      cancelLabel: 'Cancel',
      invalidMessage: 'Enter pages from 1 to 99.',
    });

    const input = document.body.querySelector<HTMLInputElement>('input[type="number"]');
    expect(input?.min).toBe('1');
    expect(input?.max).toBe('99');
    expect(input?.step).toBe('1');
    if (!input) throw new Error('Expected pages input.');
    input.value = '2';
    document.body.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();

    await expect(pending).resolves.toBe(2);
  });
});
