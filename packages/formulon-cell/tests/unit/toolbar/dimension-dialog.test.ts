import { afterEach, describe, expect, it } from 'vitest';

import { showDimensionDialog } from '../../../src/toolbar/dialogs/dimension.js';

describe('showDimensionDialog', () => {
  afterEach(() => {
    document.body.replaceChildren();
  });

  it('returns a bounded numeric dimension value', async () => {
    const pending = showDimensionDialog({
      title: 'Row Height',
      label: 'Height (px)',
      initial: 24,
      max: 409,
      okLabel: 'OK',
      cancelLabel: 'Cancel',
    });

    const dialog = document.body.querySelector<HTMLElement>('.fc-tb__dlg');
    expect(dialog?.textContent).toContain('Height (px)');
    const input = dialog?.querySelector<HTMLInputElement>('input[type="number"]');
    expect(input?.min).toBe('1');
    expect(input?.max).toBe('409');
    expect(input?.step).toBe('1');
    if (!input) throw new Error('Expected dimension input.');
    input.value = '48';
    dialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();

    await expect(pending).resolves.toBe(48);
  });
});
