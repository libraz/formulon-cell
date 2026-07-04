import { afterEach, describe, expect, it } from 'vitest';

import { showZoomDialog } from '../../../src/toolbar/dialogs/zoom.js';

describe('showZoomDialog', () => {
  afterEach(() => {
    document.body.replaceChildren();
  });

  it('returns a bounded zoom percentage', async () => {
    const pending = showZoomDialog({
      title: 'Zoom',
      label: 'Magnification',
      initial: 100,
      okLabel: 'OK',
      cancelLabel: 'Cancel',
      invalidMessage: 'Enter a zoom percentage from 50 to 400.',
    });

    const dialog = document.body.querySelector<HTMLElement>('.fc-tb__dlg');
    expect(dialog?.textContent).toContain('Magnification');
    const input = dialog?.querySelector<HTMLInputElement>('input[type="number"]');
    expect(input?.min).toBe('50');
    expect(input?.max).toBe('400');
    expect(input?.step).toBe('1');
    if (!input) throw new Error('Expected zoom input.');
    input.value = '125';
    dialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();

    await expect(pending).resolves.toBe(125);
  });

  it('shows the supplied validation message instead of reusing the field label', async () => {
    const pending = showZoomDialog({
      title: 'Zoom',
      label: 'Magnification',
      initial: 100,
      okLabel: 'OK',
      cancelLabel: 'Cancel',
      invalidMessage: 'Enter a zoom percentage from 50 to 400.',
    });

    const dialog = document.body.querySelector<HTMLElement>('.fc-tb__dlg');
    const input = dialog?.querySelector<HTMLInputElement>('input[type="number"]');
    if (!input) throw new Error('Expected zoom input.');
    input.value = '401';
    dialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();

    expect(dialog?.textContent).toContain('Enter a zoom percentage from 50 to 400.');
    input.value = '150';
    dialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await expect(pending).resolves.toBe(150);
  });
});
