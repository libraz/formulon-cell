import { afterEach, describe, expect, it } from 'vitest';
import { showTextToColumnsDialog } from '../../../src/toolbar/dialogs/text-to-columns.js';

const strings = {
  title: 'Convert Text to Columns',
  dataType: 'Original data type',
  delimited: 'Delimited',
  fixedWidth: 'Fixed width',
  fixedWidthUnavailable: 'Fixed-width splitting is not available yet.',
  delimiters: 'Delimiters',
  tab: 'Tab',
  semicolon: 'Semicolon',
  comma: 'Comma',
  space: 'Space',
  other: 'Other',
  treatConsecutive: 'Treat consecutive delimiters as one',
  preview: 'Data preview',
  noDelimited: 'No delimited text found',
  ok: 'OK',
  cancel: 'Cancel',
};

describe('showTextToColumnsDialog', () => {
  afterEach(() => {
    document.body.innerHTML = '';
  });

  it('renders an Excel-style delimited wizard surface and returns delimiter choices', async () => {
    const promise = showTextToColumnsDialog({
      strings,
      initialDelimiters: [','],
      previewRows: ['alpha,beta', 'one,two'],
    });

    const dialog = document.body.querySelector<HTMLElement>('.app__dlg');
    expect(dialog?.textContent).toContain('Original data type');
    expect(dialog?.textContent).toContain('Delimited');
    expect(dialog?.textContent).toContain('Fixed width');
    expect(dialog?.textContent).toContain('Data preview');
    expect(dialog?.querySelector('.fc-textcols__types')).toBeTruthy();
    expect(dialog?.querySelector('.fc-textcols__delimiter-grid')?.getAttribute('role')).toBe(
      'group',
    );
    expect(dialog?.querySelectorAll('.fc-textcols__delimiter')).toHaveLength(4);
    expect(dialog?.querySelector('.fc-textcols__preview')).toBeTruthy();
    expect(dialog?.querySelector('pre')?.textContent).toContain('alpha | beta');
    const fixedWidth = dialog?.querySelector<HTMLInputElement>(
      'input[name="fc-textcols-type"]:disabled',
    );
    expect(fixedWidth?.getAttribute('aria-describedby')).toBe(
      'fc-textcols-fixed-width-unavailable',
    );
    const fixedReason = dialog?.querySelector<HTMLElement>('#fc-textcols-fixed-width-unavailable');
    expect(fixedReason?.textContent).toBe('Fixed-width splitting is not available yet.');
    expect(fixedWidth?.closest('label')?.title).toBe('Fixed-width splitting is not available yet.');
    expect(fixedWidth?.title).toBe('Fixed-width splitting is not available yet.');

    const semicolon = dialog?.querySelector<HTMLInputElement>(
      '[data-dialog-field="delimiter-;"]',
    );
    const comma = dialog?.querySelector<HTMLInputElement>('[data-dialog-field="delimiter-,"]');
    const collapse = dialog?.querySelector<HTMLInputElement>(
      '[data-dialog-field="collapse-consecutive"]',
    );
    expect(comma?.checked).toBe(true);
    if (semicolon) semicolon.checked = true;
    semicolon?.dispatchEvent(new Event('change', { bubbles: true }));
    if (collapse) collapse.checked = true;
    collapse?.dispatchEvent(new Event('change', { bubbles: true }));

    dialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await expect(promise).resolves.toEqual({
      delimiters: [';', ','],
      collapseConsecutiveDelimiters: true,
    });
  });

  it('requires at least one delimiter', async () => {
    const promise = showTextToColumnsDialog({ strings, initialDelimiters: [], previewRows: ['a,b'] });
    const dialog = document.body.querySelector<HTMLElement>('.app__dlg');
    dialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    expect(dialog?.textContent).toContain('No delimited text found');
    dialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn')?.click();
    await expect(promise).resolves.toBeNull();
  });
});
