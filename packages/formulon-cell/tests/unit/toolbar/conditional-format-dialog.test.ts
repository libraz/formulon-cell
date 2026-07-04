import { afterEach, describe, expect, it } from 'vitest';

import {
  applyConditionalStylePreview,
  conditionalStyleFromValue,
  conditionalStyleOptions,
  showConditionalFormatCustomStyleDialog,
} from '../../../src/index.js';
import {
  showConditionalFormatNumberDialog,
  showConditionalFormatTextDialog,
} from '../../../src/toolbar/dialogs/conditional-format.js';

const strings = {
  ok: 'OK',
  cancel: 'Cancel',
  formatWith: 'with',
  formatPreview: 'Format preview',
  customFormat: 'Custom Format...',
  customFormatTitle: 'Format Cells',
  customFillColor: 'Fill color',
  customTextColor: 'Text color',
  customBold: 'Bold',
  customItalic: 'Italic',
  customUnderline: 'Underline',
  customStrike: 'Strikethrough',
  formatLightRed: 'Light Red Fill with Dark Red Text',
  formatYellow: 'Yellow Fill with Dark Yellow Text',
  formatGreen: 'Green Fill with Dark Green Text',
  formatLightRedFill: 'Light Red Fill',
  formatRedText: 'Red Text',
  formatRedBorder: 'Red Border',
  formatRedFill: 'Red Fill',
  formatRedTextFill: 'Red Text and Red Fill',
  invalidNumber: 'Enter a valid number.',
  invalidText: 'Enter the text to find.',
};

describe('conditional-format rule dialogs', () => {
  afterEach(() => {
    document.body.replaceChildren();
  });

  it('returns numeric thresholds with the selected Excel-style format preset', async () => {
    const pending = showConditionalFormatNumberDialog({
      title: 'Greater Than...',
      label: 'Format cells that are GREATER THAN:',
      initial: 0,
      strings,
    });

    const dialog = document.body.querySelector<HTMLElement>('.fc-tb__dlg');
    expect(dialog?.textContent).toContain('Format cells that are GREATER THAN:');
    expect(dialog?.textContent).toContain('Yellow Fill with Dark Yellow Text');
    expect(dialog?.textContent).toContain('Red Text and Red Fill');

    const input = dialog?.querySelector<HTMLInputElement>('input[type="number"]');
    const select = dialog?.querySelector<HTMLSelectElement>('select.fc-tb__dlg__select');
    const preview = dialog?.querySelector<HTMLElement>('[data-conditional-format-preview]');
    expect(select?.options).toHaveLength(9);
    expect(select?.value).toBe('light-red-dark-red');
    expect(preview?.textContent).toBe('AaBbCcYyZz');
    expect(preview?.style.background).toBe('#ffc7ce');
    if (!input || !select || !preview) {
      throw new Error('Expected number input, style select, and preview.');
    }
    input.value = '12';
    select.value = 'green-dark-green';
    select.dispatchEvent(new Event('change', { bubbles: true }));
    expect(preview.style.background).toBe('#c6efce');
    expect(preview.style.color).toBe('#006100');

    dialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await expect(pending).resolves.toEqual({
      values: [12],
      style: { fill: '#c6efce', color: '#006100' },
    });
  });

  it('exposes shared style presets and preview styling for conditional format dialogs', () => {
    const options = conditionalStyleOptions(strings);
    const preview = document.createElement('span');

    expect(options.map((option) => option.value)).toEqual([
      'light-red-dark-red',
      'yellow-dark-yellow',
      'green-dark-green',
      'light-red-fill',
      'red-text',
      'red-border',
      'red-fill',
      'red-text-fill',
      'custom',
    ]);
    expect(conditionalStyleFromValue('green-dark-green', strings)).toEqual({
      fill: '#c6efce',
      color: '#006100',
    });

    applyConditionalStylePreview(
      preview,
      { fill: '#ddeeff', color: '#123456', bold: true, strike: true },
      'Custom preview',
    );
    expect(preview.style.background).toBe('#ddeeff');
    expect(preview.style.color).toBe('#123456');
    expect(preview.style.fontWeight).toBe('700');
    expect(preview.style.textDecoration).toBe('line-through');
    expect(preview.getAttribute('aria-label')).toBe('Custom preview');
  });

  it('returns an expanded Excel-style red preset from the shared preset list', async () => {
    const pending = showConditionalFormatNumberDialog({
      title: 'Equal To...',
      label: 'Format cells that are EQUAL TO:',
      initial: 1,
      strings,
    });

    const dialog = document.body.querySelector<HTMLElement>('.fc-tb__dlg');
    const input = dialog?.querySelector<HTMLInputElement>('input[type="number"]');
    const select = dialog?.querySelector<HTMLSelectElement>('select.fc-tb__dlg__select');
    const preview = dialog?.querySelector<HTMLElement>('[data-conditional-format-preview]');
    if (!input || !select || !preview) {
      throw new Error('Expected number input, style select, and preview.');
    }
    input.value = '7';
    select.value = 'red-text-fill';
    select.dispatchEvent(new Event('change', { bubbles: true }));
    expect(preview.style.background).toBe('#ff0000');
    expect(preview.style.color).toBe('#ffffff');

    dialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await expect(pending).resolves.toEqual({
      values: [7],
      style: { fill: '#ff0000', color: '#ffffff' },
    });
  });

  it('returns text criteria with a text-only format preset', async () => {
    const pending = showConditionalFormatTextDialog({
      title: 'Text that Contains...',
      label: 'Text to contain',
      strings,
    });

    const dialog = document.body.querySelector<HTMLElement>('.fc-tb__dlg');
    const input = dialog?.querySelector<HTMLInputElement>('input[type="text"]');
    const select = dialog?.querySelector<HTMLSelectElement>('select.fc-tb__dlg__select');
    const preview = dialog?.querySelector<HTMLElement>('[data-conditional-format-preview]');
    if (!input || !select || !preview) {
      throw new Error('Expected text input, style select, and preview.');
    }
    input.value = 'late';
    select.value = 'red-text';
    select.dispatchEvent(new Event('change', { bubbles: true }));
    expect(preview.style.color).toBe('#9c0006');

    dialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await expect(pending).resolves.toEqual({
      text: 'late',
      style: { color: '#9c0006' },
    });
  });

  it('opens a shared custom-format dialog and returns the custom rule patch', async () => {
    const pending = showConditionalFormatNumberDialog({
      title: 'Top 10 Items',
      label: 'Format values that rank in the:',
      initial: 10,
      strings,
    });

    const parent = document.body.querySelector<HTMLElement>('.fc-tb__dlg');
    const select = parent?.querySelector<HTMLSelectElement>('select.fc-tb__dlg__select');
    const preview = parent?.querySelector<HTMLElement>('[data-conditional-format-preview]');
    if (!select || !preview) throw new Error('Expected parent style select and preview.');
    select.value = 'custom';
    select.dispatchEvent(new Event('change', { bubbles: true }));
    await Promise.resolve();

    const dialogs = Array.from(document.body.querySelectorAll<HTMLElement>('.fc-tb__dlg'));
    const customDialog = dialogs.find((dialog) => dialog.textContent?.includes('Format Cells'));
    expect(customDialog?.textContent).toContain('Fill color');
    expect(customDialog?.textContent).toContain('Strikethrough');

    const inputs = Array.from(customDialog?.querySelectorAll<HTMLInputElement>('input') ?? []);
    const fill = inputs.find((input) => input.placeholder === '#ffc7ce');
    const color = inputs.find((input) => input.placeholder === '#9c0006');
    const bold = inputs.find((input) => input.type === 'checkbox');
    if (!fill || !color || !bold) throw new Error('Expected custom format controls.');
    fill.value = '#ddeeff';
    color.value = '#123456';
    bold.checked = true;
    customDialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await Promise.resolve();

    expect(preview.style.background).toBe('#ddeeff');
    expect(preview.style.color).toBe('#123456');
    expect(preview.style.fontWeight).toBe('700');
    parent?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await expect(pending).resolves.toEqual({
      values: [10],
      style: { fill: '#ddeeff', color: '#123456', bold: true },
    });
  });

  it('allows conditional formatting custom styles to be opened as a reusable dialog', async () => {
    const pending = showConditionalFormatCustomStyleDialog(strings, {
      fill: '#eeeeee',
      color: '#222222',
      italic: true,
    });
    const dialog = document.body.querySelector<HTMLElement>('.fc-tb__dlg');
    expect(dialog?.textContent).toContain('Format Cells');

    const fill = dialog?.querySelector<HTMLInputElement>('input[placeholder="#ffc7ce"]');
    const color = dialog?.querySelector<HTMLInputElement>('input[placeholder="#9c0006"]');
    const italic = dialog?.querySelector<HTMLInputElement>('[data-dialog-field="italic"]');
    const underline = dialog?.querySelector<HTMLInputElement>('[data-dialog-field="underline"]');
    if (!fill || !color || !italic || !underline) {
      throw new Error('Expected shared custom style fields.');
    }
    expect(fill.value).toBe('#eeeeee');
    expect(color.value).toBe('#222222');
    expect(italic.checked).toBe(true);
    fill.value = '#ffeecc';
    color.value = '#003366';
    underline.checked = true;
    dialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();

    await expect(pending).resolves.toEqual({
      fill: '#ffeecc',
      color: '#003366',
      italic: true,
      underline: true,
    });
  });
});
