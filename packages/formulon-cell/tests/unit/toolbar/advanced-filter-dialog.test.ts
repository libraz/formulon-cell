import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { afterEach, describe, expect, it } from 'vitest';
import { showAdvancedFilterDialog } from '../../../src/toolbar/dialogs/advanced-filter.js';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');

describe('showAdvancedFilterDialog', () => {
  afterEach(() => {
    document.body.replaceChildren();
  });

  it('uses shared range pickers for list, criteria, and copy-to references', async () => {
    const listeners: Array<() => void> = [];
    let pickedRange = 'A1:C5';
    let pickedAddress = 'E1';
    const pending = showAdvancedFilterDialog({
      title: 'Advanced Filter',
      listRangeLabel: 'List range',
      criteriaRangeLabel: 'Criteria range',
      copyToLabel: 'Copy to',
      uniqueOnlyLabel: 'Unique records only',
      initialListRange: 'A1:B3',
      okLabel: 'OK',
      cancelLabel: 'Cancel',
      rangePickerLabel: 'Select range',
      pickRange: () => pickedRange,
      pickAddress: () => pickedAddress,
      subscribeToRangeChanges: (listener) => {
        listeners.push(listener);
        return () => {
          const idx = listeners.indexOf(listener);
          if (idx >= 0) listeners.splice(idx, 1);
        };
      },
      validateRange: () => null,
      validateAddress: () => null,
    });

    const listRange = document.querySelector<HTMLInputElement>(
      '[data-range-picker="advanced-filter-list-range"]',
    );
    const criteriaRange = document.querySelector<HTMLInputElement>(
      '[data-range-picker="advanced-filter-criteria-range"]',
    );
    const copyTo = document.querySelector<HTMLInputElement>(
      '[data-range-picker="advanced-filter-copy-to"]',
    );
    const inputs = document.querySelectorAll<HTMLInputElement>('.fc-range-picker input');
    const dialog = document.body.querySelector<HTMLElement>('.app__dlg');
    expect(dialog?.querySelector('.fc-advfilter__ranges')).toBeTruthy();
    expect(dialog?.querySelector('.fc-advfilter__row--list')).toBeTruthy();
    expect(dialog?.querySelector('.fc-advfilter__row--criteria')).toBeTruthy();
    expect(dialog?.querySelector('.fc-advfilter__row--copy-to')).toBeTruthy();
    expect(dialog?.querySelector('.fc-advfilter__option')?.textContent).toContain(
      'Unique records only',
    );
    expect(listRange?.getAttribute('aria-label')).toBe('Select range');
    expect(criteriaRange?.getAttribute('aria-label')).toBe('Select range');
    expect(copyTo?.getAttribute('aria-label')).toBe('Select range');

    listRange?.click();
    expect(listRange?.dataset.rangePickerActive).toBe('true');
    pickedRange = 'B2:D8';
    for (const listener of listeners) listener();
    expect(inputs[0]?.value).toBe('B2:D8');

    criteriaRange?.click();
    expect(listRange?.dataset.rangePickerActive).toBe('false');
    expect(criteriaRange?.dataset.rangePickerActive).toBe('true');
    pickedRange = 'F1:G2';
    for (const listener of listeners) listener();
    expect(inputs[1]?.value).toBe('F1:G2');

    copyTo?.click();
    expect(criteriaRange?.dataset.rangePickerActive).toBe('false');
    expect(copyTo?.dataset.rangePickerActive).toBe('true');
    pickedAddress = 'J4';
    for (const listener of listeners) listener();
    expect(inputs[2]?.value).toBe('J4');

    document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await expect(pending).resolves.toMatchObject({
      listRange: 'B2:D8',
      criteriaRange: 'F1:G2',
      copyTo: 'J4',
    });
  });

  it('focuses and selects the first invalid range field', async () => {
    const pending = showAdvancedFilterDialog({
      title: 'Advanced Filter',
      listRangeLabel: 'List range',
      criteriaRangeLabel: 'Criteria range',
      copyToLabel: 'Copy to',
      uniqueOnlyLabel: 'Unique records only',
      initialListRange: 'A1:B3',
      initialCriteriaRange: 'bad range',
      okLabel: 'OK',
      cancelLabel: 'Cancel',
      validateRange: (value) => (value.includes('bad') ? 'Enter a valid range.' : null),
      validateAddress: () => null,
    });

    const inputs = document.querySelectorAll<HTMLInputElement>('.fc-advfilter__row input');
    document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();

    expect(document.querySelector<HTMLElement>('.app__dlg__error')?.textContent).toBe(
      'Enter a valid range.',
    );
    expect(document.activeElement).toBe(inputs[1]);
    expect(inputs[1]?.selectionStart).toBe(0);
    expect(inputs[1]?.selectionEnd).toBe(inputs[1]?.value.length);

    document
      .querySelector<HTMLButtonElement>('.fc-fmtdlg__btn:not(.fc-fmtdlg__btn--primary)')
      ?.click();
    await expect(pending).resolves.toBeNull();
  });

  it('keeps Advanced Filter close to Excel 365 desktop range dialog geometry', () => {
    const css = readFileSync(
      join(root, 'src/styles/core/app/dialog-modules/advanced-filter.css'),
      'utf8',
    );

    expect(css).toMatch(
      /\.fc-advfilter__ranges\s*\{[\s\S]*?gap: 6px;[\s\S]*?padding: 8px 10px;[\s\S]*?border-radius: 2px;[\s\S]*?background: var\(--fc-bg, Canvas\);/,
    );
    expect(css).toMatch(
      /\.fc-advfilter__row\s*\{[\s\S]*?grid-template-columns: minmax\(130px, 0\.8fr\) minmax\(180px, 1\.2fr\);[\s\S]*?gap: 10px;/,
    );
    expect(css).toMatch(/\.fc-advfilter__option\s*\{[\s\S]*?gap: 6px;[\s\S]*?padding: 2px 0;/);
    expect(css).not.toContain('border-radius: 6px;');
    expect(css).not.toContain('color-mix(in srgb, var(--fc-fmtdlg-input-bg) 84%, transparent)');
  });
});
