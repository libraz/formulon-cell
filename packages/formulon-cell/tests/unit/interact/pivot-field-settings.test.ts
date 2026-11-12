import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';

import { PivotAxis, PivotFilterKind } from '../../../src/engine/types.js';
import { en } from '../../../src/i18n/strings.js';
import {
  createPivotAreaSettingsButton,
  pivotFilterConditionToSpec,
  pivotFilterSpecToCondition,
  splitFilterConditionRange,
} from '../../../src/interact/pivot-field-settings.js';
import { showPivotFilterDialog } from '../../../src/interact/pivot-filter-dialog.js';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');

describe('pivot-field-settings shared filter condition model', () => {
  it('creates Pivot area settings buttons with the shared class and aria contract', () => {
    const button = createPivotAreaSettingsButton('Field Settings', 'Field Settings: Sales');
    expect(button.type).toBe('button');
    expect(button.className).toBe('fc-pivotdlg__area-settings');
    expect(button.textContent).toBe('Field Settings');
    expect(button.getAttribute('aria-label')).toBe('Field Settings: Sales');
  });

  it('keeps Pivot settings buttons on the shared helper', () => {
    const source = readFileSync(join(root, 'src/interact/pivot-field-settings.ts'), 'utf8');
    expect(source.match(/document\.createElement\('button'\)/g)).toBeNull();
  });

  it('converts label filter conditions into PivotFilterSpec payloads', () => {
    expect(
      pivotFilterConditionToSpec('Region', { kind: PivotFilterKind.LabelContains, value: 'East' }),
    ).toEqual({
      axis: PivotAxis.Page,
      fieldName: 'Region',
      kind: PivotFilterKind.LabelContains,
      value: 'East',
    });
    expect(
      pivotFilterConditionToSpec('Region', { kind: PivotFilterKind.LabelBetween, value: 'A..M' }),
    ).toEqual({
      axis: PivotAxis.Page,
      fieldName: 'Region',
      kind: PivotFilterKind.LabelBetween,
      value: 'A',
      valueHigh: 'M',
    });
  });

  it('converts numeric range conditions into low and high bounds', () => {
    for (const [text, low, high] of [
      ['10..20', 10, 20],
      ['-10..20', -10, 20],
      ['+1.5e2..-2.5e1', 150, -25],
    ] as const) {
      expect(
        pivotFilterConditionToSpec('Sales', { kind: PivotFilterKind.ValueNotBetween, value: text }),
      ).toEqual({
        axis: PivotAxis.Page,
        fieldName: 'Sales',
        kind: PivotFilterKind.ValueNotBetween,
        value: low,
        valueHigh: high,
      });
    }
  });

  it('splits a typed number range on its separator, keeping signed bounds', () => {
    expect(splitFilterConditionRange('-10..20')).toEqual(['-10', '20']);
    expect(splitFilterConditionRange('10--20')).toEqual(['10', '-20']);
    expect(splitFilterConditionRange('+1.5e2..-2.5e1')).toEqual(['+1.5e2', '-2.5e1']);
  });

  it('drops empty and unreadable conditions', () => {
    expect(pivotFilterConditionToSpec('Sales', { kind: null, value: '' })).toBeNull();
    expect(
      pivotFilterConditionToSpec('Sales', {
        kind: PivotFilterKind.ValueGreaterThan,
        value: 'large',
      }),
    ).toBeNull();
    expect(
      pivotFilterConditionToSpec('Sales', { kind: PivotFilterKind.ValueBetween, value: '10..' }),
    ).toBeNull();
    expect(
      pivotFilterConditionToSpec('Date', { kind: PivotFilterKind.DateEqual, value: 'soon' }),
    ).toBeNull();
  });

  it('reads numbers and dates through the locale input grammar', () => {
    expect(
      pivotFilterConditionToSpec(
        'Sales',
        { kind: PivotFilterKind.ValueLessThan, value: '1.234,5' },
        PivotAxis.Page,
        { locale: 'de-DE' },
      ),
    ).toMatchObject({ kind: PivotFilterKind.ValueLessThan, value: 1234.5 });
    // 2026-05-19 is serial 46161 in the 1900 date system.
    expect(
      pivotFilterConditionToSpec('Date', {
        kind: PivotFilterKind.DateNewerThan,
        value: '5/19/2026',
      }),
    ).toMatchObject({ kind: PivotFilterKind.DateNewerThan, value: 46161 });
    expect(
      pivotFilterConditionToSpec(
        'Date',
        { kind: PivotFilterKind.DateBetween, value: '19.5.2026..31.5.2026' },
        PivotAxis.Page,
        { locale: 'de-DE' },
      ),
    ).toMatchObject({ kind: PivotFilterKind.DateBetween, value: 46161, valueHigh: 46173 });
  });

  it('converts top-count and relative date conditions', () => {
    expect(
      pivotFilterConditionToSpec('Sales', { kind: PivotFilterKind.ValueTopCount, value: '5' }),
    ).toEqual({
      axis: PivotAxis.Page,
      fieldName: 'Sales',
      kind: PivotFilterKind.ValueTopCount,
      value: 5,
    });
    expect(
      pivotFilterConditionToSpec('Date', { kind: PivotFilterKind.LastQuarter, value: '' }),
    ).toEqual({ axis: PivotAxis.Page, fieldName: 'Date', kind: PivotFilterKind.LastQuarter });
  });

  it('converts PivotFilterSpec payloads back into shared condition state', () => {
    const spec = (kind: PivotFilterKind, value?: number | string, valueHigh?: number | string) => ({
      axis: PivotAxis.Page,
      fieldName: 'F',
      kind,
      ...(value !== undefined ? { value } : {}),
      ...(valueHigh !== undefined ? { valueHigh } : {}),
    });
    expect(pivotFilterSpecToCondition(spec(PivotFilterKind.ValueBetween, -10, 20))).toEqual({
      kind: PivotFilterKind.ValueBetween,
      value: '-10..20',
    });
    expect(pivotFilterSpecToCondition(spec(PivotFilterKind.LabelEndsWith, 'st'))).toEqual({
      kind: PivotFilterKind.LabelEndsWith,
      value: 'st',
    });
    expect(pivotFilterSpecToCondition(spec(PivotFilterKind.DateBetween, 46161, 46173))).toEqual({
      kind: PivotFilterKind.DateBetween,
      value: '2026-05-19..2026-05-31',
    });
    expect(pivotFilterSpecToCondition(spec(PivotFilterKind.Yesterday))).toEqual({
      kind: PivotFilterKind.Yesterday,
      value: '',
    });
  });

  it('round-trips a shown date back to the same serial', () => {
    const condition = pivotFilterSpecToCondition({
      axis: PivotAxis.Page,
      fieldName: 'Date',
      kind: PivotFilterKind.DateOlderThan,
      value: 46161,
    });
    if (!condition) throw new Error('missing condition');
    expect(pivotFilterConditionToSpec('Date', condition)?.value).toBe(46161);
  });

  it('opens the value filter dialog on the condition kind and returns the edited condition', async () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const result = showPivotFilterDialog({
      host,
      strings: en.pivotTableDialog,
      fieldName: 'Sales',
      condition: { kind: PivotFilterKind.ValueGreaterThan, value: '100' },
      okLabel: 'OK',
      cancelLabel: 'Cancel',
      mac: false,
    });
    const dialog = document.body.querySelector<HTMLElement>('.fc-pivotdlg');
    expect(dialog?.textContent).toContain('Value Filter (Sales)');
    expect(dialog?.classList.contains('fc-pivotdlg--filter')).toBe(true);
    const operator = dialog?.querySelector<HTMLSelectElement>(
      'select[data-pivot-filter-operator="true"]',
    );
    if (!operator || !dialog) throw new Error('missing dialog controls');
    await new Promise((resolve) => requestAnimationFrame(resolve));
    expect(document.activeElement).toBe(
      operator.closest('.fc-select')?.querySelector<HTMLButtonElement>('.fc-select__button'),
    );
    expect(operator.value).toBe(String(PivotFilterKind.ValueGreaterThan));
    operator.value = String(PivotFilterKind.ValueBetween);
    operator.dispatchEvent(new Event('change', { bubbles: true }));
    const low = dialog.querySelector<HTMLInputElement>('input[data-pivot-filter-value="true"]');
    const high = dialog.querySelector<HTMLInputElement>(
      'input[data-pivot-filter-value-high="true"]',
    );
    if (!low || !high) throw new Error('missing range inputs');
    low.value = '10';
    high.value = '20';
    dialog.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await expect(result).resolves.toEqual({ kind: PivotFilterKind.ValueBetween, value: '10..20' });
    expect(document.body.querySelector('.fc-pivotdlg')).toBeNull();
    host.remove();
  });
});
