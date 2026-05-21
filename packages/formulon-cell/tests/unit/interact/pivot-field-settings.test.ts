import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';

import { PivotAxis, PivotFilterType, PivotFilterValueKind } from '../../../src/engine/types.js';
import { en } from '../../../src/i18n/strings.js';
import {
  createPivotAreaSettingsButton,
  createPivotFilterConditionControls,
  pivotFilterConditionToSpec,
  pivotFilterSpecToCondition,
  showPivotFilterDialog,
  splitFilterConditionRange,
} from '../../../src/interact/pivot-field-settings.js';

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
    expect(source).toContain('createPivotAreaSettingsButton(t.filterDialog)');
    expect(source.match(/document\.createElement\('button'\)/g)).toBeNull();
  });

  it('converts label filter conditions into PivotFilterSpec payloads', () => {
    expect(pivotFilterConditionToSpec('Region', { kind: 'label-contains', value: 'East' })).toEqual(
      {
        axis: PivotAxis.Page,
        fieldName: 'Region',
        type: PivotFilterType.LabelContains,
        valueKind: PivotFilterValueKind.Text,
        valueText: 'East',
      },
    );
    expect(pivotFilterConditionToSpec('Region', { kind: 'label-equals', value: 'East' })).toEqual({
      axis: PivotAxis.Page,
      fieldName: 'Region',
      type: PivotFilterType.LabelEquals,
      valueKind: PivotFilterValueKind.Text,
      valueText: 'East',
    });
    expect(
      pivotFilterConditionToSpec('Region', { kind: 'label-does-not-contain', value: 'East' }),
    ).toEqual({
      axis: PivotAxis.Page,
      fieldName: 'Region',
      type: PivotFilterType.LabelDoesNotContain,
      valueKind: PivotFilterValueKind.Text,
      valueText: 'East',
    });
  });

  it('converts value-between filter conditions into low/high numeric payloads', () => {
    expect(pivotFilterConditionToSpec('Sales', { kind: 'value-between', value: '10..20' })).toEqual(
      {
        axis: PivotAxis.Page,
        fieldName: 'Sales',
        type: PivotFilterType.ValueBetween,
        valueKind: PivotFilterValueKind.Double,
        valueDouble: 10,
        valueHighKind: PivotFilterValueKind.Double,
        valueHighDouble: 20,
      },
    );
    expect(
      pivotFilterConditionToSpec('Margin', { kind: 'value-between', value: '-10..20' }),
    ).toEqual({
      axis: PivotAxis.Page,
      fieldName: 'Margin',
      type: PivotFilterType.ValueBetween,
      valueKind: PivotFilterValueKind.Double,
      valueDouble: -10,
      valueHighKind: PivotFilterValueKind.Double,
      valueHighDouble: 20,
    });
    expect(splitFilterConditionRange('-10..20')).toEqual(['-10', '20']);
    expect(splitFilterConditionRange('10--20')).toEqual(['10', '-20']);
    expect(splitFilterConditionRange('+1.5e2..-2.5e1')).toEqual(['+1.5e2', '-2.5e1']);
    expect(
      pivotFilterConditionToSpec('Margin', { kind: 'value-between', value: '+1.5e2..-2.5e1' }),
    ).toEqual({
      axis: PivotAxis.Page,
      fieldName: 'Margin',
      type: PivotFilterType.ValueBetween,
      valueKind: PivotFilterValueKind.Double,
      valueDouble: 150,
      valueHighKind: PivotFilterValueKind.Double,
      valueHighDouble: -25,
    });
    expect(
      pivotFilterConditionToSpec('Sales', { kind: 'value-not-between', value: '10..20' }),
    ).toEqual({
      axis: PivotAxis.Page,
      fieldName: 'Sales',
      type: PivotFilterType.ValueNotBetween,
      valueKind: PivotFilterValueKind.Double,
      valueDouble: 10,
      valueHighKind: PivotFilterValueKind.Double,
      valueHighDouble: 20,
    });
  });

  it('drops empty and invalid numeric filter conditions', () => {
    expect(pivotFilterConditionToSpec('Sales', { kind: 'none', value: '' })).toBeNull();
    expect(
      pivotFilterConditionToSpec('Sales', { kind: 'value-greater-than', value: 'large' }),
    ).toBeNull();
    expect(
      pivotFilterConditionToSpec('Sales', { kind: 'value-between', value: '10..' }),
    ).toBeNull();
  });

  it('converts single-value numeric filter operators into typed payloads', () => {
    expect(pivotFilterConditionToSpec('Sales', { kind: 'value-less-than', value: '100' })).toEqual({
      axis: PivotAxis.Page,
      fieldName: 'Sales',
      type: PivotFilterType.ValueLessThan,
      valueKind: PivotFilterValueKind.Double,
      valueDouble: 100,
    });
    expect(pivotFilterConditionToSpec('Sales', { kind: 'value-equals', value: '100' })).toEqual({
      axis: PivotAxis.Page,
      fieldName: 'Sales',
      type: PivotFilterType.ValueEquals,
      valueKind: PivotFilterValueKind.Double,
      valueDouble: 100,
    });
  });

  it('converts top-count and date filter conditions into typed payloads', () => {
    expect(pivotFilterConditionToSpec('Sales', { kind: 'value-top-10', value: '5' })).toEqual({
      axis: PivotAxis.Page,
      fieldName: 'Sales',
      type: PivotFilterType.ValueTop10,
      valueKind: PivotFilterValueKind.Int,
      valueInt: 5,
    });
    expect(pivotFilterConditionToSpec('Date', { kind: 'label-date', value: '2026-05-19' })).toEqual(
      {
        axis: PivotAxis.Page,
        fieldName: 'Date',
        type: PivotFilterType.LabelDate,
        valueKind: PivotFilterValueKind.Text,
        valueText: '2026-05-19',
      },
    );
    expect(
      pivotFilterConditionToSpec('Date', { kind: 'date-before', value: '2026-05-19' }),
    ).toEqual({
      axis: PivotAxis.Page,
      fieldName: 'Date',
      type: PivotFilterType.DateBefore,
      valueKind: PivotFilterValueKind.Text,
      valueText: '2026-05-19',
    });
    expect(
      pivotFilterConditionToSpec('Date', { kind: 'date-between', value: '2026-05-01..2026-05-31' }),
    ).toEqual({
      axis: PivotAxis.Page,
      fieldName: 'Date',
      type: PivotFilterType.DateBetween,
      valueKind: PivotFilterValueKind.Text,
      valueText: '2026-05-01',
      valueHighKind: PivotFilterValueKind.Text,
      valueHighText: '2026-05-31',
    });
  });

  it('converts PivotFilterSpec payloads back into shared condition state', () => {
    expect(
      pivotFilterSpecToCondition({
        axis: PivotAxis.Page,
        fieldName: 'Sales',
        type: PivotFilterType.ValueBetween,
        valueKind: PivotFilterValueKind.Double,
        valueDouble: -10,
        valueHighKind: PivotFilterValueKind.Double,
        valueHighDouble: 20,
      }),
    ).toEqual({ kind: 'value-between', value: '-10..20' });
    expect(
      pivotFilterSpecToCondition({
        axis: PivotAxis.Page,
        fieldName: 'Region',
        type: PivotFilterType.LabelBeginsWith,
        valueKind: PivotFilterValueKind.Text,
        valueText: 'Ea',
      }),
    ).toEqual({ kind: 'label-begins-with', value: 'Ea' });
    expect(
      pivotFilterSpecToCondition({
        axis: PivotAxis.Page,
        fieldName: 'Region',
        type: PivotFilterType.LabelEndsWith,
        valueKind: PivotFilterValueKind.Text,
        valueText: 'st',
      }),
    ).toEqual({ kind: 'label-ends-with', value: 'st' });
    expect(
      pivotFilterSpecToCondition({
        axis: PivotAxis.Page,
        fieldName: 'Sales',
        type: PivotFilterType.ValueNotBetween,
        valueKind: PivotFilterValueKind.Double,
        valueDouble: 10,
        valueHighKind: PivotFilterValueKind.Double,
        valueHighDouble: 20,
      }),
    ).toEqual({ kind: 'value-not-between', value: '10..20' });
    expect(
      pivotFilterSpecToCondition({
        axis: PivotAxis.Page,
        fieldName: 'Sales',
        type: PivotFilterType.ValueLessThan,
        valueKind: PivotFilterValueKind.Double,
        valueDouble: 100,
      }),
    ).toEqual({ kind: 'value-less-than', value: '100' });
    expect(
      pivotFilterSpecToCondition({
        axis: PivotAxis.Page,
        fieldName: 'Date',
        type: PivotFilterType.DateAfter,
        valueKind: PivotFilterValueKind.Text,
        valueText: '2026-05-19',
      }),
    ).toEqual({ kind: 'date-after', value: '2026-05-19' });
    expect(
      pivotFilterSpecToCondition({
        axis: PivotAxis.Page,
        fieldName: 'Date',
        type: PivotFilterType.DateBetween,
        valueKind: PivotFilterValueKind.Text,
        valueText: '2026-05-01',
        valueHighKind: PivotFilterValueKind.Text,
        valueHighText: '2026-05-31',
      }),
    ).toEqual({ kind: 'date-between', value: '2026-05-01..2026-05-31' });
    expect(
      pivotFilterSpecToCondition({
        axis: PivotAxis.Page,
        fieldName: 'Sales',
        type: PivotFilterType.ValueBetween,
      }),
    ).toBeNull();
  });

  it('keeps restored value and date operators in their Excel filter categories', () => {
    for (const kind of [
      'value-greater-than',
      'value-less-than',
      'value-equals',
      'value-between',
      'value-not-between',
      'value-top-10',
    ] as const) {
      const controls = createPivotFilterConditionControls({
        strings: en.pivotTableDialog,
        condition: { kind, value: kind === 'value-top-10' ? '5' : '10..20' },
        selectClassName: 'select',
        valueClassName: 'input',
        valuesContainerClassName: 'values',
        categoryDataset: { pivotFilterCategory: 'true' },
        conditionDataset: { pivotFilterCondition: 'true' },
        fieldRow: (label, control) => {
          const row = document.createElement('label');
          row.append(label, control);
          return row;
        },
        onChange: () => undefined,
      });
      const host = document.createElement('div');
      host.append(...controls);
      expect(
        host.querySelector<HTMLSelectElement>('select[data-pivot-filter-category="true"]')?.value,
        kind,
      ).toBe('value');
      expect(
        host.querySelector<HTMLSelectElement>('select[data-pivot-filter-condition="true"]')?.value,
        kind,
      ).toBe(kind);
    }

    for (const kind of ['label-date', 'date-before', 'date-after', 'date-between'] as const) {
      const controls = createPivotFilterConditionControls({
        strings: en.pivotTableDialog,
        condition: {
          kind,
          value: kind === 'date-between' ? '2026-05-01..2026-05-31' : '2026-05-19',
        },
        selectClassName: 'select',
        valueClassName: 'input',
        valuesContainerClassName: 'values',
        categoryDataset: { pivotFilterCategory: 'true' },
        conditionDataset: { pivotFilterCondition: 'true' },
        fieldRow: (label, control) => {
          const row = document.createElement('label');
          row.append(label, control);
          return row;
        },
        onChange: () => undefined,
      });
      const host = document.createElement('div');
      host.append(...controls);
      expect(
        host.querySelector<HTMLSelectElement>('select[data-pivot-filter-category="true"]')?.value,
        kind,
      ).toBe('date');
      expect(
        host.querySelector<HTMLSelectElement>('select[data-pivot-filter-condition="true"]')?.value,
        kind,
      ).toBe(kind);
    }
  });

  it('renders shared condition controls and resets condition when category changes', () => {
    const changes: string[] = [];
    const controls = createPivotFilterConditionControls({
      strings: en.pivotTableDialog,
      condition: { kind: 'label-contains', value: 'East' },
      selectClassName: 'select',
      valueClassName: 'input',
      valuesContainerClassName: 'values',
      categoryDataset: { pivotFilterCategory: 'true' },
      conditionDataset: { pivotFilterCondition: 'true' },
      fieldRow: (label, control) => {
        const row = document.createElement('label');
        row.append(label, control);
        return row;
      },
      onChange: (condition) => changes.push(`${condition.kind}:${condition.value}`),
    });
    const host = document.createElement('div');
    host.append(...controls);
    const category = host.querySelector<HTMLSelectElement>(
      'select[data-pivot-filter-category="true"]',
    );
    const condition = host.querySelector<HTMLSelectElement>(
      'select[data-pivot-filter-condition="true"]',
    );
    const value = host.querySelector<HTMLInputElement>('.values input');
    if (!category || !condition || !value) throw new Error('missing condition controls');
    expect(category.className).toBe('select');
    expect(condition.className).toBe('select');
    expect(
      Array.from(category.options).map((option) => [option.value, option.textContent]),
    ).toEqual([
      ['label', 'Label Filters'],
      ['value', 'Value Filters'],
      ['date', 'Date Filters'],
    ]);
    expect(category.value).toBe('label');
    expect(condition.value).toBe('label-contains');
    expect(value.value).toBe('East');

    category.value = 'value';
    category.dispatchEvent(new Event('change', { bubbles: true }));
    expect(condition.value).toBe('none');
    expect(
      Array.from(condition.options)
        .filter((option) => !option.disabled)
        .map((option) => option.value),
    ).toEqual([
      'none',
      'value-greater-than',
      'value-less-than',
      'value-equals',
      'value-between',
      'value-not-between',
      'value-top-10',
    ]);
    expect(
      Array.from(condition.options)
        .filter((option) => option.disabled)
        .map((option) => option.value),
    ).toEqual([]);
    expect(changes.at(-1)).toBe('none:');
  });

  it('enables engine-backed Excel PivotTable filter operators', () => {
    const controls = createPivotFilterConditionControls({
      strings: en.pivotTableDialog,
      condition: { kind: 'label-contains', value: 'East' },
      selectClassName: 'select',
      valueClassName: 'input',
      valuesContainerClassName: 'values',
      categoryDataset: { pivotFilterCategory: 'true' },
      conditionDataset: { pivotFilterCondition: 'true' },
      fieldRow: (label, control) => {
        const row = document.createElement('label');
        row.append(label, control);
        return row;
      },
      onChange: () => undefined,
    });
    const host = document.createElement('div');
    host.append(...controls);
    const category = host.querySelector<HTMLSelectElement>(
      'select[data-pivot-filter-category="true"]',
    );
    const condition = host.querySelector<HTMLSelectElement>(
      'select[data-pivot-filter-condition="true"]',
    );
    if (!category || !condition) throw new Error('missing condition controls');

    expect(
      Array.from(condition.options)
        .filter((option) => !option.disabled)
        .map((option) => option.value),
    ).toEqual([
      'none',
      'label-equals',
      'label-does-not-equal',
      'label-contains',
      'label-does-not-contain',
      'label-begins-with',
      'label-ends-with',
    ]);
    expect(Array.from(condition.options).filter((option) => option.disabled)).toEqual([]);

    category.value = 'date';
    category.dispatchEvent(new Event('change', { bubbles: true }));
    expect(
      Array.from(condition.options)
        .filter((option) => !option.disabled)
        .map((option) => option.value),
    ).toEqual(['none', 'label-date', 'date-before', 'date-after', 'date-between']);
    expect(Array.from(condition.options).filter((option) => option.disabled)).toEqual([]);
  });

  it('opens the shared PivotTable filter dialog and returns the selected condition', async () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const result = showPivotFilterDialog({
      host,
      strings: en.pivotTableDialog,
      fieldName: 'Sales',
      condition: { kind: 'value-greater-than', value: '100' },
      okLabel: 'OK',
      cancelLabel: 'Cancel',
    });
    const dialog = document.body.querySelector<HTMLElement>('.fc-pivotdlg');
    expect(dialog?.textContent).toContain('PivotTable Filter: Sales');
    expect(dialog?.classList.contains('fc-pivotdlg--filter')).toBe(true);
    const category = dialog?.querySelector<HTMLSelectElement>(
      'select[data-pivot-filter-category="true"]',
    );
    const condition = dialog?.querySelector<HTMLSelectElement>(
      'select[data-pivot-filter-condition="true"]',
    );
    if (!category || !condition || !dialog) throw new Error('missing dialog controls');
    await new Promise((resolve) => requestAnimationFrame(resolve));
    expect(document.activeElement).toBe(
      category.closest('.fc-select')?.querySelector<HTMLButtonElement>('.fc-select__button'),
    );
    expect(category.value).toBe('value');
    expect(condition.value).toBe('value-greater-than');
    condition.value = 'value-between';
    condition.dispatchEvent(new Event('change', { bubbles: true }));
    const inputs = Array.from(dialog.querySelectorAll<HTMLInputElement>('input[type="number"]'));
    expect(inputs).toHaveLength(2);
    if (!inputs[0] || !inputs[1]) throw new Error('missing range inputs');
    inputs[0].value = '10';
    inputs[0].dispatchEvent(new Event('input', { bubbles: true }));
    inputs[1].value = '20';
    inputs[1].dispatchEvent(new Event('input', { bubbles: true }));
    dialog.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await expect(result).resolves.toEqual({ kind: 'value-between', value: '10..20' });
    expect(document.body.querySelector('.fc-pivotdlg')).toBeNull();
    host.remove();
  });
});
