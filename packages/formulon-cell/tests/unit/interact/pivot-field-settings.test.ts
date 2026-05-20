import { describe, expect, it } from 'vitest';

import { PivotAxis, PivotFilterType, PivotFilterValueKind } from '../../../src/engine/types.js';
import { en } from '../../../src/i18n/strings.js';
import {
  createPivotFilterConditionControls,
  pivotFilterConditionToSpec,
  pivotFilterSpecToCondition,
  splitFilterConditionRange,
} from '../../../src/interact/pivot-field-settings.js';

describe('pivot-field-settings shared filter condition model', () => {
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
        fieldName: 'Sales',
        type: PivotFilterType.ValueBetween,
      }),
    ).toBeNull();
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
    ).toEqual(['none', 'value-greater-than', 'value-between', 'value-top-10']);
    expect(
      Array.from(condition.options)
        .filter((option) => option.disabled)
        .map((option) => option.value),
    ).toEqual([
      'unsupported-value-less-than',
      'unsupported-value-equals',
      'unsupported-value-not-between',
    ]);
    expect(changes.at(-1)).toBe('none:');
  });

  it('shows unsupported Excel PivotTable filter operators as disabled options', () => {
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

    const disabledLabelOptions = Array.from(condition.options).filter((option) => option.disabled);
    expect(disabledLabelOptions.map((option) => option.value)).toEqual([
      'unsupported-label-equals',
      'unsupported-label-does-not-equal',
      'unsupported-label-does-not-contain',
      'unsupported-label-ends-with',
    ]);
    expect(disabledLabelOptions[0]?.textContent).toContain('requires engine support');

    category.value = 'date';
    category.dispatchEvent(new Event('change', { bubbles: true }));
    expect(
      Array.from(condition.options)
        .filter((option) => option.disabled)
        .map((option) => option.value),
    ).toEqual(['unsupported-date-before', 'unsupported-date-after', 'unsupported-date-between']);
  });
});
