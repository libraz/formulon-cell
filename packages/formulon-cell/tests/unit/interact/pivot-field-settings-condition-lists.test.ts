import { describe, expect, it } from 'vitest';

import { PivotAxis, PivotFilterKind } from '../../../src/engine/types.js';
import { en, ja } from '../../../src/i18n/strings.js';
import {
  createPivotFilterConditionControls,
  type PivotFilterConditionState,
  pivotFilterConditionToSpec,
  pivotFilterSpecToCondition,
} from '../../../src/interact/pivot-field-settings.js';
import {
  PIVOT_FILTER_COMPARISONS,
  PIVOT_FILTER_PERIODS,
} from '../../../src/interact/pivot-filter-kinds.js';

const mount = (
  condition: PivotFilterConditionState | undefined,
  changes: PivotFilterConditionState[] = [],
  strings: typeof en = en,
  dataFields?: readonly string[],
) => {
  const host = document.createElement('div');
  host.append(
    ...createPivotFilterConditionControls({
      strings: strings.pivotTableDialog,
      datePeriods: strings.datePeriods,
      condition,
      ...(dataFields ? { dataFields } : {}),
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
      onChange: (next) => changes.push(next),
    }),
  );
  const q = <T extends HTMLElement>(selector: string): T => {
    const el = host.querySelector<T>(selector);
    if (!el) throw new Error(`missing ${selector}`);
    return el;
  };
  const change = (select: HTMLSelectElement, value: string): void => {
    select.value = value;
    select.dispatchEvent(new Event('change', { bubbles: true }));
  };
  const category = q<HTMLSelectElement>('select[data-pivot-filter-category="true"]');
  const condSelect = q<HTMLSelectElement>('select[data-pivot-filter-condition="true"]');
  return { host, q, change, category, condSelect };
};

describe('Field Settings filter condition lists', () => {
  it('lists the label comparisons in desktop order, all enabled', () => {
    const { condSelect } = mount({ kind: PivotFilterKind.LabelEqual, value: 'x' });
    expect([...condSelect.options].map((o) => [o.value, o.textContent, o.disabled])).toEqual([
      ['', en.pivotTableDialog.filterConditionNone, false],
      ...PIVOT_FILTER_COMPARISONS.label.map(([kind, key]) => [
        String(kind),
        en.pivotTableDialog.filterDialogs.operators[key],
        false,
      ]),
    ]);
  });

  it('lists the value comparisons then Top 10', () => {
    const { condSelect } = mount({ kind: PivotFilterKind.ValueEqual, value: '1' });
    expect([...condSelect.options].slice(1).map((o) => [o.value, o.textContent])).toEqual([
      ...PIVOT_FILTER_COMPARISONS.value.map(([kind, key]) => [
        String(kind),
        en.pivotTableDialog.filterDialogs.operators[key],
      ]),
      [String(PivotFilterKind.ValueTopCount), en.pivotTableDialog.filterDialogs.menu.top10],
    ]);
    expect([...condSelect.options].every((o) => !o.disabled)).toBe(true);
  });

  it.each([
    ['en', en],
    ['ja', ja],
  ] as const)('ends the date list with the 32 periods in order (%s)', (_name, strings) => {
    const { condSelect } = mount({ kind: PivotFilterKind.DateEqual, value: '' }, [], strings);
    const options = [...condSelect.options];
    const enabled = options.filter((o) => !o.disabled);
    expect(options.filter((o) => o.disabled)).toHaveLength(1);
    const comparisons = PIVOT_FILTER_COMPARISONS.date.length;
    expect(enabled.slice(1, 1 + comparisons).map((o) => Number(o.value))).toEqual(
      PIVOT_FILTER_COMPARISONS.date.map(([kind]) => kind),
    );
    const periods = enabled.slice(1 + comparisons);
    expect(periods).toHaveLength(32);
    expect(periods.map((o) => [Number(o.value), o.textContent])).toEqual(
      PIVOT_FILTER_PERIODS.map(([kind, key]) => [kind, strings.datePeriods[key]]),
    );
    expect(options[1 + comparisons]?.disabled).toBe(true);
  });

  it('renders no value input for a period kind', () => {
    const { host, condSelect, change } = mount({ kind: PivotFilterKind.DateEqual, value: 'x' });
    change(condSelect, String(PivotFilterKind.Today));
    expect(host.querySelector('.values input')).toBeNull();
  });

  it('shows a Top 10 kind as the Top 10 entry with its parts and maps edits back', () => {
    const changes: PivotFilterConditionState[] = [];
    const { q, change, condSelect } = mount(
      { kind: PivotFilterKind.ValueBottomPercent, value: '20' },
      changes,
    );
    expect(condSelect.value).toBe(String(PivotFilterKind.ValueTopCount));
    const direction = q<HTMLSelectElement>('select[data-pivot-filter-top10-direction="true"]');
    const basis = q<HTMLSelectElement>('select[data-pivot-filter-top10-basis="true"]');
    const count = q<HTMLInputElement>('input[data-pivot-filter-top10-count="true"]');
    expect([direction.value, basis.value, count.value]).toEqual(['bottom', 'percent', '20']);
    change(direction, 'top');
    change(basis, 'sum');
    expect(changes.at(-1)).toEqual({ kind: PivotFilterKind.ValueTopSum, value: '20' });
    count.value = '3';
    count.dispatchEvent(new Event('input', { bubbles: true }));
    expect(changes.at(-1)).toEqual({ kind: PivotFilterKind.ValueTopSum, value: '3' });
  });

  it('starts a chosen Top 10 entry at Top / 10 / Items', () => {
    const changes: PivotFilterConditionState[] = [];
    const { change, category, condSelect } = mount(undefined, changes);
    change(category, 'value');
    change(condSelect, String(PivotFilterKind.ValueTopCount));
    expect(changes.at(-1)).toEqual({ kind: PivotFilterKind.ValueTopCount, value: '10' });
  });

  it('carries the scored data field on a value condition', () => {
    const changes: PivotFilterConditionState[] = [];
    const { q, change } = mount(
      { kind: PivotFilterKind.ValueGreaterThan, value: '5' },
      changes,
      en,
      ['A', 'B'],
    );
    change(q<HTMLSelectElement>('select[data-pivot-filter-data-field="true"]'), '1');
    expect(changes.at(-1)).toEqual({
      kind: PivotFilterKind.ValueGreaterThan,
      value: '5',
      dataField: 1,
    });
  });
});

describe('filter condition data field conversion', () => {
  it('writes dataFieldIndex only above the first data field', () => {
    const base = { kind: PivotFilterKind.ValueGreaterThan, value: '5' };
    expect(pivotFilterConditionToSpec('Sales', { ...base, dataField: 2 }, PivotAxis.Row)).toEqual({
      axis: PivotAxis.Row,
      fieldName: 'Sales',
      kind: PivotFilterKind.ValueGreaterThan,
      value: 5,
      dataFieldIndex: 2,
    });
    expect(pivotFilterConditionToSpec('Sales', { ...base, dataField: 0 })).not.toHaveProperty(
      'dataFieldIndex',
    );
  });

  it('reads dataFieldIndex back into the condition', () => {
    expect(
      pivotFilterSpecToCondition({
        axis: PivotAxis.Row,
        fieldName: 'Sales',
        kind: PivotFilterKind.ValueTopSum,
        value: 3,
        dataFieldIndex: 1,
      }),
    ).toEqual({ kind: PivotFilterKind.ValueTopSum, value: '3', dataField: 1 });
  });
});
