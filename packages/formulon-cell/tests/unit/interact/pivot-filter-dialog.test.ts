import { afterEach, describe, expect, it } from 'vitest';

import { PivotFilterKind } from '../../../src/engine/types.js';
import { en, ja } from '../../../src/i18n/strings.js';
import type { PivotFilterConditionState } from '../../../src/interact/pivot-field-settings.js';
import { showPivotFilterDialog } from '../../../src/interact/pivot-filter-dialog.js';
import {
  PIVOT_FILTER_COMPARISONS,
  type PivotTop10Basis,
  type PivotTop10Direction,
} from '../../../src/interact/pivot-filter-kinds.js';

type Locale = typeof en;

const open = (
  strings: Locale,
  condition: PivotFilterConditionState | undefined,
  extra: { dataFields?: readonly string[]; mac?: boolean } = {},
) => {
  const host = document.createElement('div');
  document.body.appendChild(host);
  const result = showPivotFilterDialog({
    host,
    strings: strings.pivotTableDialog,
    fieldName: 'Region',
    condition,
    okLabel: 'OK',
    cancelLabel: 'Cancel',
    mac: extra.mac ?? false,
    ...(extra.dataFields ? { dataFields: extra.dataFields } : {}),
  });
  const dialog = document.body.querySelector<HTMLElement>('.fc-pivotdlg') as HTMLElement;
  const q = <T extends HTMLElement>(selector: string): T | null =>
    dialog.querySelector<T>(selector);
  const ok = (): void => q<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
  const change = (select: HTMLSelectElement, value: string): void => {
    select.value = value;
    select.dispatchEvent(new Event('change', { bubbles: true }));
  };
  return { result, dialog, q, ok, change };
};

afterEach(() => {
  document.body.replaceChildren();
});

describe('pivot filter dialogs', () => {
  it('titles each dialog with the field on Windows and without it on Mac', () => {
    const win = open(en, { kind: PivotFilterKind.LabelEqual, value: '' });
    expect(win.dialog.textContent).toContain('Label Filter (Region)');
    document.body.replaceChildren();
    const mac = open(en, { kind: PivotFilterKind.LabelEqual, value: '' }, { mac: true });
    expect(mac.dialog.textContent).toContain('Label Filter');
    expect(mac.dialog.textContent).not.toContain('Region');
    document.body.replaceChildren();
    expect(open(en, { kind: PivotFilterKind.ValueEqual, value: '' }).dialog.textContent).toContain(
      'Value Filter (Region)',
    );
    document.body.replaceChildren();
    expect(open(en, { kind: PivotFilterKind.DateEqual, value: '' }).dialog.textContent).toContain(
      'Date Filter (Region)',
    );
    document.body.replaceChildren();
    expect(
      open(en, { kind: PivotFilterKind.ValueTopCount, value: '' }).dialog.textContent,
    ).toContain('Top 10 Filter (Region)');
  });

  it.each([
    ['en', en],
    ['ja', ja],
  ] as const)('lists every comparison of each dialog in desktop order (%s)', (_name, strings) => {
    const d = strings.pivotTableDialog.filterDialogs;
    for (const category of ['label', 'value', 'date'] as const) {
      const first = PIVOT_FILTER_COMPARISONS[category][0]?.[0] as PivotFilterKind;
      const { q, dialog } = open(strings, { kind: first, value: '' });
      const operator = q<HTMLSelectElement>('select[data-pivot-filter-operator="true"]');
      expect([...(operator?.options ?? [])].map((o) => [Number(o.value), o.textContent])).toEqual(
        PIVOT_FILTER_COMPARISONS[category].map(([kind, key]) => [kind, d.operators[key]]),
      );
      expect(dialog.textContent).toContain(d[category].prompt);
      expect(dialog.textContent).toContain(d[category].macTitle === '' ? '' : d[category].and);
      document.body.replaceChildren();
    }
  });

  it('shows the wildcard hints in the label dialog only', () => {
    const d = ja.pivotTableDialog.filterDialogs;
    const label = open(ja, { kind: PivotFilterKind.LabelContains, value: 'x' });
    expect(label.dialog.textContent).toContain(d.label.wildcardOne);
    expect(label.dialog.textContent).toContain(d.label.wildcardMany);
    document.body.replaceChildren();
    const value = open(ja, { kind: PivotFilterKind.ValueEqual, value: '1' });
    expect(value.dialog.textContent).not.toContain(d.label.wildcardOne);
  });

  it('shows the second field only for range operators and returns low..high', async () => {
    const { q, change, ok, result } = open(en, { kind: PivotFilterKind.LabelEqual, value: 'A' });
    const operator = q<HTMLSelectElement>('select[data-pivot-filter-operator="true"]');
    const high = q<HTMLInputElement>('input[data-pivot-filter-value-high="true"]');
    const low = q<HTMLInputElement>('input[data-pivot-filter-value="true"]');
    if (!operator || !high || !low) throw new Error('missing controls');
    expect(high.closest('label')?.hidden).toBe(true);
    change(operator, String(PivotFilterKind.LabelNotBetween));
    expect(high.closest('label')?.hidden).toBe(false);
    high.value = 'M';
    ok();
    await expect(result).resolves.toEqual({
      kind: PivotFilterKind.LabelNotBetween,
      value: 'A..M',
    });
  });

  it('restores a stored range into both fields', () => {
    const { q } = open(en, { kind: PivotFilterKind.DateBetween, value: '2026-05-01..2026-05-31' });
    expect(q<HTMLInputElement>('input[data-pivot-filter-value="true"]')?.value).toBe('2026-05-01');
    expect(q<HTMLInputElement>('input[data-pivot-filter-value-high="true"]')?.value).toBe(
      '2026-05-31',
    );
  });

  it('scores a data field in the value dialog', async () => {
    const { q, change, ok, result } = open(
      en,
      { kind: PivotFilterKind.ValueGreaterThan, value: '5' },
      { dataFields: ['Sum of Sales', 'Count of Sales'] },
    );
    const field = q<HTMLSelectElement>('select[data-pivot-filter-data-field="true"]');
    if (!field) throw new Error('missing data field select');
    expect([...field.options].map((o) => o.textContent)).toEqual([
      'Sum of Sales',
      'Count of Sales',
    ]);
    change(field, '1');
    ok();
    await expect(result).resolves.toEqual({
      kind: PivotFilterKind.ValueGreaterThan,
      value: '5',
      dataField: 1,
    });
  });

  it('offers no data field select without data fields', () => {
    const { q } = open(en, { kind: PivotFilterKind.ValueGreaterThan, value: '5' });
    expect(q('select[data-pivot-filter-data-field="true"]')).toBeNull();
  });

  it('maps every Top 10 direction and basis to its kind and carries the data field', async () => {
    const expected: [PivotTop10Direction, PivotTop10Basis, PivotFilterKind][] = [
      ['top', 'items', 0],
      ['bottom', 'items', 24],
      ['top', 'percent', 25],
      ['bottom', 'percent', 26],
      ['top', 'sum', 27],
      ['bottom', 'sum', 28],
    ];
    for (const [direction, basis, kind] of expected) {
      const { q, change, ok, result } = open(
        en,
        { kind: PivotFilterKind.ValueTopCount, value: '' },
        { dataFields: ['A', 'B'] },
      );
      const dir = q<HTMLSelectElement>('select[data-pivot-filter-top10-direction="true"]');
      const bas = q<HTMLSelectElement>('select[data-pivot-filter-top10-basis="true"]');
      const count = q<HTMLInputElement>('input[data-pivot-filter-top10-count="true"]');
      const by = q<HTMLSelectElement>('select[data-pivot-filter-data-field="true"]');
      if (!dir || !bas || !count || !by) throw new Error('missing top 10 controls');
      expect(count.value).toBe('10');
      expect([count.min, count.step]).toEqual(['1', '1']);
      change(dir, direction);
      change(bas, basis);
      change(by, '1');
      count.value = '7';
      ok();
      await expect(result).resolves.toEqual({ kind, value: '7', dataField: 1 });
      document.body.replaceChildren();
    }
  });

  it('labels the Top 10 dialog in each locale and preselects the stored parts', () => {
    for (const strings of [en, ja]) {
      const d = strings.pivotTableDialog.filterDialogs.top10;
      const { dialog, q } = open(strings, {
        kind: PivotFilterKind.ValueBottomPercent,
        value: '20',
      });
      for (const text of [d.show, d.top, d.bottom, d.items, d.percent, d.sum]) {
        expect(dialog.textContent).toContain(text);
      }
      expect(q<HTMLSelectElement>('select[data-pivot-filter-top10-direction="true"]')?.value).toBe(
        'bottom',
      );
      expect(q<HTMLSelectElement>('select[data-pivot-filter-top10-basis="true"]')?.value).toBe(
        'percent',
      );
      expect(q<HTMLInputElement>('input[data-pivot-filter-top10-count="true"]')?.value).toBe('20');
      document.body.replaceChildren();
    }
  });

  it('resolves null when cancelled', async () => {
    const { q, result } = open(en, { kind: PivotFilterKind.LabelEqual, value: '' });
    q<HTMLButtonElement>('.fc-fmtdlg__btn:not(.fc-fmtdlg__btn--primary)')?.click();
    await expect(result).resolves.toBeNull();
  });
});
