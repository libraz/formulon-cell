import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { en } from '../../../../src/i18n/strings.js';
import { attachConditionalDialog } from '../../../../src/interact/conditional-dialog.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../../src/store/store.js';
import { setRange } from './fixtures.js';

describe('attachConditionalDialog', () => {
  let host: HTMLElement;
  let store: SpreadsheetStore;

  beforeEach(() => {
    host = document.createElement('div');
    host.tabIndex = -1;
    document.body.appendChild(host);
    store = createSpreadsheetStore();
  });

  afterEach(() => {
    document.body.innerHTML = '';
  });

  it('adds color-scale rules with Excel-style threshold metadata from the classic dialog', () => {
    setRange(store, 0, 1, 5, 1);
    const handle = attachConditionalDialog({ host, store, strings: en });
    handle.open({ kind: 'color-scale' });

    const selects = Array.from(
      document.querySelectorAll<HTMLSelectElement>('.fc-conddlg__form select'),
    );
    const scaleTypeSelects = selects.filter((select) =>
      Array.from(select.options).some((option) => option.value === 'min'),
    );
    const minType = scaleTypeSelects[0] as HTMLSelectElement;
    const maxType = scaleTypeSelects[2] as HTMLSelectElement;
    const numberInputs = Array.from(
      document.querySelectorAll<HTMLInputElement>('.fc-conddlg__form input[type="number"]'),
    );
    const minValue = numberInputs.find(
      (input) => input.getAttribute('aria-label') === 'Min Value',
    ) as HTMLInputElement;

    minType.value = 'number';
    minType.dispatchEvent(new Event('change', { bubbles: true }));
    minValue.value = '10';
    maxType.value = 'percent';
    maxType.dispatchEvent(new Event('change', { bubbles: true }));
    const maxValue = numberInputs.find(
      (input) => input.getAttribute('aria-label') === 'Max Value',
    ) as HTMLInputElement;
    maxValue.value = '90';

    document
      .querySelector<HTMLButtonElement>('.fc-conddlg__addrow .fc-fmtdlg__btn--primary')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    expect(store.getState().conditional.rules[0]).toMatchObject({
      kind: 'color-scale',
      range: { sheet: 0, r0: 0, c0: 1, r1: 5, c1: 1 },
      thresholds: [
        { kind: 'number', value: 10 },
        { kind: 'percent', value: 90 },
      ],
    });
    handle.detach();
  });

  it('adds icon-set rules with icon-only and reverse-order options from the classic dialog', () => {
    setRange(store, 0, 0, 4, 0);
    const handle = attachConditionalDialog({ host, store, strings: en });
    handle.open({ kind: 'icon-set' });

    const iconSelect = Array.from(
      document.querySelectorAll<HTMLSelectElement>('.fc-conddlg__form select'),
    ).find((select) =>
      Array.from(select.options).some((option) => option.value === 'traffic3'),
    ) as HTMLSelectElement;
    iconSelect.value = 'traffic3';
    iconSelect.dispatchEvent(new Event('change', { bubbles: true }));

    const checks = Array.from(
      document.querySelectorAll<HTMLInputElement>('.fc-conddlg__sub input[type="checkbox"]'),
    );
    const reverse = checks.find(
      (input) => input.nextElementSibling?.textContent === 'Reverse order',
    );
    const iconOnly = checks.find(
      (input) => input.nextElementSibling?.textContent === 'Show icon only',
    );
    if (!reverse || !iconOnly) throw new Error('missing icon-set checkboxes');
    reverse.checked = true;
    iconOnly.checked = true;
    const thresholdValues = Array.from(
      document.querySelectorAll<HTMLInputElement>(
        '.fc-conddlg__sub input[aria-label^="Threshold"]',
      ),
    ).filter((input) => !input.closest('label')?.hidden);
    expect(thresholdValues).toHaveLength(2);
    const [firstThreshold, secondThreshold] = thresholdValues;
    if (!firstThreshold || !secondThreshold) throw new Error('missing icon-set thresholds');
    firstThreshold.value = '25';
    secondThreshold.value = '75';

    document
      .querySelector<HTMLButtonElement>('.fc-conddlg__addrow .fc-fmtdlg__btn--primary')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    expect(store.getState().conditional.rules[0]).toMatchObject({
      kind: 'icon-set',
      range: { sheet: 0, r0: 0, c0: 0, r1: 4, c1: 0 },
      icons: 'traffic3',
      reverseOrder: true,
      showValue: false,
      thresholds: [
        { kind: 'percent', value: 25 },
        { kind: 'percent', value: 75 },
      ],
    });
    handle.detach();
  });

  it('keeps mixed icon threshold operators when editing a hydrated rule', () => {
    const rule = {
      kind: 'icon-set' as const,
      range: { sheet: 0, r0: 0, c0: 0, r1: 4, c1: 0 },
      icons: 'traffic3' as const,
      thresholds: [
        { kind: 'percent' as const, value: 33, gte: false },
        { kind: 'percent' as const, value: 67, gte: true },
      ],
      floor: { kind: 'percent' as const, value: 10, gte: false },
      showValue: true,
    };
    mutators.addConditionalRule(store, rule);
    const handle = attachConditionalDialog({ host, store, strings: en });
    handle.open({ mode: 'edit', editIndex: 0 });

    const [firstOperator, secondOperator] = [
      document.querySelector<HTMLSelectElement>('select[data-cf-icon-operator="0"]'),
      document.querySelector<HTMLSelectElement>('select[data-cf-icon-operator="1"]'),
    ];
    if (!firstOperator || !secondOperator) throw new Error('missing icon threshold operators');
    expect([firstOperator, secondOperator].map((select) => select.value)).toEqual(['>', '>=']);

    document
      .querySelector<HTMLButtonElement>('.fc-conddlg__addrow .fc-fmtdlg__btn--primary')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    const saved = store.getState().conditional.rules[0];
    expect(saved).toMatchObject({
      kind: 'icon-set',
      thresholds: [
        { kind: 'percent', value: 33, gte: false },
        { kind: 'percent', value: 67 },
      ],
      floor: { kind: 'percent', value: 10, gte: false },
    });
    if (saved?.kind === 'icon-set') expect(saved.thresholds?.[1]?.gte).not.toBe(false);
    handle.detach();
  });

  it('collects a strict icon threshold while defaulting each boundary to greater-than-or-equal', () => {
    setRange(store, 0, 0, 4, 0);
    const handle = attachConditionalDialog({ host, store, strings: en });
    handle.open({ kind: 'icon-set' });

    const firstOperator = document.querySelector<HTMLSelectElement>(
      'select[data-cf-icon-operator="0"]',
    );
    const secondOperator = document.querySelector<HTMLSelectElement>(
      'select[data-cf-icon-operator="1"]',
    );
    if (!firstOperator || !secondOperator) throw new Error('missing icon threshold operators');
    expect(firstOperator.value).toBe('>=');
    expect(secondOperator.value).toBe('>=');
    firstOperator.value = '>';

    document
      .querySelector<HTMLButtonElement>('.fc-conddlg__addrow .fc-fmtdlg__btn--primary')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    const saved = store.getState().conditional.rules[0];
    expect(saved).toMatchObject({
      kind: 'icon-set',
      thresholds: [{ gte: false }, {}],
    });
    handle.detach();
  });

  it('adds above/below average rules from the classic dialog', () => {
    setRange(store, 2, 1, 5, 1);
    const handle = attachConditionalDialog({ host, store });
    handle.open({ kind: 'average', averageMode: 'equal-or-above' });

    const kindSelect = Array.from(
      document.querySelectorAll<HTMLSelectElement>('.fc-conddlg__form select'),
    ).find((select) =>
      Array.from(select.options).some((option) => option.value === 'average'),
    ) as HTMLSelectElement;
    expect(kindSelect.value).toBe('average');

    const averageSelect = Array.from(
      document.querySelectorAll<HTMLSelectElement>('.fc-conddlg__form select'),
    ).find((select) =>
      Array.from(select.options).some((option) => option.value === 'equal-or-above'),
    ) as HTMLSelectElement;
    expect(averageSelect.value).toBe('equal-or-above');

    document
      .querySelector<HTMLButtonElement>('.fc-conddlg__addrow .fc-fmtdlg__btn--primary')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    expect(store.getState().conditional.rules[0]).toMatchObject({
      kind: 'average',
      mode: 'equal-or-above',
      range: { sheet: 0, r0: 2, c0: 1, r1: 5, c1: 1 },
      apply: { fill: '#ffc7ce', color: '#9c0006' },
    });
    handle.detach();
  });

  it('adds standard-deviation average rules from the classic dialog', () => {
    setRange(store, 2, 1, 5, 1);
    const handle = attachConditionalDialog({ host, store });
    handle.open({ kind: 'average', averageMode: 'above-std-dev' });

    const averageSelect = Array.from(
      document.querySelectorAll<HTMLSelectElement>('.fc-conddlg__form select'),
    ).find((select) =>
      Array.from(select.options).some((option) => option.value === 'above-std-dev'),
    ) as HTMLSelectElement;
    expect(averageSelect.value).toBe('above-std-dev');
    const stdDevSelect = Array.from(
      document.querySelectorAll<HTMLSelectElement>('.fc-conddlg__form select'),
    ).find((select) => Array.from(select.options).some((option) => option.value === '3')) as
      | HTMLSelectElement
      | undefined;
    expect(stdDevSelect?.parentElement?.hidden).toBe(false);
    if (stdDevSelect) stdDevSelect.value = '2';

    document
      .querySelector<HTMLButtonElement>('.fc-conddlg__addrow .fc-fmtdlg__btn--primary')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    expect(store.getState().conditional.rules[0]).toMatchObject({
      kind: 'average',
      mode: 'above-std-dev',
      stdDev: 2,
      range: { sheet: 0, r0: 2, c0: 1, r1: 5, c1: 1 },
    });
    handle.detach();
  });

  it('adds text and date-occurring rules from the classic dialog', () => {
    setRange(store, 1, 1, 2, 2);
    const handle = attachConditionalDialog({ host, store });
    handle.open({ kind: 'text-contains', text: 'due' });
    const textModeSelect = Array.from(
      document.querySelectorAll<HTMLSelectElement>('.fc-conddlg__form select'),
    ).find((select) =>
      Array.from(select.options).some((option) => option.value === 'not-contains'),
    );
    if (textModeSelect) textModeSelect.value = 'begins-with';

    document
      .querySelector<HTMLButtonElement>('.fc-conddlg__addrow .fc-fmtdlg__btn--primary')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(store.getState().conditional.rules[0]).toMatchObject({
      kind: 'text-contains',
      range: { sheet: 0, r0: 1, c0: 1, r1: 2, c1: 2 },
      text: 'due',
      mode: 'begins-with',
    });

    handle.open({ kind: 'date-occurring', datePeriod: 'last7' });
    document
      .querySelector<HTMLButtonElement>('.fc-conddlg__addrow .fc-fmtdlg__btn--primary')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(store.getState().conditional.rules[1]).toMatchObject({
      kind: 'date-occurring',
      period: 'last7',
    });
    handle.detach();
  });
});
