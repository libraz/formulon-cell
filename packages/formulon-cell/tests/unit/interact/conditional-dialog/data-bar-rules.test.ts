import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { en, ja } from '../../../../src/i18n/strings.js';
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

  it('adds data-bar rules with gradient or solid fill style from the classic dialog', () => {
    setRange(store, 1, 1, 4, 1);
    const handle = attachConditionalDialog({ host, store, strings: en });
    handle.open({ kind: 'data-bar' });

    const selects = Array.from(
      document.querySelectorAll<HTMLSelectElement>('.fc-conddlg__form select'),
    );
    const fillStyleSelect = selects.find((select) =>
      Array.from(select.options).some((option) => option.value === 'gradient'),
    ) as HTMLSelectElement;
    expect(fillStyleSelect.value).toBe('gradient');

    document
      .querySelector<HTMLButtonElement>('.fc-conddlg__addrow .fc-fmtdlg__btn--primary')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(store.getState().conditional.rules[0]).toMatchObject({
      kind: 'data-bar',
      range: { sheet: 0, r0: 1, c0: 1, r1: 4, c1: 1 },
      color: '#638ec6',
      min: { kind: 'min' },
      max: { kind: 'max' },
      direction: 'context',
      gradient: true,
      showValue: true,
    });

    fillStyleSelect.value = 'solid';
    fillStyleSelect.dispatchEvent(new Event('change', { bubbles: true }));
    document
      .querySelector<HTMLButtonElement>('.fc-conddlg__addrow .fc-fmtdlg__btn--primary')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(store.getState().conditional.rules[1]).toMatchObject({
      kind: 'data-bar',
      gradient: false,
    });
    handle.detach();
  });

  it('adds data-bar bounds and direction with localized labelled controls', () => {
    setRange(store, 1, 1, 4, 1);
    const handle = attachConditionalDialog({ host, store, strings: en });
    handle.open({ kind: 'data-bar' });

    const minType = document.querySelector<HTMLSelectElement>('select[data-cf-bar-min-type]');
    const minValue = document.querySelector<HTMLInputElement>('input[data-cf-bar-min-value]');
    const maxType = document.querySelector<HTMLSelectElement>('select[data-cf-bar-max-type]');
    const maxValue = document.querySelector<HTMLInputElement>('input[data-cf-bar-max-value]');
    const direction = document.querySelector<HTMLSelectElement>('select[data-cf-bar-direction]');
    if (!minType || !minValue || !maxType || !maxValue || !direction) {
      throw new Error('missing data-bar controls');
    }

    expect(minType.value).toBe('min');
    expect(maxType.value).toBe('max');
    expect(direction.value).toBe('context');
    expect(Array.from(minType.options, (option) => option.value)).toEqual([
      'min',
      'number',
      'percent',
      'percentile',
    ]);
    expect(Array.from(maxType.options, (option) => option.value)).toEqual([
      'max',
      'number',
      'percent',
      'percentile',
    ]);
    expect(minType.getAttribute('aria-label')).toBe('Minimum Type');
    expect(minValue.getAttribute('aria-label')).toBe('Minimum Value');
    expect(direction.getAttribute('aria-label')).toBe('Direction');

    minType.value = 'number';
    minType.dispatchEvent(new Event('change', { bubbles: true }));
    minValue.value = '10';
    maxType.value = 'percent';
    maxType.dispatchEvent(new Event('change', { bubbles: true }));
    maxValue.value = '90';
    direction.value = 'right-to-left';

    document
      .querySelector<HTMLButtonElement>('.fc-conddlg__addrow .fc-fmtdlg__btn--primary')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    expect(store.getState().conditional.rules[0]).toMatchObject({
      kind: 'data-bar',
      min: { kind: 'number', value: 10 },
      max: { kind: 'percent', value: 90 },
      direction: 'right-to-left',
    });

    handle.detach();

    document.body.innerHTML = '';
    document.body.appendChild(host);
    const jaHandle = attachConditionalDialog({ host, store, strings: ja });
    jaHandle.open({ kind: 'data-bar' });
    expect(
      document
        .querySelector<HTMLSelectElement>('select[data-cf-bar-min-type]')
        ?.getAttribute('aria-label'),
    ).toBe('最小値 種類');
    expect(
      document
        .querySelector<HTMLSelectElement>('select[data-cf-bar-direction]')
        ?.getAttribute('aria-label'),
    ).toBe('方向');
    jaHandle.open({ kind: 'icon-set' });
    expect(
      document
        .querySelector<HTMLSelectElement>('select[data-cf-icon-operator="0"]')
        ?.getAttribute('aria-label'),
    ).toBe('しきい値 1 演算子');
    jaHandle.detach();
  });

  it('exposes Excel data-bar appearance defaults and keeps disabled values reversible', () => {
    const handle = attachConditionalDialog({ host, store, strings: en });
    handle.open({ mode: 'new', kind: 'data-bar' });

    const axisPosition = document.querySelector<HTMLSelectElement>(
      'select[data-cf-bar-axis-position]',
    );
    const positiveColor = document.querySelector<HTMLInputElement>(
      'input[data-cf-bar-positive-color]',
    );
    const negativeColor = document.querySelector<HTMLInputElement>(
      'input[data-cf-bar-negative-color]',
    );
    const borderStyle = document.querySelector<HTMLSelectElement>(
      'select[data-cf-bar-border-style]',
    );
    const borderColor = document.querySelector<HTMLInputElement>('input[data-cf-bar-border-color]');
    const negativeBorderColor = document.querySelector<HTMLInputElement>(
      'input[data-cf-bar-negative-border-color]',
    );
    const axisColor = document.querySelector<HTMLInputElement>('input[data-cf-bar-axis-color]');
    if (
      !axisPosition ||
      !positiveColor ||
      !negativeColor ||
      !borderStyle ||
      !borderColor ||
      !negativeBorderColor ||
      !axisColor
    ) {
      throw new Error('missing data-bar appearance controls');
    }

    expect(axisPosition.value).toBe('automatic');
    expect(Array.from(axisPosition.options, (option) => option.value)).toEqual([
      'automatic',
      'middle',
      'none',
    ]);
    expect(positiveColor.value).toBe('#638ec6');
    expect(negativeColor.value).toBe('#ff0000');
    expect(borderStyle.value).toBe('none');
    expect(borderColor.disabled).toBe(true);
    expect(negativeBorderColor.disabled).toBe(true);
    expect(axisColor.value).toBe('#000000');
    expect(axisColor.disabled).toBe(false);
    expect(axisPosition.getAttribute('aria-label')).toBe('Axis position');
    expect(negativeColor.getAttribute('aria-label')).toBe('Negative value');
    expect(borderStyle.getAttribute('aria-label')).toBe('Border');

    axisColor.value = '#006600';
    axisColor.dispatchEvent(new Event('input', { bubbles: true }));
    axisPosition.value = 'none';
    axisPosition.dispatchEvent(new Event('change', { bubbles: true }));
    expect(axisColor.disabled).toBe(true);
    expect(axisColor.value).toBe('#006600');
    axisPosition.value = 'middle';
    axisPosition.dispatchEvent(new Event('change', { bubbles: true }));
    expect(axisColor.disabled).toBe(false);
    expect(axisColor.value).toBe('#006600');

    borderStyle.value = 'solid';
    borderStyle.dispatchEvent(new Event('change', { bubbles: true }));
    expect(borderColor.disabled).toBe(false);
    expect(negativeBorderColor.disabled).toBe(false);
    handle.detach();

    document.body.innerHTML = '';
    document.body.appendChild(host);
    const jaHandle = attachConditionalDialog({ host, store, strings: ja });
    jaHandle.open({ mode: 'new', kind: 'data-bar' });
    expect(
      document
        .querySelector<HTMLSelectElement>('select[data-cf-bar-axis-position]')
        ?.getAttribute('aria-label'),
    ).toBe('軸の位置');
    expect(
      document
        .querySelector<HTMLInputElement>('input[data-cf-bar-negative-color]')
        ?.getAttribute('aria-label'),
    ).toBe('負の値');
    expect(
      document
        .querySelector<HTMLSelectElement>('select[data-cf-bar-border-style]')
        ?.getAttribute('aria-label'),
    ).toBe('罫線');
    jaHandle.detach();
  });

  it('collects data-bar appearance settings and resets an abandoned draft', () => {
    const handle = attachConditionalDialog({ host, store, strings: en });
    handle.open({ mode: 'new', kind: 'data-bar' });

    const axisPosition = document.querySelector<HTMLSelectElement>(
      'select[data-cf-bar-axis-position]',
    );
    const negativeColor = document.querySelector<HTMLInputElement>(
      'input[data-cf-bar-negative-color]',
    );
    const borderStyle = document.querySelector<HTMLSelectElement>(
      'select[data-cf-bar-border-style]',
    );
    const borderColor = document.querySelector<HTMLInputElement>('input[data-cf-bar-border-color]');
    const negativeBorderColor = document.querySelector<HTMLInputElement>(
      'input[data-cf-bar-negative-border-color]',
    );
    const axisColor = document.querySelector<HTMLInputElement>('input[data-cf-bar-axis-color]');
    if (
      !axisPosition ||
      !negativeColor ||
      !borderStyle ||
      !borderColor ||
      !negativeBorderColor ||
      !axisColor
    ) {
      throw new Error('missing data-bar appearance controls');
    }
    axisPosition.value = 'middle';
    negativeColor.value = '#cc0000';
    negativeColor.dispatchEvent(new Event('input', { bubbles: true }));
    borderStyle.value = 'solid';
    borderStyle.dispatchEvent(new Event('change', { bubbles: true }));
    borderColor.value = '#003366';
    borderColor.dispatchEvent(new Event('input', { bubbles: true }));
    negativeBorderColor.value = '#990000';
    negativeBorderColor.dispatchEvent(new Event('input', { bubbles: true }));
    axisColor.value = '#006600';
    axisColor.dispatchEvent(new Event('input', { bubbles: true }));
    axisPosition.value = 'none';
    axisPosition.dispatchEvent(new Event('change', { bubbles: true }));

    document
      .querySelector<HTMLButtonElement>('.fc-conddlg__addrow .fc-fmtdlg__btn--primary')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(store.getState().conditional.rules[0]).toMatchObject({
      kind: 'data-bar',
      axisPosition: 'none',
      negativeColor: '#cc0000',
      borderColor: '#003366',
      negativeBorderColor: '#990000',
      axisColor: '#006600',
    });

    handle.open({ mode: 'new', kind: 'data-bar' });
    expect(
      document.querySelector<HTMLSelectElement>('select[data-cf-bar-axis-position]')?.value,
    ).toBe('automatic');
    expect(
      document.querySelector<HTMLInputElement>('input[data-cf-bar-negative-color]')?.value,
    ).toBe('#ff0000');
    expect(
      document.querySelector<HTMLSelectElement>('select[data-cf-bar-border-style]')?.value,
    ).toBe('none');
    expect(document.querySelector<HTMLInputElement>('input[data-cf-bar-axis-color]')?.value).toBe(
      '#000000',
    );
    const reopenedAxis = document.querySelector<HTMLSelectElement>(
      'select[data-cf-bar-axis-position]',
    );
    if (!reopenedAxis) throw new Error('missing reopened axis position');
    reopenedAxis.value = 'middle';
    reopenedAxis.dispatchEvent(new Event('change', { bubbles: true }));
    document
      .querySelector<HTMLButtonElement>('.fc-conddlg__footer .fc-fmtdlg__btn')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(store.getState().conditional.rules[0]).toMatchObject({
      axisPosition: 'none',
      negativeColor: '#cc0000',
    });

    handle.open({ mode: 'new', kind: 'data-bar' });
    expect(
      document.querySelector<HTMLSelectElement>('select[data-cf-bar-axis-position]')?.value,
    ).toBe('automatic');
    expect(
      document.querySelector<HTMLInputElement>('input[data-cf-bar-negative-color]')?.value,
    ).toBe('#ff0000');
    expect(
      document.querySelector<HTMLSelectElement>('select[data-cf-bar-border-style]')?.value,
    ).toBe('none');
    expect(document.querySelector<HTMLInputElement>('input[data-cf-bar-axis-color]')?.value).toBe(
      '#000000',
    );
    handle.detach();
  });

  it('hydrates CSS colors for data bars without losing alpha or absent border sides', () => {
    mutators.addConditionalRule(store, {
      kind: 'data-bar',
      range: { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 0 },
      color: 'rgba(99, 142, 198, 0.5)',
      negativeColor: 'rgb(255, 0, 0)',
      negativeBorderColor: 'rgba(153, 0, 0, 0.75)',
      axisColor: 'rgba(0, 0, 0, 0.25)',
      axisPosition: 'none',
      gradient: true,
      showValue: true,
    });
    const handle = attachConditionalDialog({ host, store, strings: en });
    handle.open({ mode: 'edit', editIndex: 0 });

    expect(
      document.querySelector<HTMLInputElement>('input[data-cf-bar-positive-color]')?.value,
    ).toBe('#638ec6');
    expect(
      document.querySelector<HTMLInputElement>('input[data-cf-bar-negative-color]')?.value,
    ).toBe('#ff0000');
    expect(
      document.querySelector<HTMLSelectElement>('select[data-cf-bar-border-style]')?.value,
    ).toBe('solid');
    expect(document.querySelector<HTMLInputElement>('input[data-cf-bar-axis-color]')?.value).toBe(
      '#000000',
    );
    expect(
      document.querySelector<HTMLInputElement>('input[data-cf-bar-axis-color]')?.disabled,
    ).toBe(true);

    document
      .querySelector<HTMLButtonElement>('.fc-conddlg__addrow .fc-fmtdlg__btn--primary')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    const saved = store.getState().conditional.rules[0];
    expect(saved).toMatchObject({
      kind: 'data-bar',
      color: 'rgba(99, 142, 198, 0.5)',
      negativeColor: 'rgb(255, 0, 0)',
      negativeBorderColor: 'rgba(153, 0, 0, 0.75)',
      axisColor: 'rgba(0, 0, 0, 0.25)',
      axisPosition: 'none',
    });
    if (saved?.kind === 'data-bar') expect(saved.borderColor).toBeUndefined();
    handle.detach();
  });

  it('preserves data-bar settings while editing and cancelling a draft', () => {
    const rule = {
      kind: 'data-bar' as const,
      range: { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 0 },
      color: '#638ec6',
      min: { kind: 'number' as const, value: 10 },
      max: { kind: 'percent' as const, value: 90 },
      direction: 'right-to-left' as const,
      gradient: false,
      showValue: true,
    };
    mutators.addConditionalRule(store, rule);
    const handle = attachConditionalDialog({ host, store, strings: en });
    handle.open({ mode: 'edit', editIndex: 0 });

    const minType = document.querySelector<HTMLSelectElement>('select[data-cf-bar-min-type]');
    const minValue = document.querySelector<HTMLInputElement>('input[data-cf-bar-min-value]');
    const maxType = document.querySelector<HTMLSelectElement>('select[data-cf-bar-max-type]');
    const maxValue = document.querySelector<HTMLInputElement>('input[data-cf-bar-max-value]');
    const direction = document.querySelector<HTMLSelectElement>('select[data-cf-bar-direction]');
    if (!minType || !minValue || !maxType || !maxValue || !direction) {
      throw new Error('missing data-bar edit controls');
    }
    expect(minType.value).toBe('number');
    expect(minValue.value).toBe('10');
    expect(maxType.value).toBe('percent');
    expect(maxValue.value).toBe('90');
    expect(direction.value).toBe('right-to-left');

    minValue.value = '25';
    direction.value = 'left-to-right';
    document
      .querySelector<HTMLButtonElement>('.fc-fmtdlg__footer .fc-fmtdlg__btn')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(store.getState().conditional.rules[0]).toMatchObject({
      min: { kind: 'number', value: 10 },
      direction: 'right-to-left',
    });

    handle.open({ mode: 'edit', editIndex: 0 });
    expect(document.querySelector<HTMLInputElement>('input[data-cf-bar-min-value]')?.value).toBe(
      '10',
    );
    expect(document.querySelector<HTMLSelectElement>('select[data-cf-bar-direction]')?.value).toBe(
      'right-to-left',
    );

    document
      .querySelector<HTMLButtonElement>('.fc-conddlg__addrow .fc-fmtdlg__btn--primary')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(store.getState().conditional.rules[0]).toMatchObject({
      min: { kind: 'number', value: 10 },
      max: { kind: 'percent', value: 90 },
      direction: 'right-to-left',
    });
    const saved = store.getState().conditional.rules[0];
    if (saved?.kind === 'data-bar') {
      expect(saved.negativeColor).toBeUndefined();
      expect(saved.axisPosition).toBeUndefined();
      expect(saved.axisColor).toBeUndefined();
    }
    handle.detach();
  });
});
