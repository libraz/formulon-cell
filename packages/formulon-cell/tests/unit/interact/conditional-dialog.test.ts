import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { en, ja } from '../../../src/i18n/strings.js';
import { attachConditionalDialog } from '../../../src/interact/conditional-dialog.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

const setRange = (
  store: SpreadsheetStore,
  r0: number,
  c0: number,
  r1: number,
  c1: number,
): void => {
  store.setState((s) => ({
    ...s,
    selection: {
      active: { sheet: 0, row: r0, col: c0 },
      anchor: { sheet: 0, row: r0, col: c0 },
      range: { sheet: 0, r0, c0, r1, c1 },
    },
  }));
};

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

  it('mounts a hidden overlay and pre-fills selection range on open', () => {
    setRange(store, 0, 0, 4, 2);
    const handle = attachConditionalDialog({ host, store });
    const overlay = document.querySelector<HTMLElement>('.fc-conddlg');
    expect(overlay?.hidden).toBe(true);

    handle.open();
    expect(overlay?.hidden).toBe(false);
    const rangeInput = document.querySelector<HTMLInputElement>(
      '.fc-conddlg__form input[type="text"]',
    );
    expect(rangeInput?.value).toBe('A1:C5');
    handle.detach();
  });

  it('localizes conditional format preview sample text', () => {
    const jaHandle = attachConditionalDialog({ host, store, strings: ja });
    jaHandle.open();
    expect(document.querySelector<HTMLElement>('.fc-conddlg__preview')?.textContent).toBe(
      'Aaあぁアァ亜字',
    );
    jaHandle.detach();

    document.body.innerHTML = '';
    document.body.appendChild(host);

    const enHandle = attachConditionalDialog({ host, store, strings: en });
    enHandle.open();
    expect(document.querySelector<HTMLElement>('.fc-conddlg__preview')?.textContent).toBe(
      'AaBbCcYyZz',
    );
    enHandle.detach();
  });

  it('localizes conditional rule summaries in the manager list', () => {
    const range = { sheet: 0, r0: 0, c0: 0, r1: 4, c1: 0 };
    mutators.addConditionalRule(store, {
      kind: 'color-scale',
      range,
      stops: ['#ff0000', '#00ff00'],
    });
    mutators.addConditionalRule(store, {
      kind: 'data-bar',
      range,
      color: '#638ec6',
      gradient: true,
      showValue: true,
    });
    mutators.addConditionalRule(store, {
      kind: 'icon-set',
      range,
      icons: 'traffic3',
      showValue: false,
    });
    mutators.addConditionalRule(store, {
      kind: 'top-bottom',
      range,
      mode: 'top',
      n: 3,
      percent: true,
      apply: { fill: '#ffc7ce' },
    });
    mutators.addConditionalRule(store, {
      kind: 'average',
      range,
      mode: 'equal-or-below',
      apply: { fill: '#ffc7ce' },
    });
    mutators.addConditionalRule(store, {
      kind: 'date-occurring',
      range,
      period: 'last7',
      apply: { fill: '#ffc7ce' },
    });

    const handle = attachConditionalDialog({ host, store, strings: ja });
    handle.open();

    const summary = document.querySelector<HTMLElement>('.fc-conddlg__list')?.textContent ?? '';
    expect(summary).toContain('2 段階');
    expect(summary).toContain('塗りつぶし (グラデーション)');
    expect(summary).toContain('3 信号');
    expect(summary).toContain('アイコンのみ表示');
    expect(summary).toContain('上位 3%');
    expect(summary).toContain('平均以下');
    expect(summary).toContain('過去 7 日間');
    expect(summary).not.toContain('last7');
    expect(summary).not.toContain('traffic3');

    handle.detach();
  });

  it('adds a cell-value rule then removes it', () => {
    setRange(store, 0, 0, 1, 0);
    const history = new History();
    const handle = attachConditionalDialog({ host, store, history });
    handle.open();

    const valueA = document.querySelector<HTMLInputElement>(
      '.fc-conddlg__sub input[type="number"]',
    ) as HTMLInputElement;
    valueA.value = '50';
    valueA.dispatchEvent(new Event('input', { bubbles: true }));

    const addBtn = document.querySelector<HTMLButtonElement>(
      '.fc-conddlg__addrow .fc-fmtdlg__btn--primary',
    );
    addBtn?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    const rules = store.getState().conditional.rules;
    expect(rules).toHaveLength(1);
    expect(rules[0]?.kind).toBe('cell-value');
    if (rules[0]?.kind === 'cell-value') {
      expect(rules[0].a).toBe(50);
      expect(rules[0].op).toBe('>');
    }
    expect(history.canUndo()).toBe(true);

    const removeBtn = document.querySelector<HTMLButtonElement>(
      '.fc-conddlg__item .fc-fmtdlg__btn',
    );
    removeBtn?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(store.getState().conditional.rules).toHaveLength(0);

    history.undo();
    expect(store.getState().conditional.rules).toHaveLength(1);

    history.undo();
    expect(store.getState().conditional.rules).toHaveLength(0);

    history.redo();
    expect(store.getState().conditional.rules).toHaveLength(1);

    history.redo();
    expect(store.getState().conditional.rules).toHaveLength(0);

    handle.detach();
  });

  it('accepts a single-cell reference in the applies-to field', () => {
    setRange(store, 0, 0, 4, 2);
    const handle = attachConditionalDialog({ host, store });
    handle.open();

    const rangeInput = document.querySelector<HTMLInputElement>(
      '.fc-conddlg__form input[type="text"]',
    );
    if (!rangeInput) throw new Error('missing range input');
    rangeInput.value = '$B$2';

    document
      .querySelector<HTMLButtonElement>('.fc-conddlg__addrow .fc-fmtdlg__btn--primary')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    expect(store.getState().conditional.rules[0]?.range).toEqual({
      sheet: 0,
      r0: 1,
      c0: 1,
      r1: 1,
      c1: 1,
    });
    handle.detach();
  });

  it('clear-all removes every rule', () => {
    const history = new History();
    mutators.addConditionalRule(store, {
      kind: 'data-bar',
      range: { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 },
      color: '#638ec6',
      showValue: true,
    });
    mutators.addConditionalRule(store, {
      kind: 'color-scale',
      range: { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 },
      stops: ['#ff0000', '#00ff00'],
    });

    const handle = attachConditionalDialog({ host, store, history });
    handle.open();

    const buttons = Array.from(document.querySelectorAll<HTMLButtonElement>('.fc-fmtdlg__btn'));
    const clearAll = buttons.find((b) => b.textContent === 'すべて削除') as HTMLButtonElement;
    clearAll.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    expect(store.getState().conditional.rules).toHaveLength(0);

    history.undo();
    expect(store.getState().conditional.rules).toHaveLength(2);

    history.redo();
    expect(store.getState().conditional.rules).toHaveLength(0);
    handle.detach();
  });

  it('switches subforms when kind changes', () => {
    const handle = attachConditionalDialog({ host, store });
    handle.open();

    const subs = document.querySelectorAll<HTMLDivElement>('.fc-conddlg__sub');
    expect(subs[0]?.hidden).toBe(false); // cell-value visible by default
    expect(subs[1]?.hidden).toBe(true);
    expect(subs[2]?.hidden).toBe(true);

    const kindSelect = Array.from(
      document.querySelectorAll<HTMLSelectElement>('.fc-conddlg__form select'),
    ).find((select) =>
      Array.from(select.options).some((option) => option.value === 'data-bar'),
    ) as HTMLSelectElement;
    const iconSelect = Array.from(
      document.querySelectorAll<HTMLSelectElement>('.fc-conddlg__form select'),
    ).find((select) =>
      Array.from(select.options).some((option) => option.value === 'symbols3'),
    );
    expect(iconSelect).toBeTruthy();
    expect(Array.from(iconSelect?.options ?? []).some((option) => option.value === 'boxes5')).toBe(
      true,
    );
    expect(
      Array.from(kindSelect.options).some((option) => option.value === 'average'),
    ).toBe(true);
    kindSelect.value = 'data-bar';
    kindSelect.dispatchEvent(new Event('change', { bubbles: true }));
    expect(subs[0]?.hidden).toBe(true);
    expect(subs[2]?.hidden).toBe(false);

    handle.detach();
  });

  it('opens with a preset rule kind from ribbon menu actions', () => {
    const handle = attachConditionalDialog({ host, store });
    handle.open({ kind: 'cell-value', cellValueOp: 'between' });

    const selects = Array.from(
      document.querySelectorAll<HTMLSelectElement>('.fc-conddlg__form select'),
    );
    const kindSelect = selects.find((select) =>
      Array.from(select.options).some((option) => option.value === 'cell-value'),
    );
    const opSelect = selects.find((select) =>
      Array.from(select.options).some((option) => option.value === 'between'),
    );
    expect(kindSelect?.value).toBe('cell-value');
    expect(opSelect?.value).toBe('between');

    handle.open({ kind: 'duplicates' });
    expect(kindSelect?.value).toBe('duplicates');
    handle.detach();
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
      thresholds: [{ kind: 'number', value: 10 }, { kind: 'percent', value: 90 }],
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
    const reverse = checks.find((input) => input.nextElementSibling?.textContent === 'Reverse order');
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
    thresholdValues[0]!.value = '25';
    thresholdValues[1]!.value = '75';

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

  it('opens a New Formatting Rule mode with OK/Cancel and no rule manager list', () => {
    setRange(store, 0, 0, 1, 0);
    const handle = attachConditionalDialog({ host, store, strings: en });
    handle.open({ mode: 'new', kind: 'cell-value' });

    const overlay = document.querySelector<HTMLElement>('.fc-conddlg') as HTMLElement;
    expect(overlay.getAttribute('aria-label')).toBe('New Formatting Rule');
    expect(document.querySelector<HTMLElement>('.fc-fmtdlg__header')?.textContent).toBe(
      'New Formatting Rule',
    );
    expect(document.querySelector<HTMLElement>('.fc-conddlg__list')?.hidden).toBe(true);
    expect(document.querySelector<HTMLButtonElement>('.fc-conddlg__clear')?.hidden).toBe(true);

    const buttons = Array.from(document.querySelectorAll<HTMLButtonElement>('.fc-fmtdlg__btn'));
    expect(buttons.map((button) => button.textContent)).toContain('OK');
    expect(buttons.map((button) => button.textContent)).toContain('Cancel');

    const ok = buttons.find((button) => button.textContent === 'OK') as HTMLButtonElement;
    ok.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(store.getState().conditional.rules).toHaveLength(1);
    expect(overlay.hidden).toBe(true);

    handle.detach();
  });

  it('adds text and date-occurring rules from the classic dialog', () => {
    setRange(store, 1, 1, 2, 2);
    const handle = attachConditionalDialog({ host, store });
    handle.open({ kind: 'text-contains', text: 'due' });

    document
      .querySelector<HTMLButtonElement>('.fc-conddlg__addrow .fc-fmtdlg__btn--primary')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(store.getState().conditional.rules[0]).toMatchObject({
      kind: 'text-contains',
      range: { sheet: 0, r0: 1, c0: 1, r1: 2, c1: 2 },
      text: 'due',
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

  it('labels apply-format color controls and treats Enter as Add Rule', () => {
    setRange(store, 0, 0, 1, 0);
    const handle = attachConditionalDialog({ host, store });
    handle.open();

    const colorInputs = Array.from(
      document.querySelectorAll<HTMLInputElement>('.fc-conddlg__form input[type="color"]'),
    );
    expect(colorInputs.length).toBeGreaterThan(0);
    for (const input of colorInputs) expect(input.getAttribute('aria-label')).toBeTruthy();

    const toggleInputs = Array.from(
      document.querySelectorAll<HTMLInputElement>(
        '.fc-conddlg__form .fc-fmtdlg__row > input[type="checkbox"]',
      ),
    );
    expect(toggleInputs.length).toBeGreaterThan(0);
    for (const input of toggleInputs) expect(input.getAttribute('aria-label')).toBeTruthy();

    document
      .querySelector<HTMLElement>('.fc-conddlg')
      ?.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
    expect(store.getState().conditional.rules).toHaveLength(1);
    handle.detach();
  });

  it('Escape closes the overlay', () => {
    const handle = attachConditionalDialog({ host, store });
    handle.open();
    const overlay = document.querySelector<HTMLElement>('.fc-conddlg') as HTMLElement;
    overlay.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    expect(overlay.hidden).toBe(true);
    handle.detach();
  });
});
