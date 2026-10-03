import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { History } from '../../../../src/commands/history.js';
import { en, ja } from '../../../../src/i18n/strings.js';
import {
  appendConditionalApplyFormatControls,
  applyPatchToConditionalApplyControls,
  attachRangePickerButton,
  collectConditionalApplyPatch,
  updateRangePickerLabel,
} from '../../../../src/index.js';
import { attachConditionalDialog } from '../../../../src/interact/conditional-dialog.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../../src/store/store.js';
import { root, setRange } from './fixtures.js';

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

  it('initializes the style from the requested kind on every open', () => {
    const handle = attachConditionalDialog({ host, store });
    const style = () =>
      document.querySelector<HTMLSelectElement>('.fc-conddlg__style-row select')?.value;
    try {
      handle.open({ mode: 'new', kind: 'data-bar' });
      expect(style()).toBe('data-bar');
      handle.close();
      handle.open({ mode: 'new', kind: 'icon-set' });
      expect(style()).toBe('icon-set');
      expect(document.querySelector('[data-cf-icon-operator="0"]')?.closest('[hidden]')).toBeNull();
      handle.close();
      handle.open({ mode: 'new', kind: 'cell-value' });
      expect(style()).toBe('classic');
    } finally {
      handle.detach();
    }
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

  it('updates the applies-to range through the shared range picker', () => {
    setRange(store, 0, 0, 4, 2);
    const handle = attachConditionalDialog({ host, store, strings: en });
    handle.open();

    const rangeInput = document.querySelector<HTMLInputElement>(
      '.fc-conddlg__form input[type="text"]',
    );
    const picker = document.querySelector<HTMLButtonElement>(
      '[data-range-picker="conditional-format-range"]',
    );
    expect(rangeInput?.closest('.fc-range-picker')).toBeTruthy();
    expect(picker?.getAttribute('aria-label')).toBe('Select range');

    picker?.click();
    expect(picker?.dataset.rangePickerActive).toBe('true');
    expect(
      document.querySelector('.fc-conddlg')?.classList.contains('fc-fmtdlg--range-picking'),
    ).toBe(true);
    setRange(store, 1, 1, 5, 3);
    expect(rangeInput?.value).toBe('B2:D6');

    handle.close();
    expect(picker?.dataset.rangePickerActive).toBe('false');
    expect(
      document.querySelector('.fc-conddlg')?.classList.contains('fc-fmtdlg--range-picking'),
    ).toBe(false);
    handle.detach();
  });

  it('exposes shared range picker controls through the public entrypoint', () => {
    const input = document.createElement('input');
    host.appendChild(input);

    const button = attachRangePickerButton(input, {
      label: 'Select range',
      getValue: () => 'C3:D4',
      kind: 'public-range',
    });

    expect(button.dataset.rangePicker).toBe('public-range');
    expect(input.closest('.fc-range-picker')).toBeTruthy();
    button.click();
    expect(input.value).toBe('C3:D4');

    updateRangePickerLabel(button, 'Pick cells');
    expect(button.title).toBe('Pick cells');
    expect(button.getAttribute('aria-label')).toBe('Pick cells');

    const source = readFileSync(join(root, 'src/interact/range-picker-control.ts'), 'utf8');
    expect(source).toContain("import { createInteractionButton } from './chip-button.js'");
    expect(source).toContain('const button = createInteractionButton({');
    expect(source).not.toContain("document.createElement('button')");
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

  it('uses shared controls for conditional rule apply format patches', () => {
    const controls = appendConditionalApplyFormatControls(host, en.conditionalDialog);

    applyPatchToConditionalApplyControls(controls, {
      fill: '#ffc7ce',
      color: '#9c0006',
      bold: true,
      underline: true,
    });

    expect(controls.fillToggle.checked).toBe(true);
    expect(controls.fillInput.value).toBe('#ffc7ce');
    expect(controls.fontToggle.checked).toBe(true);
    expect(controls.fontInput.value).toBe('#9c0006');
    expect(controls.bold.checked).toBe(true);
    expect(controls.underline.checked).toBe(true);
    expect(collectConditionalApplyPatch(controls)).toEqual({
      fill: '#ffc7ce',
      color: '#9c0006',
      bold: true,
      underline: true,
    });
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
    const onChanged = vi.fn();
    const handle = attachConditionalDialog({ host, store, history, onChanged });
    handle.open();

    const valueA = document.querySelector<HTMLInputElement>(
      '.fc-conddlg__sub input[type="text"]',
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
    expect(onChanged).toHaveBeenCalledTimes(1);

    const removeBtn = document.querySelector<HTMLButtonElement>(
      '.fc-conddlg__item .fc-fmtdlg__btn',
    );
    removeBtn?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(store.getState().conditional.rules).toHaveLength(0);
    expect(onChanged).toHaveBeenCalledTimes(2);

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

  it('adds a text cell-value rule from the classic dialog', () => {
    setRange(store, 0, 0, 1, 0);
    const handle = attachConditionalDialog({ host, store });
    handle.open({ kind: 'cell-value', cellValueOp: '=' });

    const valueA = document.querySelector<HTMLInputElement>(
      '.fc-conddlg__sub input[type="text"]',
    ) as HTMLInputElement;
    valueA.value = 'Done';

    document
      .querySelector<HTMLButtonElement>('.fc-conddlg__addrow .fc-fmtdlg__btn--primary')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    expect(store.getState().conditional.rules[0]).toMatchObject({
      kind: 'cell-value',
      op: '=',
      a: 'Done',
    });
    handle.detach();
  });

  it('edits an existing session rule through the shared rule dialog', () => {
    const history = new History();
    mutators.addConditionalRule(store, {
      kind: 'cell-value',
      range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
      op: '>',
      a: 10,
      apply: { fill: '#ffc7ce', color: '#9c0006' },
    });
    const handle = attachConditionalDialog({ host, store, history });
    handle.open({ mode: 'edit', editIndex: 0 });

    expect(document.querySelector<HTMLElement>('.fc-fmtdlg__header')?.textContent).toBe(
      ja.conditionalDialog.editRuleTitle,
    );
    const rangeInput = document.querySelector<HTMLInputElement>(
      '.fc-conddlg__form input[type="text"]',
    );
    const valueInput = document.querySelector<HTMLInputElement>(
      '.fc-conddlg__sub input[type="text"]',
    );
    expect(rangeInput?.value).toBe('A1:A1');
    expect(valueInput?.value).toBe('10');
    if (!rangeInput || !valueInput) throw new Error('missing edit inputs');
    rangeInput.value = 'B2:B3';
    valueInput.value = '25';

    document
      .querySelector<HTMLButtonElement>('.fc-conddlg__addrow .fc-fmtdlg__btn--primary')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    expect(store.getState().conditional.rules).toHaveLength(1);
    expect(store.getState().conditional.rules[0]).toMatchObject({
      kind: 'cell-value',
      range: { sheet: 0, r0: 1, c0: 1, r1: 2, c1: 1 },
      op: '>',
      a: 25,
    });
    expect(history.undo()).toBe(true);
    expect(store.getState().conditional.rules[0]).toMatchObject({
      range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
      a: 10,
    });

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
    ).find(
      (select) =>
        Array.from(select.options).some((option) => option.value === 'data-bar') &&
        Array.from(select.options).some((option) => option.value === 'cell-value'),
    ) as HTMLSelectElement;
    const iconSelect = Array.from(
      document.querySelectorAll<HTMLSelectElement>('.fc-conddlg__form select'),
    ).find((select) => Array.from(select.options).some((option) => option.value === 'symbols3'));
    expect(iconSelect).toBeTruthy();
    expect(Array.from(iconSelect?.options ?? []).some((option) => option.value === 'boxes5')).toBe(
      true,
    );
    expect(Array.from(kindSelect.options).some((option) => option.value === 'average')).toBe(true);
    kindSelect.value = 'data-bar';
    kindSelect.dispatchEvent(new Event('change', { bubbles: true }));
    expect(subs[0]?.hidden).toBe(true);
    expect(subs[2]?.hidden).toBe(false);

    handle.detach();
  });

  it('renders classic rule selects with the shared option contract', () => {
    const handle = attachConditionalDialog({ host, store, strings: en });
    handle.open({ kind: 'icon-set' });

    const selects = Array.from(
      document.querySelectorAll<HTMLSelectElement>('.fc-conddlg__form select'),
    );
    const styleSelect = selects.find((select) =>
      Array.from(select.options).some((option) => option.value === 'classic'),
    );
    const kindSelect = selects.find((select) =>
      Array.from(select.options).some((option) => option.value === 'cell-value'),
    );
    const iconSelect = selects.find((select) =>
      Array.from(select.options).some((option) => option.value === 'traffic3'),
    );
    expect(
      Array.from(styleSelect?.options ?? [], (option) => [option.value, option.textContent]),
    ).toEqual([
      ['two-color-scale', '2-Color Scale'],
      ['three-color-scale', '3-Color Scale'],
      ['data-bar', 'Data bar'],
      ['icon-set', 'Icon set'],
      ['classic', 'Classic'],
    ]);
    expect(Array.from(kindSelect?.options ?? [], (option) => option.value)).toContain(
      'date-occurring',
    );
    expect(Array.from(iconSelect?.options ?? [], (option) => option.value)).toContain('boxes5');

    handle.detach();
  });

  it('switches the five Excel rule styles into their matching form kinds', () => {
    const handle = attachConditionalDialog({ host, store });
    handle.open();
    const selects = Array.from(
      document.querySelectorAll<HTMLSelectElement>('.fc-conddlg__form select'),
    );
    const style = selects.find((select) =>
      Array.from(select.options).some((option) => option.value === 'two-color-scale'),
    ) as HTMLSelectElement;
    const kind = selects.find((select) =>
      Array.from(select.options).some((option) => option.value === 'cell-value'),
    ) as HTMLSelectElement;
    for (const [value, expected] of [
      ['two-color-scale', 'color-scale'],
      ['three-color-scale', 'color-scale'],
      ['data-bar', 'data-bar'],
      ['icon-set', 'icon-set'],
    ] as const) {
      style.value = value;
      style.dispatchEvent(new Event('change', { bubbles: true }));
      expect(kind.value).toBe(expected);
    }
    style.value = 'classic';
    style.dispatchEvent(new Event('change', { bubbles: true }));
    expect(kind.hidden).toBe(false);
    handle.detach();
  });

  it('opens Format Cells in dxf mode from the classic Custom format and applies its patch', () => {
    const handle = attachConditionalDialog({ host, store, strings: en });
    handle.open({ mode: 'new', kind: 'cell-value' });
    const preset = Array.from(
      document.querySelectorAll<HTMLSelectElement>('.fc-conddlg__form select'),
    ).find((select) => Array.from(select.options).some((option) => option.value === 'red-border'));
    if (!preset) throw new Error('missing conditional format preset select');
    preset.value = 'custom';
    preset.dispatchEvent(new Event('change', { bubbles: true }));

    const dxf = Array.from(document.querySelectorAll<HTMLElement>('.fc-fmtdlg__panel')).find(
      (panel) => !panel.classList.contains('fc-conddlg__panel'),
    );
    const fillTab = dxf?.querySelector<HTMLButtonElement>('button[data-fc-tab="fill"]');
    const fontTab = dxf?.querySelector<HTMLButtonElement>('button[data-fc-tab="font"]');
    const fill = dxf?.querySelector<HTMLInputElement>('input[data-fc-color="fill"]');
    const bold = dxf?.querySelector<HTMLInputElement>('input[data-fc-check="bold"]');
    if (!dxf || !fillTab || !fontTab || !fill || !bold)
      throw new Error('missing dxf format controls');
    expect(dxf.querySelector('button[data-fc-tab="align"]')).toHaveProperty('hidden', true);
    fillTab.click();
    fill.value = '#ddeeff';
    fill.dispatchEvent(new Event('input', { bubbles: true }));
    fontTab.click();
    bold.checked = true;
    bold.dispatchEvent(new Event('change', { bubbles: true }));
    dxf.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();

    document
      .querySelector<HTMLButtonElement>('.fc-conddlg__addrow .fc-fmtdlg__btn--primary')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(store.getState().conditional.rules[0]).toMatchObject({
      kind: 'cell-value',
      apply: { fill: '#ddeeff', bold: true },
    });
    handle.detach();
  });

  it('selects the matching Excel style when editing a non-classic rule', () => {
    mutators.addConditionalRule(store, {
      kind: 'color-scale',
      range: { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 0 },
      stops: ['#f8696b', '#ffeb84', '#63be7b'],
    });
    const handle = attachConditionalDialog({ host, store });
    handle.open({ mode: 'edit', editIndex: 0 });
    const style = Array.from(
      document.querySelectorAll<HTMLSelectElement>('.fc-conddlg__form select'),
    ).find((select) => Array.from(select.options).some((option) => option.value === 'classic'));
    expect(style?.value).toBe('three-color-scale');
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

  it('keeps the conditional format manager on compact desktop dialog geometry', () => {
    const css = readFileSync(
      join(root, 'src/styles/core/app/dialog-modules/conditional-and-names.css'),
      'utf8',
    );

    expect(css).toMatch(
      /\.fc-conddlg__list\s*\{[\s\S]*?gap: 0;[\s\S]*?border-radius: 0;[\s\S]*?padding: 0;/,
    );
    expect(css).toMatch(
      /\.fc-conddlg__item\s*\{[\s\S]*?min-height: 28px;[\s\S]*?padding: 4px 8px;[\s\S]*?border-bottom: 1px solid var\(--fc-rule-subtle\);[\s\S]*?border-radius: 0;/,
    );
    expect(css).toMatch(
      /\.fc-conddlg__form\s*\{[\s\S]*?border-radius: 2px;[\s\S]*?background: var\(--fc-bg, Canvas\);[\s\S]*?padding: 10px 12px;/,
    );
  });
});
