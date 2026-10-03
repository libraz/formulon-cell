import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { addrKey } from '../../../../src/engine/address.js';
import { Spreadsheet } from '../../../../src/mount.js';
import { mutators } from '../../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../../test-utils/mount.js';
import { seedNumber, seedText, stubHelpers } from './fixtures.js';

vi.setConfig({ testTimeout: 20_000 });

describe('Spreadsheet.mountToolbar', () => {
  let sheet: MountedStubSheet;
  let host: HTMLElement;

  beforeEach(async () => {
    sheet = await mountStubSheet({ locale: 'en' });
    host = document.createElement('div');
    document.body.appendChild(host);
  });

  afterEach(() => {
    sheet.dispose();
    host.remove();
  });

  it('sorts the active column through the Home Sort dropdown', () => {
    seedNumber(sheet, 0, 0, 3);
    seedNumber(sheet, 1, 0, 1);
    seedNumber(sheet, 2, 0, 2);
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 0 });
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 0, col: 0 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    const sortButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="sortFilterHome"]',
    );
    expect(sortButton).toBeTruthy();
    sortButton?.click();
    expect(host.querySelectorAll('#menu-sort-home .fc-tb__menu-item--iconic').length).toBe(11);
    const ascendingButton = host.querySelector<HTMLButtonElement>('[data-sort="asc"]');
    expect(ascendingButton).toBeTruthy();
    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: ascendingButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);

    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
      kind: 'number',
      value: 1,
    });
    expect(sheet.workbook.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({
      kind: 'number',
      value: 2,
    });
    expect(sheet.workbook.getValue({ sheet: 0, row: 2, col: 0 })).toEqual({
      kind: 'number',
      value: 3,
    });

    tb.dispose();
  });

  it('reflects filter state in the Sort & Filter dropdown', () => {
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 0 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    const sortButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="sortFilterHome"]',
    );
    expect(sortButton).toBeTruthy();
    sortButton?.click();
    const menu = host.querySelector<HTMLElement>('#menu-sort-home');
    const filterButton = menu?.querySelector<HTMLButtonElement>('[data-sort="filter"]');
    const clearButton = menu?.querySelector<HTMLButtonElement>('[data-sort="filter-clear"]');
    const reapplyButton = menu?.querySelector<HTMLButtonElement>('[data-sort="filter-reapply"]');
    expect(filterButton?.getAttribute('aria-pressed')).toBe('false');
    expect(clearButton?.disabled).toBe(true);
    expect(clearButton?.dataset.menuDisabledReason).toBe('There is no filter to clear.');
    expect(reapplyButton?.disabled).toBe(true);
    expect(reapplyButton?.dataset.menuDisabledReason).toBe(
      'There are no filter criteria to reapply.',
    );

    const range = { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 0 };
    mutators.setFilterRange(sheet.instance.store, range);
    sortButton?.click();
    expect(filterButton?.getAttribute('aria-pressed')).toBe('true');
    expect(filterButton?.classList.contains('fc-tb__menu-item--active')).toBe(true);
    expect(clearButton?.disabled).toBe(false);
    expect(clearButton?.dataset.menuDisabledReason).toBeUndefined();
    expect(reapplyButton?.disabled).toBe(true);

    sheet.instance.store.setState((state) => ({
      ...state,
      ui: {
        ...state.ui,
        filterCriteria: [{ range, byCol: 0, hiddenValues: ['1'] }],
      },
    }));
    sortButton?.click();
    expect(reapplyButton?.disabled).toBe(false);
    expect(reapplyButton?.getAttribute('aria-disabled')).toBe('false');
    expect(reapplyButton?.dataset.menuDisabledReason).toBeUndefined();

    tb.dispose();
  });

  it('routes filter and manager actions through the Home Sort dropdown', async () => {
    seedText(sheet, 0, 0, 'Name');
    seedText(sheet, 0, 1, 'Score');
    seedText(sheet, 1, 0, 'Alice');
    seedNumber(sheet, 1, 1, 10);
    seedText(sheet, 2, 0, 'Bob');
    seedNumber(sheet, 2, 1, 20);
    seedText(sheet, 4, 0, 'Name');
    seedText(sheet, 5, 0, 'Bob');
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 });
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 2, col: 0 });
    const openConditionalDialog = vi
      .spyOn(sheet.instance, 'openConditionalDialog')
      .mockImplementation(() => undefined);
    const openNamedRangeDialog = vi
      .spyOn(sheet.instance, 'openNamedRangeDialog')
      .mockImplementation(() => undefined);
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    const sortButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="sortFilterHome"]',
    );
    expect(sortButton).toBeTruthy();
    const clickSort = async (action: string): Promise<void> => {
      tb.dropdownsApi?.openDynamicRibbonDropdown(
        { command: 'sortFilterHome', menuId: 'menu-sort-home' },
        sortButton as HTMLButtonElement,
      );
      const button = host.querySelector<HTMLButtonElement>(`[data-sort="${action}"]`);
      expect(button).toBeTruthy();
      const event = new MouseEvent('click', { bubbles: true });
      Object.defineProperty(event, 'target', { value: button });
      expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);
      await Promise.resolve();
    };

    await clickSort('filter');
    expect(sheet.instance.store.getState().ui.filterRange).toEqual({
      sheet: 0,
      r0: 0,
      c0: 0,
      r1: 2,
      c1: 1,
    });

    await clickSort('filter-by-value');
    expect(sheet.instance.store.getState().layout.hiddenRows.has(1)).toBe(true);
    expect(sheet.instance.store.getState().layout.hiddenRows.has(2)).toBe(false);

    await clickSort('filter-clear');
    expect(sheet.instance.store.getState().ui.filterRange).toBeNull();
    expect(sheet.instance.store.getState().layout.hiddenRows.has(1)).toBe(false);

    await clickSort('filter-advanced');
    const advancedDialog = document.body.querySelector<HTMLElement>('.fc-tb__dlg');
    expect(advancedDialog?.textContent).toContain('Advanced Filter');
    expect(advancedDialog?.querySelector('.fc-advfilter__ranges')).toBeTruthy();
    const inputs = Array.from(advancedDialog?.querySelectorAll<HTMLInputElement>('input') ?? []);
    expect(inputs).toHaveLength(4);
    const listInput = inputs[0];
    const criteriaInput = inputs[1];
    if (!listInput || !criteriaInput) throw new Error('Expected Advanced Filter range inputs.');
    listInput.value = 'A1:B3';
    criteriaInput.value = 'A5:A6';
    advancedDialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await Promise.resolve();
    expect(sheet.instance.store.getState().layout.hiddenRows.has(1)).toBe(true);
    expect(sheet.instance.store.getState().layout.hiddenRows.has(2)).toBe(false);

    await clickSort('conditional');
    await clickSort('named');
    expect(openConditionalDialog).toHaveBeenCalledTimes(1);
    expect(openNamedRangeDialog).toHaveBeenCalledTimes(1);

    tb.dispose();
  });

  it('opens custom sort and remove duplicates from the Home Sort dropdown', async () => {
    seedNumber(sheet, 0, 0, 2);
    seedNumber(sheet, 1, 0, 1);
    seedNumber(sheet, 2, 0, 1);
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 0 });
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 0, col: 0 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    const sortButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="sortFilterHome"]',
    );
    expect(sortButton).toBeTruthy();
    sortButton?.click();
    const customButton = host.querySelector<HTMLButtonElement>('[data-sort="custom"]');
    expect(customButton).toBeTruthy();
    const customEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(customEvent, 'target', { value: customButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(customEvent)).toBe(true);

    const sortDialog = document.body.querySelector<HTMLElement>('.fc-tb__dlg');
    expect(sortDialog?.textContent).toContain('Sort by');
    expect(sortDialog?.textContent).toContain('Add Level');
    expect(sortDialog?.querySelectorAll('.fc-sortdlg__level')).toHaveLength(1);
    sortDialog?.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'Enter' }));
    await Promise.resolve();
    await Promise.resolve();

    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
      kind: 'number',
      value: 1,
    });

    sortButton?.click();
    const dedupeButton = host.querySelector<HTMLButtonElement>('[data-sort="dedupe"]');
    expect(dedupeButton).toBeTruthy();
    const dedupeEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(dedupeEvent, 'target', { value: dedupeButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(dedupeEvent)).toBe(true);

    const dedupeDialog = document.body.querySelector<HTMLElement>('.fc-tb__dlg');
    expect(dedupeDialog?.textContent).toContain('Remove Duplicates');
    expect(dedupeDialog?.querySelectorAll('.fc-dedupedlg__column')).toHaveLength(1);
    dedupeDialog?.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'Enter' }));
    await Promise.resolve();
    await Promise.resolve();

    expect(sheet.workbook.getValue({ sheet: 0, row: 2, col: 0 }).kind).toBe('blank');

    tb.dispose();
  });

  it('opens Text to Columns delimiter dialog from primary click and keeps presets secondary', async () => {
    seedText(sheet, 0, 0, 'alpha,beta');
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('data');

    const textToColumnsButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="textToColumns"]',
    );
    expect(textToColumnsButton).toBeTruthy();
    textToColumnsButton?.click();
    const dialog = document.body.querySelector<HTMLElement>('.fc-tb__dlg');
    expect(dialog?.textContent).toContain('Convert Text to Columns');
    expect(dialog?.textContent).toContain('Original data type');
    expect(dialog?.textContent).toContain('Data preview');
    const comma = dialog?.querySelector<HTMLInputElement>('[data-dialog-field="delimiter-,"]');
    expect(comma?.checked).toBe(true);
    dialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await Promise.resolve();

    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
      kind: 'text',
      value: 'alpha',
    });
    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({
      kind: 'text',
      value: 'beta',
    });

    seedText(sheet, 1, 0, 'one,two');
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 1, c0: 0, r1: 1, c1: 0 });
    tb.dropdownsApi?.openDynamicRibbonDropdown({
      command: 'textToColumns',
      menuId: 'menu-text-to-columns',
    });
    expect(host.querySelectorAll('#menu-text-to-columns .fc-tb__menu-item--iconic').length).toBe(5);
    const commaButton = host.querySelector<HTMLButtonElement>(
      '[data-text-to-columns-delimiter=","]',
    );
    expect(commaButton).toBeTruthy();
    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: commaButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);

    expect(sheet.workbook.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({
      kind: 'text',
      value: 'one',
    });
    expect(sheet.workbook.getValue({ sheet: 0, row: 1, col: 1 })).toEqual({
      kind: 'text',
      value: 'two',
    });

    tb.dispose();
  });

  it('opens Data Validation from primary click and keeps circle actions secondary', () => {
    seedNumber(sheet, 0, 0, 5);
    seedNumber(sheet, 1, 0, 20);
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 });
    sheet.instance.store.setState((state) => {
      const formats = new Map(state.format.formats);
      formats.set(addrKey({ sheet: 0, row: 0, col: 0 }), {
        validation: { kind: 'whole', op: '<=', a: 10 },
      });
      formats.set(addrKey({ sheet: 0, row: 1, col: 0 }), {
        validation: { kind: 'whole', op: '<=', a: 10 },
      });
      return { ...state, format: { ...state.format, formats } };
    });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('data');

    const validationButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="dataValidation"]',
    );
    expect(validationButton).toBeTruthy();
    validationButton?.click();
    const formatDialog = document.body.querySelector<HTMLElement>('.fc-fmtdlg:not([hidden])');
    expect(formatDialog?.hidden).toBe(false);
    expect(formatDialog?.classList.contains('fc-fmtdlg--data-validation')).toBe(true);
    expect(formatDialog?.querySelector<HTMLElement>('.fc-fmtdlg__title')?.textContent).toBe(
      'Data validation',
    );
    expect(formatDialog?.querySelector<HTMLElement>('.fc-fmtdlg__tabs')?.hidden).toBe(true);
    expect(
      formatDialog?.querySelector<HTMLSelectElement>('select[aria-label="Kind"]'),
    ).toBeTruthy();
    formatDialog
      ?.querySelector<HTMLButtonElement>('.fc-fmtdlg__close')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(formatDialog?.hidden).toBe(true);

    tb.dropdownsApi?.openDynamicRibbonDropdown({
      command: 'dataValidation',
      menuId: 'menu-data-validation',
    });
    expect(host.querySelectorAll('#menu-data-validation .fc-tb__menu-item--iconic').length).toBe(4);
    const circleButton = host.querySelector<HTMLButtonElement>(
      '[data-validation-action="circle-invalid"]',
    );
    const initialClearCirclesButton = host.querySelector<HTMLButtonElement>(
      '[data-validation-action="clear-circles"]',
    );
    const initialClearRulesButton = host.querySelector<HTMLButtonElement>(
      '[data-validation-action="clear-rules"]',
    );
    expect(circleButton).toBeTruthy();
    expect(circleButton?.disabled).toBe(false);
    expect(initialClearCirclesButton?.disabled).toBe(true);
    expect(initialClearCirclesButton?.dataset.menuDisabledReason).toBe(
      'There are no validation circles to clear.',
    );
    expect(initialClearRulesButton?.disabled).toBe(false);
    const circleEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(circleEvent, 'target', { value: circleButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(circleEvent)).toBe(true);
    expect(sheet.instance.store.getState().errorIndicators.validationCircles).toEqual(
      new Set([addrKey({ sheet: 0, row: 1, col: 0 })]),
    );

    tb.dropdownsApi?.openDynamicRibbonDropdown({
      command: 'dataValidation',
      menuId: 'menu-data-validation',
    });
    const clearCirclesButton = host.querySelector<HTMLButtonElement>(
      '[data-validation-action="clear-circles"]',
    );
    expect(clearCirclesButton).toBeTruthy();
    expect(clearCirclesButton?.disabled).toBe(false);
    const clearCirclesEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(clearCirclesEvent, 'target', { value: clearCirclesButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(clearCirclesEvent)).toBe(true);
    expect(sheet.instance.store.getState().errorIndicators.validationCircles.size).toBe(0);

    tb.dropdownsApi?.openDynamicRibbonDropdown({
      command: 'dataValidation',
      menuId: 'menu-data-validation',
    });
    expect(clearCirclesButton?.disabled).toBe(true);
    expect(clearCirclesButton?.dataset.menuDisabledReason).toBe(
      'There are no validation circles to clear.',
    );

    tb.dropdownsApi?.openDynamicRibbonDropdown({
      command: 'dataValidation',
      menuId: 'menu-data-validation',
    });
    const clearRulesButton = host.querySelector<HTMLButtonElement>(
      '[data-validation-action="clear-rules"]',
    );
    expect(clearRulesButton).toBeTruthy();
    const clearRulesEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(clearRulesEvent, 'target', { value: clearRulesButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(clearRulesEvent)).toBe(true);
    expect(
      sheet.instance.store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))
        ?.validation,
    ).toBeUndefined();
    expect(
      sheet.instance.store.getState().format.formats.get(addrKey({ sheet: 0, row: 1, col: 0 }))
        ?.validation,
    ).toBeUndefined();

    tb.dropdownsApi?.openDynamicRibbonDropdown({
      command: 'dataValidation',
      menuId: 'menu-data-validation',
    });
    expect(circleButton?.disabled).toBe(true);
    expect(circleButton?.dataset.menuDisabledReason).toBe(
      'The selection does not contain data validation.',
    );
    expect(clearRulesButton?.disabled).toBe(true);
    expect(clearRulesButton?.dataset.menuDisabledReason).toBe(
      'The selection does not contain data validation.',
    );

    tb.dispose();
  });

  it('circles invalid validation cells in huge selections by scanning validation formats only', () => {
    seedNumber(sheet, 900_000, 0, 20);
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 1048575, c1: 0 });
    sheet.instance.store.setState((state) => {
      const formats = new Map(state.format.formats);
      formats.set(addrKey({ sheet: 0, row: 900_000, col: 0 }), {
        validation: { kind: 'whole', op: '<=', a: 10 },
      });
      return { ...state, format: { ...state.format, formats } };
    });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    tb.dropdownsApi?.openDynamicRibbonDropdown({
      command: 'dataValidation',
      menuId: 'menu-data-validation',
    });
    const circleButton = host.querySelector<HTMLButtonElement>(
      '[data-validation-action="circle-invalid"]',
    );
    expect(circleButton?.disabled).toBe(false);

    const circleEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(circleEvent, 'target', { value: circleButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(circleEvent)).toBe(true);
    expect(sheet.instance.store.getState().errorIndicators.validationCircles).toEqual(
      new Set([addrKey({ sheet: 0, row: 900_000, col: 0 })]),
    );

    tb.dispose();
  });
});
