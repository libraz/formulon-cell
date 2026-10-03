import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { hyperlinkAt, setHyperlink } from '../../../../src/commands/hyperlinks.js';
import { addrKey } from '../../../../src/engine/address.js';
import { Spreadsheet } from '../../../../src/mount.js';
import { mutators } from '../../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../../test-utils/mount.js';
import { dynamicDropdownNoopOverrides, seedNumber, stubHelpers } from './fixtures.js';

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

  it('keeps AutoSum presets as icon menu items through the shared dispatcher', () => {
    const dropdowns = dynamicDropdownNoopOverrides();
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: dropdowns,
      helpers: stubHelpers(),
    });

    const autosumButton = host.querySelector<HTMLButtonElement>('[data-ribbon-command="autosum"]');
    expect(autosumButton).toBeTruthy();
    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'autosum', menuId: 'menu-autosum-home' },
      autosumButton as HTMLButtonElement,
    );
    expect(host.querySelectorAll('#menu-autosum-home [data-autosum-fn]').length).toBe(6);
    expect(
      host.querySelectorAll('#menu-autosum-home .fc-tb__menu-icon--svg .fc-tb__menu-icon-svg')
        .length,
    ).toBe(1);
    const averageButton = host.querySelector<HTMLButtonElement>('[data-autosum-fn="AVERAGE"]');
    expect(averageButton).toBeTruthy();
    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: averageButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);
    expect(dropdowns.applyAutoSumFormula).toHaveBeenCalledWith('AVERAGE');

    tb.dispose();
  });

  it('applies AutoSum dropdown presets through the shared default action', () => {
    seedNumber(sheet, 0, 0, 10);
    seedNumber(sheet, 1, 0, 20);
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 2, col: 0 });
    const openFunctionArguments = vi
      .spyOn(sheet.instance, 'openFunctionArguments')
      .mockImplementation(() => undefined);
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    const autosumButton = host.querySelector<HTMLButtonElement>('[data-ribbon-command="autosum"]');
    expect(autosumButton).toBeTruthy();
    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'autosum', menuId: 'menu-autosum-home' },
      autosumButton as HTMLButtonElement,
    );
    const averageButton = host.querySelector<HTMLButtonElement>('[data-autosum-fn="AVERAGE"]');
    const averageEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(averageEvent, 'target', { value: averageButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(averageEvent)).toBe(true);
    expect(sheet.workbook.cellFormula({ sheet: 0, row: 2, col: 0 })).toBe('=AVERAGE(A1:A2)');
    expect(sheet.instance.store.getState().selection.active).toEqual({ sheet: 0, row: 2, col: 0 });

    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'autosum', menuId: 'menu-autosum-home' },
      autosumButton as HTMLButtonElement,
    );
    const moreButton = host.querySelector<HTMLButtonElement>('[data-autosum-fn="MORE"]');
    const moreEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(moreEvent, 'target', { value: moreButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(moreEvent)).toBe(true);
    expect(openFunctionArguments).toHaveBeenCalledTimes(1);

    tb.dispose();
  });

  it('keeps Calculation Options radio items icon-backed and dispatchable', () => {
    const dropdowns = dynamicDropdownNoopOverrides();
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: dropdowns,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('formulas');

    const calcButton = host.querySelector<HTMLButtonElement>('[data-ribbon-command="calcOptions"]');
    expect(calcButton).toBeTruthy();
    calcButton?.click();
    const menu = host.querySelector<HTMLElement>('#menu-calc-options');
    expect(menu?.querySelectorAll('.fc-tb__menu-item--iconic').length).toBe(6);
    const radios = Array.from(
      menu?.querySelectorAll<HTMLButtonElement>('[role="menuitemradio"]') ?? [],
    );
    expect(radios.map((button) => button.dataset.calcOption)).toEqual([
      'auto',
      'auto-no-table',
      'manual',
    ]);
    const manual = host.querySelector<HTMLButtonElement>('[data-calc-option="manual"]');
    expect(manual).toBeTruthy();
    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: manual });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);
    expect(dropdowns.applyCalcOptionAction).toHaveBeenCalledWith('manual');

    tb.dispose();
  });

  it('reflects the current calculation mode in the Calculation Options menu', () => {
    vi.spyOn(sheet.workbook, 'calcMode').mockReturnValue(1);
    const setCalcMode = vi.spyOn(sheet.workbook, 'setCalcMode').mockReturnValue(true);
    const recalc = vi.spyOn(sheet.instance, 'recalc').mockImplementation(() => undefined);
    const openIterativeDialog = vi
      .spyOn(sheet.instance, 'openIterativeDialog')
      .mockImplementation(() => undefined);
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('formulas');

    const calcButton = host.querySelector<HTMLButtonElement>('[data-ribbon-command="calcOptions"]');
    expect(calcButton).toBeTruthy();
    calcButton?.click();

    const auto = host.querySelector<HTMLButtonElement>('[data-calc-option="auto"]');
    const manual = host.querySelector<HTMLButtonElement>('[data-calc-option="manual"]');
    const autoNoTable = host.querySelector<HTMLButtonElement>('[data-calc-option="auto-no-table"]');
    expect(auto?.getAttribute('aria-checked')).toBe('false');
    expect(manual?.getAttribute('aria-checked')).toBe('true');
    expect(manual?.classList.contains('fc-tb__menu-item--active')).toBe(true);
    expect(autoNoTable?.getAttribute('aria-checked')).toBe('false');

    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: autoNoTable });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);
    expect(setCalcMode).toHaveBeenCalledWith(2);

    const calculateNow = host.querySelector<HTMLButtonElement>(
      '[data-calc-option="calculate-now"]',
    );
    const calculateSheet = host.querySelector<HTMLButtonElement>(
      '[data-calc-option="calculate-sheet"]',
    );
    const iterative = host.querySelector<HTMLButtonElement>('[data-calc-option="iterative"]');
    expect(calculateNow).toBeTruthy();
    expect(calculateSheet).toBeTruthy();
    expect(iterative).toBeTruthy();

    const calculateNowEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(calculateNowEvent, 'target', { value: calculateNow });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(calculateNowEvent)).toBe(true);
    const calculateSheetEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(calculateSheetEvent, 'target', { value: calculateSheet });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(calculateSheetEvent)).toBe(true);
    expect(recalc).toHaveBeenCalledTimes(2);

    const iterativeEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(iterativeEvent, 'target', { value: iterative });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(iterativeEvent)).toBe(true);
    expect(openIterativeDialog).toHaveBeenCalledTimes(1);

    tb.dispose();
  });

  it('clears formula audit arrows by kind through the Clear Arrows dropdown', () => {
    const precedent = {
      kind: 'precedent' as const,
      from: { sheet: 0, row: 0, col: 0 },
      to: { sheet: 0, row: 0, col: 2 },
    };
    const dependent = {
      kind: 'dependent' as const,
      from: { sheet: 0, row: 0, col: 1 },
      to: { sheet: 0, row: 0, col: 2 },
    };
    mutators.addTrace(sheet.instance.store, precedent);
    mutators.addTrace(sheet.instance.store, dependent);
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('formulas');

    const clearArrowsButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="clearArrows"]',
    );
    expect(clearArrowsButton).toBeTruthy();
    clearArrowsButton?.click();
    expect(host.querySelectorAll('#menu-clear-arrows .fc-tb__menu-item--iconic').length).toBe(3);
    const clearAllButton = host.querySelector<HTMLButtonElement>(
      '[data-formula-audit-action="clear-all"]',
    );
    const clearPrecedentsButton = host.querySelector<HTMLButtonElement>(
      '[data-formula-audit-action="clear-precedents"]',
    );
    const clearDependentsButton = host.querySelector<HTMLButtonElement>(
      '[data-formula-audit-action="clear-dependents"]',
    );
    expect(clearAllButton?.disabled).toBe(false);
    expect(clearPrecedentsButton).toBeTruthy();
    expect(clearPrecedentsButton?.disabled).toBe(false);
    expect(clearDependentsButton?.disabled).toBe(false);
    const clearPrecedentsEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(clearPrecedentsEvent, 'target', { value: clearPrecedentsButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(clearPrecedentsEvent)).toBe(true);
    expect(sheet.instance.store.getState().traces.items).toEqual([dependent]);

    clearArrowsButton?.click();
    expect(clearAllButton?.disabled).toBe(false);
    expect(clearPrecedentsButton?.disabled).toBe(true);
    expect(clearDependentsButton?.disabled).toBe(false);
    expect(clearDependentsButton).toBeTruthy();
    const clearDependentsEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(clearDependentsEvent, 'target', { value: clearDependentsButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(clearDependentsEvent)).toBe(true);
    expect(sheet.instance.store.getState().traces.items).toEqual([]);

    clearArrowsButton?.click();
    expect(clearAllButton?.disabled).toBe(true);
    expect(clearPrecedentsButton?.disabled).toBe(true);
    expect(clearDependentsButton?.disabled).toBe(true);

    tb.dispose();
  });

  it('runs Error Checking from the primary button and keeps the menu secondary', async () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('formulas');

    const errorButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="errorChecking"]',
    );
    expect(errorButton).toBeTruthy();
    expect(errorButton?.dataset.ribbonActivation).toBe('splitPrimary');
    errorButton?.click();
    await Promise.resolve();
    await Promise.resolve();

    expect(host.querySelector<HTMLDivElement>('#menu-error-checking')?.hidden).toBe(true);
    expect(document.body.querySelector<HTMLElement>('.fc-tb__dlg')?.textContent).toContain(
      'Error Checking',
    );
    document.body.querySelector<HTMLButtonElement>('.fc-tb__dlg .fc-fmtdlg__btn--primary')?.click();

    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'errorChecking', menuId: 'menu-error-checking' },
      errorButton as HTMLButtonElement,
    );
    expect(host.querySelector<HTMLDivElement>('#menu-error-checking')?.hidden).toBe(false);
    expect(host.querySelectorAll('#menu-error-checking .fc-tb__menu-item--iconic').length).toBe(3);
    expect(
      host.querySelector<HTMLButtonElement>('[data-formula-audit-action="trace-error"]'),
    ).toBeTruthy();
    expect(
      host
        .querySelector<HTMLButtonElement>('[data-formula-audit-action="trace-error"]')
        ?.getAttribute('aria-disabled'),
    ).toBe('true');
    expect(
      host.querySelector<HTMLButtonElement>('[data-formula-audit-action="trace-error"]')?.dataset
        .menuDisabledReason,
    ).toBe('The active cell does not contain a formula error.');
    expect(
      host
        .querySelector<HTMLButtonElement>('[data-formula-audit-action="ignore-error"]')
        ?.getAttribute('aria-disabled'),
    ).toBe('true');

    const errorAddr = { sheet: 0, row: 1, col: 1 };
    sheet.workbook.setFormula(errorAddr, '=1/0');
    sheet.instance.store.setState((state) => {
      const cells = new Map(state.data.cells);
      cells.set(addrKey(errorAddr), {
        value: { kind: 'error', code: 7, text: '#DIV/0!' },
        formula: '=1/0',
      });
      return { ...state, data: { ...state.data, cells } };
    });
    mutators.setActive(sheet.instance.store, errorAddr);
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 });
    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'errorChecking', menuId: 'menu-error-checking' },
      errorButton as HTMLButtonElement,
    );
    expect(
      host
        .querySelector<HTMLButtonElement>('[data-formula-audit-action="trace-error"]')
        ?.getAttribute('aria-disabled'),
    ).toBe('false');
    expect(
      host.querySelector<HTMLButtonElement>('[data-formula-audit-action="trace-error"]')?.dataset
        .menuDisabledReason,
    ).toBeUndefined();
    expect(
      host
        .querySelector<HTMLButtonElement>('[data-formula-audit-action="ignore-error"]')
        ?.getAttribute('aria-disabled'),
    ).toBe('false');

    tb.dispose();
  });

  it('opens Watch Window from the primary button and keeps Add/Delete secondary', () => {
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 0, col: 0 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('formulas');

    const watchButton = host.querySelector<HTMLButtonElement>('[data-ribbon-command="watch"]');
    expect(watchButton).toBeTruthy();
    expect(watchButton?.dataset.ribbonActivation).toBe('splitPrimary');
    watchButton?.click();

    expect(host.querySelector<HTMLDivElement>('#menu-watch-formulas')?.hidden).toBe(true);
    expect(sheet.instance.store.getState().ui.watchPanelOpen).toBe(true);

    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'watch', menuId: 'menu-watch-formulas' },
      watchButton as HTMLButtonElement,
    );
    expect(host.querySelector<HTMLDivElement>('#menu-watch-formulas')?.hidden).toBe(false);
    expect(host.querySelectorAll('#menu-watch-formulas .fc-tb__menu-item--iconic').length).toBe(4);
    const addButton = host.querySelector<HTMLButtonElement>('[data-watch-action="add"]');
    expect(addButton).toBeTruthy();
    const deleteButton = host.querySelector<HTMLButtonElement>('[data-watch-action="delete"]');
    const deleteAllButton = host.querySelector<HTMLButtonElement>(
      '[data-watch-action="delete-all"]',
    );
    expect(deleteButton).toBeTruthy();
    expect(deleteButton?.disabled).toBe(true);
    expect(deleteButton?.dataset.menuDisabledReason).toBe('The active cell is not being watched.');
    expect(deleteAllButton?.disabled).toBe(true);
    expect(deleteAllButton?.dataset.menuDisabledReason).toBe('There are no watches to delete.');
    const addEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(addEvent, 'target', { value: addButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(addEvent)).toBe(true);
    expect(sheet.instance.store.getState().watch.watches).toContainEqual({
      sheet: 0,
      row: 0,
      col: 0,
    });
    expect(sheet.instance.store.getState().ui.watchPanelOpen).toBe(true);

    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'watch', menuId: 'menu-watch-formulas' },
      watchButton as HTMLButtonElement,
    );
    expect(deleteButton?.disabled).toBe(false);
    expect(deleteButton?.getAttribute('aria-disabled')).toBe('false');
    expect(deleteButton?.dataset.menuDisabledReason).toBeUndefined();
    expect(deleteAllButton?.disabled).toBe(false);
    expect(deleteAllButton?.dataset.menuDisabledReason).toBeUndefined();
    const deleteEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(deleteEvent, 'target', { value: deleteButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(deleteEvent)).toBe(true);
    expect(sheet.instance.store.getState().watch.watches).toEqual([]);

    mutators.addWatch(sheet.instance.store, { sheet: 0, row: 0, col: 0 });

    mutators.setActive(sheet.instance.store, { sheet: 0, row: 1, col: 0 });
    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'watch', menuId: 'menu-watch-formulas' },
      watchButton as HTMLButtonElement,
    );
    expect(deleteButton?.disabled).toBe(true);
    expect(deleteButton?.dataset.menuDisabledReason).toBe('The active cell is not being watched.');
    expect(deleteAllButton?.disabled).toBe(false);
    const deleteAllEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(deleteAllEvent, 'target', { value: deleteAllButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(deleteAllEvent)).toBe(true);
    expect(sheet.instance.store.getState().watch.watches).toEqual([]);

    tb.dispose();
  });

  it('opens External Links from primary click and keeps hyperlink actions secondary', () => {
    const originalOpen = window.open;
    const open = vi.fn();
    window.open = open;
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 0, col: 0 });
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    setHyperlink(
      sheet.instance.store,
      { sheet: 0, row: 0, col: 0 },
      'https://example.test',
      sheet.workbook,
    );
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('data');

    const linksButton = host.querySelector<HTMLButtonElement>('[data-ribbon-command="linksData"]');
    expect(linksButton).toBeTruthy();
    linksButton?.click();
    const linksDialog = document.body.querySelector<HTMLElement>('.fc-extlinkdlg');
    expect(linksDialog?.hidden).toBe(false);
    linksDialog
      ?.querySelector<HTMLButtonElement>('.fc-extlinkdlg__close')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(linksDialog?.hidden).toBe(true);

    tb.dropdownsApi?.openDynamicRibbonDropdown({
      command: 'linksData',
      menuId: 'menu-links-data',
    });
    const linksMenu = host.querySelector<HTMLElement>('#menu-links-data');
    expect(linksMenu?.querySelectorAll('.fc-tb__menu-item--iconic').length).toBe(4);
    const openButton = linksMenu?.querySelector<HTMLButtonElement>('[data-link-action="open"]');
    const clearButton = linksMenu?.querySelector<HTMLButtonElement>('[data-link-action="clear"]');
    expect(openButton?.disabled).toBe(false);
    expect(clearButton).toBeTruthy();
    expect(clearButton?.disabled).toBe(false);
    const openEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(openEvent, 'target', { value: openButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(openEvent)).toBe(true);
    expect(open).toHaveBeenCalledWith('https://example.test', '_blank', 'noopener,noreferrer');

    tb.dropdownsApi?.openDynamicRibbonDropdown({
      command: 'linksData',
      menuId: 'menu-links-data',
    });
    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: clearButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);

    expect(hyperlinkAt(sheet.instance.store.getState(), { sheet: 0, row: 0, col: 0 })).toBeNull();

    tb.dropdownsApi?.openDynamicRibbonDropdown({
      command: 'linksData',
      menuId: 'menu-links-data',
    });
    expect(openButton?.disabled).toBe(true);
    expect(openButton?.dataset.menuDisabledReason).toBe(
      'The active cell does not contain a hyperlink.',
    );
    expect(clearButton?.disabled).toBe(true);
    expect(clearButton?.dataset.menuDisabledReason).toBe(
      'The active cell does not contain a hyperlink.',
    );

    tb.dispose();
    window.open = originalOpen;
  });

  it('opens Name Manager from primary click and keeps Define Name secondary', async () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('formulas');

    const namesButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="namedRanges"]',
    );
    expect(namesButton).toBeTruthy();
    namesButton?.click();
    await new Promise((resolve) => requestAnimationFrame(resolve));
    const manager = document.body.querySelector<HTMLElement>('.fc-namedlg');
    expect(manager?.hidden).toBe(false);
    expect(manager?.querySelector<HTMLElement>('.fc-namedlg__list')).toBeTruthy();

    Array.from(manager?.querySelectorAll<HTMLButtonElement>('button') ?? [])
      .find((button) => button.textContent === 'Close')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(manager?.hidden).toBe(true);

    tb.dropdownsApi?.openDynamicRibbonDropdown({
      command: 'namedRanges',
      menuId: 'menu-defined-names',
    });
    const namesMenu = host.querySelector<HTMLElement>('#menu-defined-names');
    expect(namesMenu?.querySelectorAll('.fc-tb__menu-item--iconic').length).toBe(7);
    expect(
      namesMenu?.querySelector<HTMLButtonElement>('[data-defined-name-action="use-formula"]')
        ?.disabled,
    ).toBe(true);
    for (const createButton of namesMenu?.querySelectorAll<HTMLButtonElement>(
      '[data-defined-name-action^="create-"]',
    ) ?? []) {
      expect(createButton.disabled).toBe(!sheet.workbook.capabilities.definedNameMutate);
    }
    vi.spyOn(sheet.workbook, 'definedNames').mockImplementation(function* () {
      yield { name: 'TaxRate', formula: '=Sheet1!$A$1', localSheetId: -1 };
      yield { name: 'NetSales', formula: '=Sheet1!$B$1', localSheetId: -1 };
    });
    tb.dropdownsApi?.openDynamicRibbonDropdown({
      command: 'namedRanges',
      menuId: 'menu-defined-names',
    });
    const useFormulaButton = namesMenu?.querySelector<HTMLButtonElement>(
      '[data-defined-name-action="use-formula"]',
    );
    expect(useFormulaButton?.disabled).toBe(false);
    const useFormulaEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(useFormulaEvent, 'target', { value: useFormulaButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(useFormulaEvent)).toBe(true);
    await Promise.resolve();
    const useFormulaDialog = document.body.querySelector<HTMLElement>('.fc-tb__dlg');
    expect(useFormulaDialog?.textContent).toContain('Use in Formula');
    expect(useFormulaDialog?.textContent).toContain('TaxRate');
    expect(useFormulaDialog?.textContent).toContain('NetSales');
    expect(sheet.workbook.cellFormula({ sheet: 0, row: 0, col: 0 })).toBeNull();
    const netSalesRadio =
      useFormulaDialog?.querySelector<HTMLInputElement>('input[value="NetSales"]');
    expect(netSalesRadio).toBeTruthy();
    netSalesRadio?.click();
    useFormulaDialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await Promise.resolve();
    expect(sheet.workbook.cellFormula({ sheet: 0, row: 0, col: 0 })).toBe('=NetSales');

    tb.dropdownsApi?.openDynamicRibbonDropdown({
      command: 'namedRanges',
      menuId: 'menu-defined-names',
    });
    const defineButton = namesMenu?.querySelector<HTMLButtonElement>(
      '[data-defined-name-action="define"]',
    );
    expect(defineButton).toBeTruthy();
    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: defineButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);
    await new Promise((resolve) => requestAnimationFrame(resolve));

    expect(document.body.querySelector<HTMLElement>('.fc-namedlg')?.hidden).toBe(false);

    tb.dispose();
  });
});
