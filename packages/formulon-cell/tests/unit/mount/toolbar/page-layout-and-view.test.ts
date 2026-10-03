import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { Spreadsheet } from '../../../../src/mount.js';
import { getPageSetup, mutators } from '../../../../src/store/store.js';
import {
  HOME_MIXED_LAYOUT_GROUP_VARIANTS,
  HOME_STACKED_LAYOUT_GROUP_VARIANTS,
  HOME_TILE_LAYOUT_GROUP_VARIANTS,
} from '../../../../src/toolbar/ribbon-model.js';
import { type MountedStubSheet, mountStubSheet } from '../../../test-utils/mount.js';
import { stubHelpers, waitFor } from './fixtures.js';

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

  it('opens the View Zoom dialog from the shared default hook', async () => {
    const refreshZoom = vi.fn();
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      helpers: stubHelpers(),
      refreshZoom,
    });
    tb.setActiveTab('view');

    const zoomButton = host.querySelector<HTMLButtonElement>('[data-ribbon-command="zoomDialog"]');
    expect(zoomButton).toBeTruthy();
    expect(zoomButton?.dataset.ribbonActivation).toBe('dialog');
    zoomButton?.click();

    const dialog = document.body.querySelector<HTMLElement>('.fc-tb__dlg');
    const input = dialog?.querySelector<HTMLInputElement>('input');
    expect(input).toBeTruthy();
    if (!input) throw new Error('Expected Zoom dialog input.');
    expect(input.value).toBe('100');
    input.value = '401';
    dialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    expect(dialog?.textContent).toContain('Enter a zoom percentage from 50 to 400.');
    input.value = '125';
    dialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await Promise.resolve();
    await Promise.resolve();

    expect(sheet.instance.store.getState().viewport.zoom).toBe(1.25);
    expect(refreshZoom).toHaveBeenCalled();

    tb.dispose();
  });

  it('sets, adds, and clears the current selection through the Print Area dropdown', () => {
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 1, c0: 1, r1: 2, c1: 3 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('pageLayout');

    const printAreaButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="printArea"]',
    );
    expect(printAreaButton).toBeTruthy();
    printAreaButton?.click();
    const setPrintAreaButton = host.querySelector<HTMLButtonElement>(
      '[data-print-area-action="set"]',
    );
    const initialAddPrintAreaButton = host.querySelector<HTMLButtonElement>(
      '[data-print-area-action="add"]',
    );
    const initialClearPrintAreaButton = host.querySelector<HTMLButtonElement>(
      '[data-print-area-action="clear"]',
    );
    expect(setPrintAreaButton).toBeTruthy();
    expect(initialAddPrintAreaButton?.disabled).toBe(true);
    expect(initialAddPrintAreaButton?.getAttribute('aria-disabled')).toBe('true');
    expect(initialAddPrintAreaButton?.dataset.menuDisabledReason).toBe(
      'No print area has been set.',
    );
    expect(initialClearPrintAreaButton?.disabled).toBe(true);
    expect(initialClearPrintAreaButton?.dataset.menuDisabledReason).toBe(
      'No print area has been set.',
    );
    expect(
      host.querySelector<HTMLElement>('#menu-print-area .fc-tb__menu-icon--print-area-set'),
    ).toBeTruthy();
    const setEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(setEvent, 'target', { value: setPrintAreaButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(setEvent)).toBe(true);

    expect(getPageSetup(sheet.instance.store.getState(), 0).printArea).toBe('B2:D3');

    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 5, c0: 5, r1: 5, c1: 6 });
    printAreaButton?.click();
    const addPrintAreaButton = host.querySelector<HTMLButtonElement>(
      '[data-print-area-action="add"]',
    );
    expect(addPrintAreaButton).toBeTruthy();
    expect(addPrintAreaButton?.disabled).toBe(false);
    expect(addPrintAreaButton?.dataset.menuDisabledReason).toBeUndefined();
    expect(
      host
        .querySelector<HTMLButtonElement>('[data-print-area-action="clear"]')
        ?.getAttribute('aria-disabled'),
    ).toBe('false');
    const addEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(addEvent, 'target', { value: addPrintAreaButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(addEvent)).toBe(true);

    expect(getPageSetup(sheet.instance.store.getState(), 0).printArea).toBe('B2:D3,F6:G6');

    printAreaButton?.click();
    const clearPrintAreaButton = host.querySelector<HTMLButtonElement>(
      '[data-print-area-action="clear"]',
    );
    expect(clearPrintAreaButton).toBeTruthy();
    const clearEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(clearEvent, 'target', { value: clearPrintAreaButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(clearEvent)).toBe(true);

    expect(getPageSetup(sheet.instance.store.getState(), 0).printArea).toBeUndefined();

    printAreaButton?.click();
    expect(host.querySelector<HTMLButtonElement>('[data-print-area-action="add"]')?.disabled).toBe(
      true,
    );
    expect(
      host.querySelector<HTMLButtonElement>('[data-print-area-action="clear"]')?.disabled,
    ).toBe(true);

    tb.dispose();
  });

  it('renders Page Theme as a visual gallery and applies the selected theme', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('pageLayout');

    const themeButton = host.querySelector<HTMLButtonElement>('[data-ribbon-command="pageTheme"]');
    expect(themeButton).toBeTruthy();
    expect(themeButton?.dataset.ribbonActivation).toBe('gallery');
    themeButton?.click();
    const menu = host.querySelector<HTMLElement>('#menu-page-theme');
    expect(menu?.classList.contains('fc-tb__menu--visual')).toBe(true);
    expect(menu?.querySelectorAll('.fc-tb__visual-tile')).toHaveLength(3);

    const paperButton = host.querySelector<HTMLButtonElement>('[data-page-theme-action="paper"]');
    const inkButton = host.querySelector<HTMLButtonElement>('[data-page-theme-action="ink"]');
    expect(paperButton?.getAttribute('role')).toBe('menuitemradio');
    expect(paperButton?.getAttribute('aria-checked')).toBe('true');
    expect(paperButton?.classList.contains('fc-tb__visual-tile--active')).toBe(true);
    expect(inkButton?.getAttribute('aria-checked')).toBe('false');
    expect(inkButton).toBeTruthy();
    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: inkButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);
    // The Ink tile applies the grid's `ink` ThemeName directly — tiles and the
    // grid share the single paper/ink/contrast vocabulary.
    expect(sheet.instance.store.getState().ui.theme).toBe('ink');

    themeButton?.click();
    expect(paperButton?.getAttribute('aria-checked')).toBe('false');
    expect(inkButton?.getAttribute('aria-checked')).toBe('true');
    expect(inkButton?.classList.contains('fc-tb__visual-tile--active')).toBe(true);
    const contrastButton = host.querySelector<HTMLButtonElement>(
      '[data-page-theme-action="contrast"]',
    );
    expect(contrastButton).toBeTruthy();
    const contrastEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(contrastEvent, 'target', { value: contrastButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(contrastEvent)).toBe(true);
    expect(sheet.instance.store.getState().ui.theme).toBe('contrast');

    tb.dispose();
  });

  it('disables Page Layout Arrange ordering actions until an illustration exists', () => {
    const openWorkbookObjects = vi.spyOn(sheet.instance, 'openWorkbookObjects');
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('pageLayout');

    const arrangeButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="arrangeObjectsPageLayout"]',
    );
    expect(arrangeButton).toBeTruthy();
    arrangeButton?.click();
    const bringForward = host.querySelector<HTMLButtonElement>(
      '#menu-arrange-objects [data-arrange-action="bring-forward"]',
    );
    const selectionPane = host.querySelector<HTMLButtonElement>(
      '#menu-arrange-objects [data-arrange-action="selection-pane"]',
    );
    expect(bringForward?.disabled).toBe(true);
    expect(bringForward?.getAttribute('aria-disabled')).toBe('true');
    expect(bringForward?.dataset.menuDisabledReason).toBe('Select an object to arrange.');
    expect(selectionPane?.disabled).toBe(false);

    const disabledEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(disabledEvent, 'target', { value: bringForward });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(disabledEvent)).toBe(true);
    expect(openWorkbookObjects).not.toHaveBeenCalled();

    if (!bringForward) throw new Error('Expected Bring Forward item.');
    bringForward.disabled = false;
    bringForward.setAttribute('aria-disabled', 'true');
    const ariaDisabledEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(ariaDisabledEvent, 'target', { value: bringForward });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(ariaDisabledEvent)).toBe(true);
    expect(openWorkbookObjects).not.toHaveBeenCalled();

    mutators.upsertIllustration(sheet.instance.store, {
      id: 'shape-back',
      kind: 'shape',
      shape: 'rectangle',
      sheet: 0,
    });
    mutators.upsertIllustration(sheet.instance.store, {
      id: 'shape-front',
      kind: 'shape',
      shape: 'arrow',
      sheet: 0,
    });

    tb.setActiveTab('pageLayout');
    host
      .querySelector<HTMLButtonElement>('[data-ribbon-command="arrangeObjectsPageLayout"]')
      ?.click();
    const frontBringForward = host.querySelector<HTMLButtonElement>(
      '#menu-arrange-objects [data-arrange-action="bring-forward"]',
    );
    const frontSendBackward = host.querySelector<HTMLButtonElement>(
      '#menu-arrange-objects [data-arrange-action="send-backward"]',
    );
    expect(frontBringForward?.disabled).toBe(true);
    expect(frontBringForward?.dataset.menuDisabledReason).toBe(
      'The selected object is already at the front.',
    );
    expect(frontSendBackward?.disabled).toBe(false);

    const selectedBack = document.createElement('div');
    selectedBack.className = 'fc-illustration';
    selectedBack.dataset.illustrationId = 'shape-back';
    selectedBack.setAttribute('aria-selected', 'true');
    sheet.instance.host.appendChild(selectedBack);
    host
      .querySelector<HTMLButtonElement>('[data-ribbon-command="arrangeObjectsPageLayout"]')
      ?.click();
    const backBringForward = host.querySelector<HTMLButtonElement>(
      '#menu-arrange-objects [data-arrange-action="bring-forward"]',
    );
    const backSendBackward = host.querySelector<HTMLButtonElement>(
      '#menu-arrange-objects [data-arrange-action="send-backward"]',
    );
    expect(backBringForward?.disabled).toBe(false);
    expect(backBringForward?.getAttribute('aria-disabled')).toBe('false');
    expect(backBringForward?.dataset.menuDisabledReason).toBeUndefined();
    expect(backSendBackward?.disabled).toBe(true);
    expect(backSendBackward?.getAttribute('aria-disabled')).toBe('true');
    expect(backSendBackward?.dataset.menuDisabledReason).toBe(
      'The selected object is already at the back.',
    );

    tb.dispose();
    openWorkbookObjects.mockRestore();
  });

  it('updates manual page breaks through the Excel-style Page Layout Breaks dropdown', () => {
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 4, c0: 2, r1: 4, c1: 2 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('pageLayout');

    const breaksButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="pageBreaks"]',
    );
    expect(breaksButton).toBeTruthy();
    breaksButton?.click();
    expect(host.querySelectorAll('#menu-page-breaks .fc-tb__menu-item--iconic')).toHaveLength(3);
    const insertButton = host.querySelector<HTMLButtonElement>('[data-page-break-action="insert"]');
    const removeButton = host.querySelector<HTMLButtonElement>('[data-page-break-action="remove"]');
    const resetButton = host.querySelector<HTMLButtonElement>(
      '[data-page-break-action="reset-all"]',
    );
    expect(insertButton).toBeTruthy();
    expect(removeButton).toBeTruthy();
    expect(
      host.querySelector<HTMLElement>('#menu-page-breaks .fc-tb__menu-icon--break-page'),
    ).toBeTruthy();
    expect(insertButton?.textContent).toContain('Insert Page Break');
    expect(removeButton?.textContent).toContain('Remove Page Break');
    expect(insertButton?.disabled).toBe(false);
    expect(removeButton?.disabled).toBe(true);
    expect(removeButton?.getAttribute('aria-disabled')).toBe('true');
    expect(removeButton?.dataset.menuDisabledReason).toBe(
      'There is no manual page break at the selection.',
    );
    expect(resetButton?.disabled).toBe(true);
    expect(resetButton?.dataset.menuDisabledReason).toBe(
      'There are no manual page breaks to reset.',
    );

    const insertEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(insertEvent, 'target', { value: insertButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(insertEvent)).toBe(true);

    expect(getPageSetup(sheet.instance.store.getState(), 0).manualPageBreakRows).toEqual([4]);
    expect(getPageSetup(sheet.instance.store.getState(), 0).manualPageBreakCols).toEqual([2]);

    breaksButton?.click();
    const enabledRemoveButton = host.querySelector<HTMLButtonElement>(
      '[data-page-break-action="remove"]',
    );
    const enabledResetButton = host.querySelector<HTMLButtonElement>(
      '[data-page-break-action="reset-all"]',
    );
    expect(enabledRemoveButton?.disabled).toBe(false);
    expect(enabledRemoveButton?.getAttribute('aria-disabled')).toBe('false');
    expect(enabledRemoveButton?.dataset.menuDisabledReason).toBeUndefined();
    expect(enabledResetButton?.disabled).toBe(false);
    expect(enabledResetButton?.dataset.menuDisabledReason).toBeUndefined();
    const removeEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(removeEvent, 'target', { value: enabledRemoveButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(removeEvent)).toBe(true);
    expect(getPageSetup(sheet.instance.store.getState(), 0).manualPageBreakRows).toBeUndefined();
    expect(getPageSetup(sheet.instance.store.getState(), 0).manualPageBreakCols).toBeUndefined();

    breaksButton?.click();
    expect(enabledRemoveButton?.disabled).toBe(true);
    expect(enabledRemoveButton?.dataset.menuDisabledReason).toBe(
      'There is no manual page break at the selection.',
    );
    expect(enabledResetButton?.disabled).toBe(true);

    const insertAgainEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(insertAgainEvent, 'target', { value: insertButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(insertAgainEvent)).toBe(true);
    breaksButton?.click();
    const resetAfterRemoveColButton = host.querySelector<HTMLButtonElement>(
      '[data-page-break-action="reset-all"]',
    );
    expect(resetAfterRemoveColButton?.disabled).toBe(false);
    const resetEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(resetEvent, 'target', { value: resetAfterRemoveColButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(resetEvent)).toBe(true);

    expect(getPageSetup(sheet.instance.store.getState(), 0).manualPageBreakRows).toBeUndefined();
    expect(getPageSetup(sheet.instance.store.getState(), 0).manualPageBreakCols).toBeUndefined();

    tb.dispose();
  });

  it('sets and deletes sheet background from the Excel-style primary button', async () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('pageLayout');

    const backgroundButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="sheetBackground"]',
    );
    expect(backgroundButton).toBeTruthy();
    expect(backgroundButton?.dataset.ribbonActivation).toBe('primaryAction');
    expect(backgroundButton?.dataset.ribbonMenuId).toBeUndefined();
    expect(backgroundButton?.getAttribute('aria-haspopup')).toBeNull();
    expect(backgroundButton?.textContent).toContain('Background');
    expect(backgroundButton?.getAttribute('aria-label')).toBe('Background');
    backgroundButton?.click();

    const input = document.body.querySelector<HTMLInputElement>('input[type="file"]');
    expect(input).toBeTruthy();
    if (!input) throw new Error('Sheet background file picker input was not rendered');
    const file = new File(['background'], 'background.png', { type: 'image/png' });
    Object.defineProperty(input, 'files', { value: [file], configurable: true });
    input.dispatchEvent(new Event('change', { bubbles: true }));

    await waitFor(() =>
      Boolean(sheet.instance.store.getState().ui.sheetBackgroundImages.get(0)?.startsWith('data:')),
    );
    expect(document.body.querySelector('input[type="file"]')).toBeNull();
    expect(backgroundButton?.textContent).toContain('Delete Background');
    expect(backgroundButton?.getAttribute('aria-label')).toBe('Delete Background');
    expect(host.querySelector('#menu-sheet-background')).toBeNull();

    mutators.setSheetBackgroundImage(
      sheet.instance.store,
      0,
      'https://example.test/second-background.png',
    );
    backgroundButton?.click();
    expect(sheet.instance.store.getState().ui.sheetBackgroundImages.has(0)).toBe(false);
    expect(backgroundButton?.textContent).toContain('Background');
    expect(backgroundButton?.getAttribute('aria-label')).toBe('Background');
    expect(document.body.querySelector('.fc-tb__dlg')).toBeNull();

    tb.dispose();
  });

  it('opens Page Setup Sheet tab from Print Titles primary click and keeps secondary actions', () => {
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 1, c0: 2, r1: 3, c1: 4 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('pageLayout');

    const printTitlesButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="printTitles"]',
    );
    expect(printTitlesButton).toBeTruthy();
    expect(printTitlesButton?.dataset.ribbonActivation).toBe('dialog');
    expect(printTitlesButton?.dataset.ribbonMenuId).toBeUndefined();
    expect(printTitlesButton?.getAttribute('aria-haspopup')).toBeNull();
    printTitlesButton?.click();
    const pageSetupDialog = document.body.querySelector<HTMLElement>('.fc-pgsetup');
    expect(pageSetupDialog?.hidden).toBe(false);
    expect(
      pageSetupDialog
        ?.querySelector<HTMLButtonElement>('[data-pgsetup-tab="sheet"]')
        ?.getAttribute('aria-selected'),
    ).toBe('true');
    pageSetupDialog
      ?.querySelector<HTMLButtonElement>('.fc-fmtdlg__close')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(pageSetupDialog?.hidden).toBe(true);
    expect(host.querySelector('#menu-print-titles')).toBeNull();

    tb.dispose();
  });

  it('keeps Home Merge and Find rendering in the audited Excel 365 ribbon layout', () => {
    const helpers = stubHelpers();
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: {
        ...helpers,
        createIcon: (name) => {
          const icon = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
          icon.setAttribute('class', 'fc-tb__rb-icon');
          icon.dataset.icon = name;
          return icon;
        },
      },
    });

    const mergeButton = host.querySelector<HTMLButtonElement>('[data-ribbon-command="merge"]');
    expect(mergeButton).toBeTruthy();
    expect(mergeButton?.closest('.fc-tb__ribbon-group')?.classList).toContain(
      'fc-tb__ribbon-group--alignment',
    );
    expect(mergeButton?.classList.contains('fc-tb__rb--wide')).toBe(false);
    expect(mergeButton?.dataset.ribbonActivation).toBe('splitPrimary');
    expect(mergeButton?.querySelector('.fc-tb__rb-icon')?.getAttribute('data-icon')).toBe('merge');
    expect(
      Array.from(mergeButton?.querySelectorAll('span') ?? []).filter(
        (span) => !span.classList.contains('fc-tb__rb-split-chevron'),
      ),
    ).toEqual([]);

    const findButton = host.querySelector<HTMLButtonElement>('[data-ribbon-command="findHome"]');
    expect(findButton).toBeTruthy();
    expect(findButton?.closest('.fc-tb__ribbon-group')?.classList).toContain(
      'fc-tb__ribbon-group--editing',
    );
    expect(findButton?.classList.contains('fc-tb__rb--wide')).toBe(true);
    expect(findButton?.dataset.ribbonActivation).toBe('dropdown');
    expect(findButton?.dataset.ribbonMenuId).toBe('menu-find-select');
    expect(findButton?.querySelector('.fc-tb__rb-icon')?.getAttribute('data-icon')).toBe('find');
    expect(findButton?.querySelector('span')?.textContent).toBe('Find & Select');
    expect(findButton?.querySelector('.fc-tb__rb-split-chevron')).toBeTruthy();

    tb.dispose();
  });

  it('projects Home dense groups through shared layout classes', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, { helpers: stubHelpers() });

    for (const variant of HOME_TILE_LAYOUT_GROUP_VARIANTS) {
      const group = host.querySelector<HTMLElement>(`.fc-tb__ribbon-group--${variant}`);
      expect(group, variant).toBeTruthy();
      expect(group?.classList.contains('fc-tb__ribbon-group--tiles'), variant).toBe(true);
      const nonWideCommands = Array.from(
        group?.querySelectorAll<HTMLButtonElement>('[data-ribbon-command]') ?? [],
      )
        .filter((button) => !button.classList.contains('fc-tb__rb--wide'))
        .map((button) => button.dataset.ribbonCommand);
      expect(nonWideCommands, variant).toEqual([]);
    }
    for (const variant of HOME_STACKED_LAYOUT_GROUP_VARIANTS) {
      const group = host.querySelector<HTMLElement>(`.fc-tb__ribbon-group--${variant}`);
      expect(group, variant).toBeTruthy();
      expect(group?.classList.contains('fc-tb__ribbon-group--stacked'), variant).toBe(true);
      expect(
        Array.from(group?.querySelectorAll<HTMLButtonElement>('[data-ribbon-command]') ?? []).map(
          (button) => button.classList.contains('fc-tb__rb--stacked'),
        ),
        variant,
      ).toEqual([true, true, true]);
    }
    for (const variant of HOME_MIXED_LAYOUT_GROUP_VARIANTS) {
      const group = host.querySelector<HTMLElement>(`.fc-tb__ribbon-group--${variant}`);
      expect(group, variant).toBeTruthy();
      expect(group?.classList.contains('fc-tb__ribbon-group--mixed'), variant).toBe(true);
      expect(
        Array.from(group?.querySelectorAll<HTMLButtonElement>('[data-ribbon-command]') ?? []).map(
          (button) => button.classList.contains('fc-tb__rb--stacked'),
        ),
        variant,
      ).toEqual([true, true, true, false, false]);
    }

    tb.dispose();
  });

  it('applies Freeze Panes dropdown actions through default dynamic wiring', () => {
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 3, col: 2 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('view');

    const freezeButton = host.querySelector<HTMLButtonElement>('[data-ribbon-command="freeze"]');
    expect(freezeButton).toBeTruthy();
    if (freezeButton) {
      Object.defineProperty(freezeButton, 'getBoundingClientRect', {
        value: () =>
          ({
            x: 120,
            y: 32,
            width: 40,
            height: 28,
            top: 32,
            right: 160,
            bottom: 60,
            left: 120,
            toJSON: () => ({}),
          }) as DOMRect,
      });
    }
    freezeButton?.click();
    const freezeMenu = host.querySelector<HTMLElement>('#menu-freeze');
    expect(freezeMenu?.style.position).toBe('fixed');
    expect(freezeMenu?.style.left).toBe('120px');
    expect(freezeMenu?.style.top).toBe('63px');
    document.body.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
    expect(freezeMenu?.hidden).toBe(true);
    freezeButton?.click();
    const freezeItems = Array.from(
      host.querySelectorAll<HTMLButtonElement>('#menu-freeze .fc-tb__menu-item--iconic'),
    );
    expect(freezeItems).toHaveLength(3);
    expect(freezeItems.map((item) => item.textContent)).toEqual([
      'Freeze Panes',
      'Freeze Top Row',
      'Freeze First Column',
    ]);
    expect(host.querySelector('#menu-freeze [data-freeze="off"]')).toBeNull();
    const selectionFreeze = host.querySelector<HTMLButtonElement>(
      '#menu-freeze [data-freeze="selection"]',
    );
    expect(selectionFreeze).toBeTruthy();
    const freezeEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(freezeEvent, 'target', { value: selectionFreeze });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(freezeEvent)).toBe(true);
    expect(sheet.instance.store.getState().layout.freezeRows).toBe(3);
    expect(sheet.instance.store.getState().layout.freezeCols).toBe(2);

    freezeButton?.click();
    const unfreeze = host.querySelector<HTMLButtonElement>('#menu-freeze [data-freeze="off"]');
    expect(unfreeze).toBeTruthy();
    expect(host.querySelectorAll('#menu-freeze .fc-tb__menu-item--iconic').length).toBe(3);
    expect(unfreeze?.textContent).toBe('Unfreeze Panes');
    expect(unfreeze?.disabled).toBe(false);
    expect(unfreeze?.getAttribute('aria-disabled')).toBe('false');
    const unfreezeEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(unfreezeEvent, 'target', { value: unfreeze });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(unfreezeEvent)).toBe(true);
    expect(sheet.instance.store.getState().layout.freezeRows).toBe(0);
    expect(sheet.instance.store.getState().layout.freezeCols).toBe(0);

    tb.dispose();
  });
});
