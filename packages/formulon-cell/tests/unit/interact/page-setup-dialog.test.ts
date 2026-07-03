import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { setMarginPreset } from '../../../src/commands/page-setup.js';
import { defaultStrings, en } from '../../../src/i18n/strings.js';
import { attachPageSetupDialog } from '../../../src/interact/page-setup-dialog.js';
import {
  createSpreadsheetStore,
  defaultPageSetup,
  getPageSetup,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');

const dialog = (): HTMLElement | null => document.querySelector<HTMLElement>('.fc-pgsetup');
const selects = (): HTMLSelectElement[] =>
  Array.from(
    document.querySelectorAll<HTMLSelectElement>('.fc-pgsetup__select:not([data-pgsetup-printer])'),
  );
const marginInputs = (): HTMLInputElement[] =>
  Array.from(document.querySelectorAll<HTMLInputElement>('.fc-pgsetup__margins input'));
const pageMarginInputs = (): HTMLInputElement[] =>
  Array.from(
    document.querySelectorAll<HTMLInputElement>(
      '.fc-pgsetup__margins:not(.fc-pgsetup__printable) input',
    ),
  );
const printableInputs = (): HTMLInputElement[] =>
  Array.from(document.querySelectorAll<HTMLInputElement>('.fc-pgsetup__printable input'));
const buttons = (): HTMLButtonElement[] =>
  Array.from(document.querySelectorAll<HTMLButtonElement>('.fc-pgsetup .fc-fmtdlg__btn'));
const footerButtons = (): HTMLButtonElement[] =>
  Array.from(
    document.querySelectorAll<HTMLButtonElement>('.fc-pgsetup .fc-fmtdlg__footer .fc-fmtdlg__btn'),
  );
const tab = (id: string): HTMLButtonElement | null =>
  document.querySelector<HTMLButtonElement>(`.fc-pgsetup [data-pgsetup-tab="${id}"]`);
const panel = (id: string): HTMLDivElement | null =>
  document.querySelector<HTMLDivElement>(`.fc-pgsetup [data-pgsetup-tab="${id}"][role="tabpanel"]`);
const inputByLabel = (label: string): HTMLInputElement | null =>
  document.querySelector<HTMLInputElement>(`.fc-pgsetup input[aria-label="${label}"]`);
const selectByLabel = (label: string): HTMLSelectElement | null =>
  document.querySelector<HTMLSelectElement>(`.fc-pgsetup select[aria-label="${label}"]`);
const referenceError = (): HTMLElement | null =>
  document.querySelector<HTMLElement>('.fc-pgsetup__error');
const printableWarning = (): HTMLElement | null =>
  document.querySelector<HTMLElement>('.fc-pgsetup__warning');

describe('attachPageSetupDialog', () => {
  let host: HTMLElement;
  let store: SpreadsheetStore;

  beforeEach(() => {
    host = document.createElement('div');
    document.body.appendChild(host);
    store = createSpreadsheetStore();
  });

  afterEach(() => {
    while (document.body.firstChild) document.body.removeChild(document.body.firstChild);
  });

  it('mounts a hidden overlay and shows the dialog on open', () => {
    const handle = attachPageSetupDialog({ host, store });
    expect(dialog()?.hidden).toBe(true);
    handle.open();
    expect(dialog()?.hidden).toBe(false);
    handle.detach();
  });

  it('hydrates inputs from defaults and reflects the orientation/paper select values', () => {
    const handle = attachPageSetupDialog({ host, store });
    handle.open();
    const [orientSelect, paperSelect] = selects();
    expect(orientSelect?.value).toBe(defaultPageSetup().orientation);
    expect(paperSelect?.value).toBe(defaultPageSetup().paperSize);
    expect(orientSelect?.options.length).toBe(2);
    expect(paperSelect?.options.length).toBe(6);
    expect(orientSelect?.classList.contains('fc-pgsetup__select')).toBe(true);
    expect(
      Array.from(orientSelect?.options ?? [], (option) => [option.value, option.textContent]),
    ).toEqual([
      ['portrait', defaultStrings.pageSetup.orientPortrait],
      ['landscape', defaultStrings.pageSetup.orientLandscape],
    ]);
    const printQuality = selectByLabel(defaultStrings.pageSetup.printQuality);
    expect(printQuality?.classList.contains('fc-pgsetup__select')).toBe(true);
    expect(
      Array.from(printQuality?.options ?? [], (option) => [option.value, option.textContent]),
    ).toEqual([
      ['automatic', defaultStrings.pageSetup.printQualityAutomatic],
      ['300', '300 dpi'],
      ['600', '600 dpi'],
      ['1200', '1200 dpi'],
    ]);
    handle.detach();
  });

  it('labels Page Setup controls for dialog-style keyboard and assistive navigation', () => {
    const handle = attachPageSetupDialog({ host, store });
    handle.open();
    const tablist = document.querySelector<HTMLElement>('.fc-pgsetup [role="tablist"]');
    expect(tablist?.getAttribute('aria-label')).toBeTruthy();
    const [orientSelect, paperSelect] = selects();
    expect(orientSelect?.getAttribute('aria-label')).toBeTruthy();
    expect(paperSelect?.getAttribute('aria-label')).toBeTruthy();
    for (const input of marginInputs()) {
      expect(input.getAttribute('aria-label')).toBeTruthy();
    }
    const centerChecks = Array.from(
      document.querySelectorAll<HTMLInputElement>('.fc-pgsetup__center input[type="checkbox"]'),
    );
    expect(centerChecks).toHaveLength(2);
    for (const input of centerChecks) expect(input.getAttribute('aria-label')).toBeTruthy();
    expect(
      document.querySelector<HTMLInputElement>('.fc-pgsetup__text')?.getAttribute('aria-label'),
    ).toBeTruthy();
    handle.detach();
  });

  it('renders Excel-style Page Setup tabs with arrow, Home, and End navigation', () => {
    const handle = attachPageSetupDialog({ host, store });
    handle.open();

    expect(tab('page')?.getAttribute('aria-selected')).toBe('true');
    expect(panel('page')?.hidden).toBe(false);
    expect(panel('margins')?.hidden).toBe(true);

    tab('page')?.focus();
    tab('page')?.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowRight', bubbles: true }));
    expect(tab('margins')?.getAttribute('aria-selected')).toBe('true');
    expect(document.activeElement).toBe(tab('margins'));
    expect(panel('margins')?.hidden).toBe(false);

    tab('margins')?.dispatchEvent(new KeyboardEvent('keydown', { key: 'End', bubbles: true }));
    expect(tab('sheet')?.getAttribute('aria-selected')).toBe('true');
    expect(document.activeElement).toBe(tab('sheet'));

    tab('sheet')?.dispatchEvent(new KeyboardEvent('keydown', { key: 'Home', bubbles: true }));
    expect(tab('page')?.getAttribute('aria-selected')).toBe('true');
    expect(document.activeElement).toBe(tab('page'));

    handle.detach();
  });

  it('OK with a flipped orientation/paper persists via mutators.setPageSetup', () => {
    const handle = attachPageSetupDialog({ host, store });
    handle.open();
    const [orientSelect, paperSelect] = selects();
    if (!orientSelect || !paperSelect) throw new Error('selects missing');
    orientSelect.value = 'landscape';
    paperSelect.value = 'letter';
    const ok = buttons().find((b) => b.classList.contains('fc-fmtdlg__btn--primary'));
    ok?.click();
    const setup = getPageSetup(store.getState(), 0);
    expect(setup.orientation).toBe('landscape');
    expect(setup.paperSize).toBe('letter');
    expect(dialog()?.hidden).toBe(true);
    handle.detach();
  });

  it('refreshes printable bounds from a host printer profile when paper changes', () => {
    const resolvePrintableBounds = vi.fn(() => ({
      top: 0.45,
      right: 0.35,
      bottom: 0.45,
      left: 0.35,
    }));
    const handle = attachPageSetupDialog({
      host,
      store,
      resolvePrintableBounds,
    });
    handle.open();
    const [, paperSelect] = selects();
    if (!paperSelect) throw new Error('paper select missing');
    paperSelect.value = 'letter';
    const ok = buttons().find((b) => b.classList.contains('fc-fmtdlg__btn--primary'));
    ok?.click();

    const setup = getPageSetup(store.getState(), 0);
    expect(resolvePrintableBounds).toHaveBeenCalledTimes(1);
    expect(setup.paperSize).toBe('letter');
    expect(setup.printableBounds).toEqual({
      top: 0.45,
      right: 0.35,
      bottom: 0.45,
      left: 0.35,
    });
    handle.detach();
  });

  it('lets Page Setup choose a host-provided printer profile', () => {
    const setPrinterProfileId = vi.fn();
    const resolvePrintableBounds = vi.fn(() => ({
      top: 0.2,
      right: 0.15,
      bottom: 0.2,
      left: 0.15,
    }));
    const handle = attachPageSetupDialog({
      host,
      store,
      strings: en,
      getPrinterProfiles: () => [
        { id: 'office', name: 'Office Printer', printableBounds: { top: 0.1 } },
        { id: 'label', name: 'Label Printer', printableBounds: { top: 0.2 } },
      ],
      getPrinterProfileId: () => ' office ',
      setPrinterProfileId,
      resolvePrintableBounds,
    });
    handle.open();
    const printer = selectByLabel('Printer');
    if (!printer) throw new Error('printer profile select missing');
    expect(printer.hidden).toBe(false);
    expect(Array.from(printer.options).map((option) => option.textContent)).toEqual([
      'Automatically match paper and orientation',
      'Office Printer',
      'Label Printer',
    ]);
    expect(printer.value).toBe('office');

    printer.value = 'label';
    const ok = buttons().find((b) => b.classList.contains('fc-fmtdlg__btn--primary'));
    ok?.click();

    expect(setPrinterProfileId).toHaveBeenCalledWith('label');
    expect(resolvePrintableBounds).toHaveBeenCalledWith(
      expect.objectContaining({ paperSize: 'A4', orientation: 'portrait' }),
      0,
      expect.objectContaining({ paperSize: 'A4', orientation: 'portrait' }),
      'label',
    );
    expect(getPageSetup(store.getState(), 0).printableBounds).toEqual({
      top: 0.2,
      right: 0.15,
      bottom: 0.2,
      left: 0.15,
    });
    handle.detach();
  });

  it('normalizes the selected printer profile id before comparing dialog changes', () => {
    const setPrinterProfileId = vi.fn();
    const resolvePrintableBounds = vi.fn();
    const handle = attachPageSetupDialog({
      host,
      store,
      strings: en,
      getPrinterProfiles: () => [
        { id: 'office', name: 'Office Printer', printableBounds: { top: 0.1 } },
      ],
      getPrinterProfileId: () => ' office ',
      setPrinterProfileId,
      resolvePrintableBounds,
    });
    handle.open();
    const printer = selectByLabel('Printer');
    if (!printer) throw new Error('printer profile select missing');
    expect(printer.value).toBe('office');

    const ok = buttons().find((b) => b.classList.contains('fc-fmtdlg__btn--primary'));
    ok?.click();

    expect(setPrinterProfileId).not.toHaveBeenCalled();
    expect(resolvePrintableBounds).not.toHaveBeenCalled();
    handle.detach();
  });

  it('normalizes raw dialog printer profiles before rendering the selector', () => {
    const handle = attachPageSetupDialog({
      host,
      store,
      strings: en,
      getPrinterProfiles: () => [
        {
          id: ' office ',
          name: ' Office Printer ',
          printableBounds: { top: -1, right: Number.NaN, bottom: 0.1, left: 0.1 },
        },
        { id: ' office ', name: 'Duplicate', printableBounds: { top: 0.2 } },
      ],
      getPrinterProfileId: () => ' office ',
    });
    handle.open();

    const printer = selectByLabel('Printer');
    if (!printer) throw new Error('printer profile select missing');
    expect(Array.from(printer.options).map((option) => [option.value, option.textContent])).toEqual(
      [
        ['', 'Automatically match paper and orientation'],
        ['office', 'Office Printer'],
      ],
    );
    expect(printer.value).toBe('office');
    handle.detach();
  });

  it('refreshes host-provided printer profiles from Page Setup', async () => {
    const setPrinterProfileId = vi.fn();
    const refreshPrinterProfiles = vi
      .fn()
      .mockResolvedValue([
        { id: 'native', name: 'Native Printer', printableBounds: { top: 0.15 } },
      ]);
    const handle = attachPageSetupDialog({
      host,
      store,
      strings: en,
      getPrinterProfiles: () => [],
      getPrinterProfileId: () => undefined,
      setPrinterProfileId,
      refreshPrinterProfiles,
    });

    handle.open();
    const refresh = Array.from(document.querySelectorAll<HTMLButtonElement>('button')).find(
      (button) => button.textContent === 'Refresh printers',
    );
    if (!refresh) throw new Error('missing printer refresh button');
    refresh.click();
    await Promise.resolve();

    const printer = selectByLabel('Printer');
    expect(refreshPrinterProfiles).toHaveBeenCalledTimes(1);
    expect(printer?.value).toBe('');
    expect(Array.from(printer?.options ?? []).map((option) => option.value)).toEqual([
      '',
      'native',
    ]);

    if (!printer) throw new Error('missing printer select');
    printer.value = 'native';
    printer.dispatchEvent(new Event('change', { bubbles: true }));
    footerButtons()
      .find((button) => button.textContent === 'OK')
      ?.click();
    expect(setPrinterProfileId).toHaveBeenCalledWith('native');
    handle.detach();
  });

  it('projects a disabled reason while printer profiles are refreshing', async () => {
    let resolveRefresh: (
      profiles: readonly { id: string; name: string; printableBounds: { top: number } }[],
    ) => void = () => {};
    const refreshPrinterProfiles = vi.fn(
      () =>
        new Promise<readonly { id: string; name: string; printableBounds: { top: number } }[]>(
          (resolve) => {
            resolveRefresh = resolve;
          },
        ),
    );
    const handle = attachPageSetupDialog({
      host,
      store,
      strings: en,
      getPrinterProfiles: () => [],
      refreshPrinterProfiles,
    });

    handle.open();
    const refresh = Array.from(document.querySelectorAll<HTMLButtonElement>('button')).find(
      (button) => button.textContent === 'Refresh printers',
    );
    if (!refresh) throw new Error('missing printer refresh button');
    refresh.click();
    await Promise.resolve();

    expect(refresh.disabled).toBe(true);
    expect(refresh.dataset.disabledReason).toBe(en.pageSetup.printerProfileRefreshInProgress);
    expect(refresh.getAttribute('aria-description')).toBe(
      en.pageSetup.printerProfileRefreshInProgress,
    );

    resolveRefresh([{ id: 'native', name: 'Native Printer', printableBounds: { top: 0.15 } }]);
    await Promise.resolve();
    await Promise.resolve();
    expect(refresh.disabled).toBe(false);
    expect(refresh.dataset.disabledReason).toBeUndefined();
    handle.detach();
  });

  it('clears stale printable bounds when host reports no matching printer profile', () => {
    store.setState((s) => ({
      ...s,
      pageSetup: {
        setupBySheet: new Map([
          [
            0,
            {
              ...defaultPageSetup(),
              printableBounds: { top: 0.4, right: 0.4, bottom: 0.4, left: 0.4 },
            },
          ],
        ]),
      },
    }));
    const resolvePrintableBounds = vi.fn(() => null);
    const handle = attachPageSetupDialog({
      host,
      store,
      resolvePrintableBounds,
    });
    handle.open();
    const [, paperSelect] = selects();
    if (!paperSelect) throw new Error('paper select missing');
    paperSelect.value = 'letter';
    const ok = buttons().find((b) => b.classList.contains('fc-fmtdlg__btn--primary'));
    ok?.click();

    expect(resolvePrintableBounds).toHaveBeenCalledTimes(1);
    expect(getPageSetup(store.getState(), 0).printableBounds).toBeUndefined();
    handle.detach();
  });

  it('OK after editing Page tab scaling, quality, and first page number persists them', () => {
    const handle = attachPageSetupDialog({ host, store, strings: en });
    handle.open();
    const fit = inputByLabel('Fit to');
    const fitWidth = inputByLabel('Fit to width (pages)');
    const fitHeight = inputByLabel('Fit to height (pages)');
    const quality = selectByLabel('Print quality');
    const firstPage = inputByLabel('First page number');
    if (!fit || !fitWidth || !fitHeight || !quality || !firstPage) {
      throw new Error('page scaling controls missing');
    }

    fit.checked = true;
    fitWidth.value = '1';
    fitHeight.value = '3';
    quality.value = '600';
    firstPage.value = '7';

    const ok = buttons().find((b) => b.classList.contains('fc-fmtdlg__btn--primary'));
    ok?.click();
    const setup = getPageSetup(store.getState(), 0);
    expect(setup.fitWidth).toBe(1);
    expect(setup.fitHeight).toBe(3);
    expect(setup.printQuality).toBe('600');
    expect(setup.firstPageNumber).toBe(7);
    handle.detach();
  });

  it('Cancel does not mutate the page setup', () => {
    // Seed with a known value so the dialog can be canceled and we can confirm
    // the slice didn't move.
    setMarginPreset(store, 0, 'narrow');
    const before = getPageSetup(store.getState(), 0);
    const handle = attachPageSetupDialog({ host, store });
    handle.open();
    // Mutate inputs in the dialog.
    const [orientSelect] = selects();
    if (orientSelect) orientSelect.value = 'landscape';
    const inputs = pageMarginInputs();
    if (inputs[0]) inputs[0].value = '9';
    const cancel = footerButtons().find((b) => !b.classList.contains('fc-fmtdlg__btn--primary'));
    cancel?.click();
    const after = getPageSetup(store.getState(), 0);
    expect(after).toEqual(before);
    expect(dialog()?.hidden).toBe(true);
    handle.detach();
  });

  it('header close button does not mutate the page setup', () => {
    setMarginPreset(store, 0, 'narrow');
    const before = getPageSetup(store.getState(), 0);
    const handle = attachPageSetupDialog({ host, store });
    handle.open();
    const [orientSelect] = selects();
    if (orientSelect) orientSelect.value = 'landscape';
    const close = document.querySelector<HTMLButtonElement>('.fc-pgsetup .fc-fmtdlg__close');
    expect(close?.textContent).toBe('');
    close?.click();

    expect(getPageSetup(store.getState(), 0)).toEqual(before);
    expect(dialog()?.hidden).toBe(true);
    handle.detach();
  });

  it('Escape closes the overlay without mutating', () => {
    const before = getPageSetup(store.getState(), 0);
    const handle = attachPageSetupDialog({ host, store });
    handle.open();
    const overlay = dialog();
    overlay?.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    expect(overlay?.hidden).toBe(true);
    expect(getPageSetup(store.getState(), 0)).toEqual(before);
    handle.detach();
  });

  it('OK after editing margin inputs writes the updated margins to the slice', () => {
    const handle = attachPageSetupDialog({ host, store });
    handle.open();
    const inputs = pageMarginInputs();
    if (inputs.length !== 6) throw new Error('expected 6 margin inputs');
    const [top, right, bottom, left, header, footer] = inputs as [
      HTMLInputElement,
      HTMLInputElement,
      HTMLInputElement,
      HTMLInputElement,
      HTMLInputElement,
      HTMLInputElement,
    ];
    top.value = '0.5';
    right.value = '0.4';
    bottom.value = '0.3';
    left.value = '0.2';
    header.value = '0.15';
    footer.value = '0.25';
    const centerChecks = Array.from(
      document.querySelectorAll<HTMLInputElement>('.fc-pgsetup__center input[type="checkbox"]'),
    );
    if (centerChecks.length !== 2) throw new Error('expected 2 center checkboxes');
    const [centerH, centerV] = centerChecks as [HTMLInputElement, HTMLInputElement];
    centerH.checked = true;
    centerV.checked = true;
    const ok = buttons().find((b) => b.classList.contains('fc-fmtdlg__btn--primary'));
    ok?.click();
    const setup = getPageSetup(store.getState(), 0);
    expect(setup.margins).toEqual({ top: 0.5, right: 0.4, bottom: 0.3, left: 0.2 });
    expect(setup.headerMargin).toBe(0.15);
    expect(setup.footerMargin).toBe(0.25);
    expect(setup.centerHorizontally).toBe(true);
    expect(setup.centerVertically).toBe(true);
    handle.detach();
  });

  it('OK after editing printer minimum margins persists printable bounds separately', () => {
    const handle = attachPageSetupDialog({ host, store, strings: en });
    handle.open();
    const inputs = printableInputs();
    if (inputs.length !== 4) throw new Error('expected 4 printable bound inputs');
    const [top, right, bottom, left] = inputs as [
      HTMLInputElement,
      HTMLInputElement,
      HTMLInputElement,
      HTMLInputElement,
    ];
    top.value = '0.25';
    right.value = '0.2';
    bottom.value = '0.3';
    left.value = '0.2';

    const ok = buttons().find((b) => b.classList.contains('fc-fmtdlg__btn--primary'));
    ok?.click();

    const setup = getPageSetup(store.getState(), 0);
    expect(setup.printableBounds).toEqual({ top: 0.25, right: 0.2, bottom: 0.3, left: 0.2 });
    expect(setup.margins).toEqual(defaultPageSetup().margins);
    handle.detach();
  });

  it('warns when requested margins are below printer minimum margins', () => {
    const handle = attachPageSetupDialog({ host, store, strings: en });
    handle.open();
    tab('margins')?.click();
    const pageInputs = pageMarginInputs();
    const printable = printableInputs();
    if (pageInputs.length !== 6 || printable.length !== 4) {
      throw new Error('expected margin inputs');
    }
    const [topMargin] = pageInputs as [HTMLInputElement, ...HTMLInputElement[]];
    const [printableTop] = printable as [HTMLInputElement, ...HTMLInputElement[]];
    topMargin.value = '0.1';
    printableTop.value = '0.25';
    printableTop.dispatchEvent(new Event('input', { bubbles: true }));

    expect(printableWarning()?.hidden).toBe(false);
    expect(printableWarning()?.textContent).toContain('printer minimum');
    expect(printableWarning()?.textContent).toContain('Top 0.25in');

    topMargin.value = '0.3';
    topMargin.dispatchEvent(new Event('input', { bubbles: true }));
    expect(printableWarning()?.hidden).toBe(true);
    handle.detach();
  });

  it('Header/Footer tab exposes built-in presets and custom edit entry points', () => {
    const handle = attachPageSetupDialog({ host, store, strings: en });
    handle.open();
    tab('headerFooter')?.click();

    expect(selectByLabel('Header')).toBeTruthy();
    expect(selectByLabel('Footer')).toBeTruthy();
    expect(
      buttons().some((button) => button.getAttribute('aria-label') === 'Custom Header...'),
    ).toBe(true);
    expect(
      buttons().some((button) => button.getAttribute('aria-label') === 'Custom Footer...'),
    ).toBe(true);

    handle.detach();
  });

  it('OK after choosing built-in Header/Footer presets persists the generated slots', () => {
    const handle = attachPageSetupDialog({ host, store, strings: en });
    handle.open();
    tab('headerFooter')?.click();

    const headerSelect = selectByLabel('Header');
    const footerSelect = selectByLabel('Footer');
    if (!headerSelect || !footerSelect) throw new Error('header/footer selects missing');
    headerSelect.value = 'page';
    headerSelect.dispatchEvent(new Event('change', { bubbles: true }));
    footerSelect.value = 'path';
    footerSelect.dispatchEvent(new Event('change', { bubbles: true }));

    expect(inputByLabel('Header Center')?.value).toBe('Page 1');
    expect(inputByLabel('Footer Center')?.value).toBe('Book1.xlsx');

    const ok = buttons().find((b) => b.classList.contains('fc-fmtdlg__btn--primary'));
    ok?.click();
    const setup = getPageSetup(store.getState(), 0);
    expect(setup.headerLeft).toBe('');
    expect(setup.headerCenter).toBe('Page 1');
    expect(setup.headerRight).toBe('');
    expect(setup.footerLeft).toBe('');
    expect(setup.footerCenter).toBe('Book1.xlsx');
    expect(setup.footerRight).toBe('');
    handle.detach();
  });

  it('Header/Footer tab persists Excel-style header and footer option checkboxes', () => {
    const handle = attachPageSetupDialog({ host, store, strings: en });
    handle.open();
    tab('headerFooter')?.click();

    const oddEven = inputByLabel('Different odd and even pages');
    const first = inputByLabel('Different first page');
    const scale = inputByLabel('Scale with document');
    const align = inputByLabel('Align with page margins');
    if (!oddEven || !first || !scale || !align) {
      throw new Error('header/footer option checkboxes missing');
    }
    expect(scale.checked).toBe(true);
    expect(align.checked).toBe(true);

    oddEven.checked = true;
    first.checked = true;
    scale.checked = false;
    align.checked = false;

    const ok = buttons().find((b) => b.classList.contains('fc-fmtdlg__btn--primary'));
    ok?.click();
    const setup = getPageSetup(store.getState(), 0);
    expect(setup.differentOddEvenPages).toBe(true);
    expect(setup.differentFirstPage).toBe(true);
    expect(setup.scaleHeaderFooterWithDocument).toBe(false);
    expect(setup.alignHeaderFooterWithMargins).toBe(false);
    handle.detach();
  });

  it('Custom Header button focuses the direct left-header input without applying a preset', () => {
    const handle = attachPageSetupDialog({ host, store, strings: en });
    handle.open();
    tab('headerFooter')?.click();

    const customHeader = buttons().find(
      (button) => button.getAttribute('aria-label') === 'Custom Header...',
    );
    customHeader?.click();

    expect(document.activeElement).toBe(inputByLabel('Header Left'));
    expect(selectByLabel('Header')?.value).toBe('custom');
    handle.detach();
  });

  it('Sheet tab persists Excel-style print area, print options, comments, errors, and page order', () => {
    const handle = attachPageSetupDialog({ host, store, strings: en });
    handle.open();
    tab('sheet')?.click();

    const printArea = inputByLabel('Print area');
    const rows = inputByLabel('Print title rows');
    const cols = inputByLabel('Print title columns');
    const comments = selectByLabel('Comments and notes');
    const errors = selectByLabel('Cell errors as');
    const blackWhite = inputByLabel('Black and white');
    const draft = inputByLabel('Draft quality');
    const overThenDown = inputByLabel('Over, then down');
    if (
      !printArea ||
      !rows ||
      !cols ||
      !comments ||
      !errors ||
      !blackWhite ||
      !draft ||
      !overThenDown
    ) {
      throw new Error('sheet tab controls missing');
    }

    printArea.value = 'B2:D10';
    rows.value = '1:2';
    cols.value = 'A:B';
    blackWhite.checked = true;
    draft.checked = true;
    comments.value = 'endOfSheet';
    errors.value = 'dash';
    overThenDown.checked = true;

    const ok = buttons().find((b) => b.classList.contains('fc-fmtdlg__btn--primary'));
    ok?.click();
    const setup = getPageSetup(store.getState(), 0);
    expect(setup.printArea).toBe('B2:D10');
    expect(setup.printTitleRows).toBe('1:2');
    expect(setup.printTitleCols).toBe('A:B');
    expect(setup.blackAndWhite).toBe(true);
    expect(setup.draftQuality).toBe(true);
    expect(setup.comments).toBe('endOfSheet');
    expect(setup.cellErrorsAs).toBe('dash');
    expect(setup.pageOrder).toBe('overThenDown');
    handle.detach();
  });

  it('Sheet tab range pickers use the live selection for print area and titles', () => {
    const handle = attachPageSetupDialog({ host, store, strings: en });
    handle.open('sheet');

    const printArea = inputByLabel('Print area');
    const rows = inputByLabel('Print title rows');
    const cols = inputByLabel('Print title columns');
    if (!printArea || !rows || !cols) throw new Error('sheet reference controls missing');

    const printAreaPicker = document.querySelector<HTMLButtonElement>(
      '[data-range-picker="page-setup-print-area"]',
    );
    const rowsPicker = document.querySelector<HTMLButtonElement>(
      '[data-range-picker="page-setup-print-title-rows"]',
    );
    const colsPicker = document.querySelector<HTMLButtonElement>(
      '[data-range-picker="page-setup-print-title-cols"]',
    );
    expect(printAreaPicker?.getAttribute('aria-label')).toBe('Select range');
    expect(rowsPicker?.getAttribute('aria-label')).toBe('Select range');
    expect(colsPicker?.getAttribute('aria-label')).toBe('Select range');

    printAreaPicker?.click();
    expect(printAreaPicker?.dataset.rangePickerActive).toBe('true');
    expect(dialog()?.classList.contains('fc-fmtdlg--range-picking')).toBe(true);
    mutators.setRange(store, { sheet: 0, r0: 1, c0: 1, r1: 9, c1: 3 });
    expect(printArea.value).toBe('B2:D10');

    rowsPicker?.click();
    expect(printAreaPicker?.dataset.rangePickerActive).toBe('false');
    expect(rowsPicker?.dataset.rangePickerActive).toBe('true');
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 2, r1: 1, c1: 5 });
    expect(rows.value).toBe('1:2');

    colsPicker?.click();
    expect(rowsPicker?.dataset.rangePickerActive).toBe('false');
    expect(colsPicker?.dataset.rangePickerActive).toBe('true');
    mutators.setRange(store, { sheet: 0, r0: 4, c0: 0, r1: 12, c1: 1 });
    expect(cols.value).toBe('A:B');

    handle.close();
    expect(colsPicker?.dataset.rangePickerActive).toBe('false');
    expect(dialog()?.classList.contains('fc-fmtdlg--range-picking')).toBe(false);
    handle.detach();
  });

  it('rejects invalid Sheet tab print references without mutating page setup', () => {
    const handle = attachPageSetupDialog({ host, store, strings: en });
    handle.open();
    tab('sheet')?.click();

    const printArea = inputByLabel('Print area');
    const rows = inputByLabel('Print title rows');
    const cols = inputByLabel('Print title columns');
    if (!printArea || !rows || !cols) throw new Error('sheet reference controls missing');

    printArea.value = '1:3';
    rows.value = '1:2';
    cols.value = 'A:B';
    const ok = buttons().find((b) => b.classList.contains('fc-fmtdlg__btn--primary'));
    ok?.click();

    expect(dialog()?.hidden).toBe(false);
    expect(getPageSetup(store.getState(), 0).printArea).toBeUndefined();
    expect(printArea.getAttribute('aria-invalid')).toBe('true');
    expect(referenceError()?.textContent).toBe('Enter a valid cell range, such as A1:D20.');
    expect(tab('sheet')?.getAttribute('aria-selected')).toBe('true');
    expect(document.activeElement).toBe(printArea);

    printArea.value = 'B2:D10';
    printArea.dispatchEvent(new Event('input', { bubbles: true }));
    expect(referenceError()?.hidden).toBe(true);
    expect(printArea.hasAttribute('aria-invalid')).toBe(false);

    rows.value = 'A:B';
    ok?.click();
    expect(getPageSetup(store.getState(), 0).printArea).toBeUndefined();
    expect(rows.getAttribute('aria-invalid')).toBe('true');
    expect(referenceError()?.textContent).toBe('Enter valid rows to repeat, such as 1:3.');

    rows.value = '1:2';
    rows.dispatchEvent(new Event('input', { bubbles: true }));
    cols.value = '1:3';
    ok?.click();
    expect(getPageSetup(store.getState(), 0).printArea).toBeUndefined();
    expect(cols.getAttribute('aria-invalid')).toBe('true');
    expect(referenceError()?.textContent).toBe('Enter valid columns to repeat, such as A:B.');

    handle.detach();
  });

  it('detach removes the overlay node from the DOM', () => {
    const handle = attachPageSetupDialog({ host, store });
    handle.detach();
    expect(dialog()).toBeNull();
  });

  it('keeps Page Setup controls on compact desktop dialog geometry', () => {
    const css = readFileSync(join(root, 'src/styles/core/app/panels/page-setup.css'), 'utf8');

    expect(css).toMatch(/\.fc-pgsetup__mini-btn\s*\{[\s\S]*?border-radius: 2px;/);
    expect(css).toMatch(
      /\.fc-pgsetup__text\[aria-invalid="true"\]\s*\{[\s\S]*?outline: 1px solid[\s\S]*?outline-offset: -1px;/,
    );
    expect(css).not.toContain('outline: 2px solid color-mix(in srgb, #c50f1f');
  });
});
