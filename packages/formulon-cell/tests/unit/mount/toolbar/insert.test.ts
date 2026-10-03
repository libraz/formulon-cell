import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { Spreadsheet } from '../../../../src/mount.js';
import { mutators } from '../../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../../test-utils/mount.js';
import { seedNumber, seedText, stubHelpers, waitFor } from './fixtures.js';

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

  it('opens Create Table from the Insert ribbon primary button', async () => {
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 2 });
    seedText(sheet, 0, 0, 'Region');
    seedText(sheet, 0, 1, 'Sales');
    seedText(sheet, 0, 2, 'Qty');
    seedText(sheet, 1, 0, 'East');
    seedNumber(sheet, 1, 1, 10);
    seedNumber(sheet, 1, 2, 2);
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('insert');

    const tableButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="formatTableInsert"]',
    );
    expect(tableButton).toBeTruthy();
    tableButton?.click();
    await Promise.resolve();

    expect(document.body.textContent).toContain('Create Table');
    const dialog = document.body.querySelector<HTMLElement>('.fc-tb__dlg');
    const rangeInput = dialog?.querySelector<HTMLInputElement>('input[type="text"]');
    expect(rangeInput?.value).toBe('Sheet1!$A$1:$C$4');
    await waitFor(() => document.activeElement === rangeInput);
    expect(rangeInput?.closest('.fc-range-picker')).toBeTruthy();
    const rangePicker = dialog?.querySelector<HTMLButtonElement>(
      '[data-range-picker="table-range"]',
    );
    expect(rangePicker).toBeTruthy();
    expect(rangePicker?.getAttribute('aria-label')).toBe('Select range');
    expect(dialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.textContent).toBe(
      'OK',
    );
    expect(
      Array.from(dialog?.querySelectorAll<HTMLButtonElement>('.fc-fmtdlg__btn') ?? []).some(
        (button) => button.textContent === 'Cancel',
      ),
    ).toBe(true);
    rangePicker?.click();
    expect(rangePicker?.dataset.rangePickerActive).toBe('true');
    expect(rangePicker?.getAttribute('aria-pressed')).toBe('true');
    expect(
      rangeInput?.closest('.fc-range-picker')?.classList.contains('fc-range-picker--picking'),
    ).toBe(true);
    expect(dialog?.closest('.fc-fmtdlg')?.classList.contains('fc-fmtdlg--range-picking')).toBe(
      true,
    );
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 1, c0: 1, r1: 4, c1: 3 });
    expect(rangeInput?.value).toBe('Sheet1!$B$2:$D$5');
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    expect(rangePicker?.dataset.rangePickerActive).toBe('false');
    expect(rangePicker?.getAttribute('aria-pressed')).toBe('false');
    expect(dialog?.closest('.fc-fmtdlg')?.classList.contains('fc-fmtdlg--range-picking')).toBe(
      false,
    );
    const headersCheckbox = dialog?.querySelector<HTMLInputElement>('input[type="checkbox"]');
    expect(headersCheckbox?.checked).toBe(true);
    if (headersCheckbox) headersCheckbox.checked = false;
    dialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await Promise.resolve();

    expect(sheet.instance.store.getState().tables.tables).toMatchObject([
      {
        range: { sheet: 0, r0: 1, c0: 1, r1: 4, c1: 3 },
        style: 'medium',
        showHeader: false,
        banded: true,
      },
    ]);

    tb.dispose();
  });

  it('opens custom script from primary click and keeps built-in script actions secondary', async () => {
    seedText(sheet, 0, 0, ' alpha ');
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('automate');

    const scriptButton = host.querySelector<HTMLButtonElement>('[data-ribbon-command="script"]');
    expect(scriptButton).toBeTruthy();
    expect(scriptButton?.dataset.ribbonActivation).toBe('splitPrimary');
    scriptButton?.click();
    expect(host.querySelector<HTMLDivElement>('#menu-script')?.hidden).toBe(true);
    const scriptDialog = document.body.querySelector<HTMLElement>('.fc-tb__dlg');
    expect(scriptDialog?.textContent).toContain('Script');
    expect(scriptDialog?.textContent).toContain('Trim whitespace');
    const select = scriptDialog?.querySelector<HTMLSelectElement>('[data-script-command-select]');
    expect(select).toBeTruthy();
    if (!select) throw new Error('Expected script command select.');
    select.value = 'trim';
    scriptDialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await Promise.resolve();
    await Promise.resolve();

    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
      kind: 'text',
      value: 'alpha',
    });
    document.body.querySelector<HTMLButtonElement>('.fc-tb__dlg .fc-fmtdlg__btn--primary')?.click();

    seedText(sheet, 1, 0, ' BETA ');
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 1, c0: 0, r1: 1, c1: 0 });
    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'script', menuId: 'menu-script' },
      scriptButton as HTMLButtonElement,
    );
    expect(host.querySelectorAll('#menu-script .fc-tb__menu-item--iconic').length).toBe(5);
    const trimButton = host.querySelector<HTMLButtonElement>('[data-script-action="trim"]');
    expect(trimButton).toBeTruthy();
    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: trimButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);
    await Promise.resolve();

    expect(sheet.workbook.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({
      kind: 'text',
      value: 'BETA',
    });
    expect(document.body.querySelector<HTMLElement>('.fc-tb__dlg')?.textContent).toContain(
      '1 cell(s) changed',
    );
    document.body.querySelector<HTMLButtonElement>('.fc-tb__dlg .fc-fmtdlg__btn--primary')?.click();

    tb.dispose();
  });

  it('runs PDF and Add-ins primary actions while keeping secondary menus available', async () => {
    const print = vi.spyOn(sheet.instance, 'print').mockImplementation(() => undefined);
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('acrobat');

    const pdfButton = host.querySelector<HTMLButtonElement>('[data-ribbon-command="pdf"]');
    expect(pdfButton).toBeTruthy();
    expect(pdfButton?.dataset.ribbonActivation).toBe('splitPrimary');
    pdfButton?.click();
    await Promise.resolve();
    expect(host.querySelector<HTMLDivElement>('#menu-pdf')?.hidden).toBe(true);
    expect(print).toHaveBeenCalledWith('pdf');
    expect(document.body.querySelector<HTMLElement>('.fc-tb__dlg')?.textContent).toContain(
      'PDF export has been sent',
    );
    document.body.querySelector<HTMLButtonElement>('.fc-tb__dlg .fc-fmtdlg__btn--primary')?.click();

    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'pdf', menuId: 'menu-pdf' },
      pdfButton as HTMLButtonElement,
    );
    expect(host.querySelectorAll('#menu-pdf .fc-tb__menu-item--iconic').length).toBe(3);
    const createPdfButton = host.querySelector<HTMLButtonElement>('[data-pdf-action="create"]');
    expect(createPdfButton).toBeTruthy();
    const createPdfEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(createPdfEvent, 'target', { value: createPdfButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(createPdfEvent)).toBe(true);
    await Promise.resolve();
    expect(print).toHaveBeenCalledWith('pdf');
    expect(document.body.querySelector<HTMLElement>('.fc-tb__dlg')?.textContent).toContain(
      'PDF export has been sent',
    );
    document.body.querySelector<HTMLButtonElement>('.fc-tb__dlg .fc-fmtdlg__btn--primary')?.click();

    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'pdf', menuId: 'menu-pdf' },
      pdfButton as HTMLButtonElement,
    );
    const shareButton = host.querySelector<HTMLButtonElement>('[data-pdf-action="share"]');
    expect(shareButton).toBeTruthy();
    const shareEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(shareEvent, 'target', { value: shareButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(shareEvent)).toBe(true);
    await Promise.resolve();
    expect(print).toHaveBeenCalledWith('pdf');
    expect(document.body.querySelector<HTMLElement>('.fc-tb__dlg')?.textContent).toContain(
      'PDF export is ready',
    );
    document.body.querySelector<HTMLButtonElement>('.fc-tb__dlg .fc-fmtdlg__btn--primary')?.click();

    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'pdf', menuId: 'menu-pdf' },
      pdfButton as HTMLButtonElement,
    );
    const preferencesButton = host.querySelector<HTMLButtonElement>(
      '[data-pdf-action="preferences"]',
    );
    expect(preferencesButton).toBeTruthy();
    const pdfEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(pdfEvent, 'target', { value: preferencesButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(pdfEvent)).toBe(true);
    expect(document.body.querySelector<HTMLElement>('.fc-pgsetup')?.hidden).toBe(false);
    document.body
      .querySelector<HTMLButtonElement>('.fc-pgsetup .fc-fmtdlg__close')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    const addInButton = host.querySelector<HTMLButtonElement>('[data-ribbon-command="addIn"]');
    expect(addInButton).toBeTruthy();
    expect(addInButton?.dataset.ribbonActivation).toBe('splitPrimary');
    addInButton?.click();
    await Promise.resolve();
    expect(host.querySelector<HTMLDivElement>('#menu-add-ins')?.hidden).toBe(true);
    expect(document.body.querySelector<HTMLElement>('.fc-tb__dlg')?.textContent).toContain(
      'Add-in management',
    );
    document.body.querySelector<HTMLButtonElement>('.fc-tb__dlg .fc-fmtdlg__btn--primary')?.click();

    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'addIn', menuId: 'menu-add-ins' },
      addInButton as HTMLButtonElement,
    );
    expect(host.querySelectorAll('#menu-add-ins .fc-tb__menu-item--iconic').length).toBe(3);
    const getButton = host.querySelector<HTMLButtonElement>('[data-add-in-action="get"]');
    expect(getButton).toBeTruthy();
    const getAddInEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(getAddInEvent, 'target', { value: getButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(getAddInEvent)).toBe(true);
    await Promise.resolve();
    expect(document.body.querySelector<HTMLElement>('.fc-tb__dlg')?.textContent).toContain(
      'Get Add-ins',
    );
    expect(document.body.querySelector<HTMLElement>('.fc-tb__dlg')?.textContent).toContain(
      'Office Add-ins',
    );
    document.body.querySelector<HTMLButtonElement>('.fc-tb__dlg .fc-fmtdlg__btn--primary')?.click();

    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'addIn', menuId: 'menu-add-ins' },
      addInButton as HTMLButtonElement,
    );
    const myButton = host.querySelector<HTMLButtonElement>('[data-add-in-action="my"]');
    expect(myButton).toBeTruthy();
    const addInEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(addInEvent, 'target', { value: myButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(addInEvent)).toBe(true);
    await Promise.resolve();
    expect(document.body.querySelector<HTMLElement>('.fc-tb__dlg')?.textContent).toContain(
      'External add-ins',
    );
    document.body.querySelector<HTMLButtonElement>('.fc-tb__dlg .fc-fmtdlg__btn--primary')?.click();

    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'addIn', menuId: 'menu-add-ins' },
      addInButton as HTMLButtonElement,
    );
    const manageButton = host.querySelector<HTMLButtonElement>('[data-add-in-action="manage"]');
    expect(manageButton).toBeTruthy();
    const manageEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(manageEvent, 'target', { value: manageButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(manageEvent)).toBe(true);
    await Promise.resolve();
    expect(document.body.querySelector<HTMLElement>('.fc-tb__dlg')?.textContent).toContain(
      'Add-in management',
    );
    document.body.querySelector<HTMLButtonElement>('.fc-tb__dlg .fc-fmtdlg__btn--primary')?.click();

    tb.dispose();
    print.mockRestore();
  });

  it('surfaces Insert media reports and creates session shapes', async () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('insert');

    const pictureButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="pictureInsert"]',
    );
    expect(pictureButton).toBeTruthy();
    pictureButton?.click();
    const pictureMenu = host.querySelector<HTMLElement>('#menu-picture-insert');
    expect(pictureMenu?.classList.contains('fc-tb__menu--visual')).toBe(true);
    expect(pictureMenu?.querySelectorAll('.fc-tb__visual-tile')).toHaveLength(3);
    const stockButton = host.querySelector<HTMLButtonElement>('[data-picture-insert="stock"]');
    expect(stockButton).toBeTruthy();
    const stockEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(stockEvent, 'target', { value: stockButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(stockEvent)).toBe(true);
    await Promise.resolve();
    expect(document.body.querySelector<HTMLElement>('.fc-tb__dlg')?.textContent).toContain(
      'Stock Images',
    );
    expect(document.body.querySelector<HTMLElement>('.fc-tb__dlg')?.textContent).toContain(
      'host-provided media picker',
    );
    document.body.querySelector<HTMLButtonElement>('.fc-tb__dlg .fc-fmtdlg__btn--primary')?.click();

    pictureButton?.click();
    const onlineButton = host.querySelector<HTMLButtonElement>('[data-picture-insert="online"]');
    expect(onlineButton).toBeTruthy();
    const pictureEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(pictureEvent, 'target', { value: onlineButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(pictureEvent)).toBe(true);
    await Promise.resolve();
    const pictureDialog = document.body.querySelector<HTMLElement>('.fc-tb__dlg');
    expect(pictureDialog?.textContent).toContain('Online Pictures');
    expect(pictureDialog?.textContent).toContain('host-provided media picker');
    expect(pictureDialog?.querySelector('input')).toBeNull();
    pictureDialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    expect(sheet.instance.store.getState().illustrations.illustrations).toEqual([]);

    const toDataUrl = vi
      .spyOn(HTMLCanvasElement.prototype, 'toDataURL')
      .mockReturnValue('data:image/png;base64,current-view');
    const screenshotButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="screenshotInsert"]',
    );
    expect(screenshotButton).toBeTruthy();
    screenshotButton?.click();
    const screenshotMenu = host.querySelector<HTMLElement>('#menu-screenshot-insert');
    expect(screenshotMenu?.classList.contains('fc-tb__menu--visual')).toBe(true);
    expect(screenshotMenu?.querySelectorAll('.fc-tb__visual-tile')).toHaveLength(2);
    expect(screenshotMenu?.querySelector('.fc-tb__menu-heading')?.textContent).toBe(
      'Available Windows',
    );
    const currentViewButton = host.querySelector<HTMLButtonElement>(
      '[data-screenshot-insert="current-view"]',
    );
    expect(currentViewButton).toBeTruthy();
    expect(currentViewButton?.classList.contains('fc-tb__visual-tile--screenshot-preview')).toBe(
      true,
    );
    const screenshotEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(screenshotEvent, 'target', { value: currentViewButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(screenshotEvent)).toBe(true);
    expect(sheet.instance.store.getState().illustrations.illustrations).toMatchObject([
      {
        kind: 'image',
        src: 'data:image/png;base64,current-view',
        sheet: 0,
      },
    ]);
    toDataUrl.mockRestore();

    const OriginalFileReader = globalThis.FileReader;
    vi.stubGlobal(
      'FileReader',
      class {
        result: string | ArrayBuffer | null = null;
        private listeners = new Map<string, EventListener[]>();
        addEventListener(type: string, listener: EventListener): void {
          this.listeners.set(type, [...(this.listeners.get(type) ?? []), listener]);
        }
        readAsDataURL(file: File): void {
          this.result = `data:${file.type};base64,from-device`;
          for (const listener of this.listeners.get('load') ?? []) {
            listener.call(this, new Event('load'));
          }
        }
      } as unknown as typeof FileReader,
    );
    const inputClick = vi.spyOn(HTMLInputElement.prototype, 'click').mockImplementation(() => {});
    const createElement = vi.spyOn(document, 'createElement');
    const originalCreateElement = createElement.getMockImplementation();
    let fileInput: HTMLInputElement | null = null;
    createElement.mockImplementation(((tagName: string, options?: ElementCreationOptions) => {
      const el = originalCreateElement
        ? originalCreateElement.call(document, tagName, options)
        : Document.prototype.createElement.call(document, tagName, options);
      if (tagName.toLowerCase() === 'input') fileInput = el as HTMLInputElement;
      return el;
    }) as typeof document.createElement);
    pictureButton?.click();
    const deviceButton = host.querySelector<HTMLButtonElement>('[data-picture-insert="device"]');
    expect(deviceButton).toBeTruthy();
    const deviceEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(deviceEvent, 'target', { value: deviceButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(deviceEvent)).toBe(true);
    await Promise.resolve();
    expect(inputClick).toHaveBeenCalled();
    expect(fileInput).toBeTruthy();
    const selectedFileInput = fileInput as HTMLInputElement | null;
    if (!selectedFileInput) throw new Error('Expected device image file input.');
    Object.defineProperty(selectedFileInput, 'files', {
      value: [new File(['device'], 'device.png', { type: 'image/png' })],
      configurable: true,
    });
    selectedFileInput.dispatchEvent(new Event('change'));
    await Promise.resolve();
    expect(sheet.instance.store.getState().illustrations.illustrations).toMatchObject([
      {
        kind: 'image',
        src: 'data:image/png;base64,current-view',
      },
      {
        kind: 'image',
        src: 'data:image/png;base64,from-device',
        alt: 'device.png',
        sheet: 0,
      },
    ]);
    createElement.mockRestore();
    inputClick.mockRestore();
    vi.stubGlobal('FileReader', OriginalFileReader);

    const shapeButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="shapesInsert"]',
    );
    expect(shapeButton).toBeTruthy();
    shapeButton?.click();
    const shapesMenu = host.querySelector<HTMLElement>('#menu-shapes-insert');
    expect(shapesMenu?.classList.contains('fc-tb__menu--visual')).toBe(true);
    expect(shapesMenu?.querySelectorAll('.fc-tb__visual-tile')).toHaveLength(7);
    expect(
      Array.from(shapesMenu?.querySelectorAll<HTMLElement>('.fc-tb__menu-heading') ?? []).map(
        (heading) => heading.textContent,
      ),
    ).toEqual(['Lines', 'Rectangles', 'Basic Shapes']);
    const arrowButton = host.querySelector<HTMLButtonElement>('[data-shape-insert="arrow"]');
    const diamondButton = host.querySelector<HTMLButtonElement>('[data-shape-insert="diamond"]');
    expect(arrowButton).toBeTruthy();
    expect(diamondButton).toBeTruthy();
    const shapeEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(shapeEvent, 'target', { value: diamondButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(shapeEvent)).toBe(true);
    expect(sheet.instance.store.getState().illustrations.illustrations).toMatchObject([
      {
        kind: 'image',
        src: 'data:image/png;base64,current-view',
      },
      {
        kind: 'image',
        src: 'data:image/png;base64,from-device',
      },
      {
        kind: 'shape',
        shape: 'diamond',
        sheet: 0,
        w: 160,
        h: 96,
      },
    ]);

    tb.dispose();
  });

  it('inserts host-provided screen clippings from the Screenshot menu', async () => {
    const screenSheet = await mountStubSheet({
      locale: 'en',
      captureScreenClip: () => ({
        src: 'data:image/png;base64,screen-clip',
        alt: 'Screen clipping',
      }),
    });
    const screenHost = document.createElement('div');
    document.body.appendChild(screenHost);
    try {
      const tb = Spreadsheet.mountToolbar(screenHost, screenSheet.instance, {
        dynamicDropdowns: true,
        helpers: stubHelpers(),
      });
      tb.setActiveTab('insert');
      const screenshotButton = screenHost.querySelector<HTMLButtonElement>(
        '[data-ribbon-command="screenshotInsert"]',
      );
      expect(screenshotButton).toBeTruthy();
      screenshotButton?.click();
      const screenClippingButton = screenHost.querySelector<HTMLButtonElement>(
        '[data-screenshot-insert="screen-clipping"]',
      );
      expect(screenClippingButton).toBeTruthy();
      const event = new MouseEvent('click', { bubbles: true });
      Object.defineProperty(event, 'target', { value: screenClippingButton });
      expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);
      await Promise.resolve();
      await Promise.resolve();
      expect(screenSheet.instance.store.getState().illustrations.illustrations).toMatchObject([
        {
          alt: 'Screen clipping',
          kind: 'image',
          sheet: 0,
          src: 'data:image/png;base64,screen-clip',
          w: 240,
          h: 160,
        },
      ]);
      tb.dispose();
    } finally {
      screenSheet.dispose();
      screenHost.remove();
    }
  });

  it('reports Screen Clipping as host-provided when no capture hook is available', async () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('insert');
    const screenshotButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="screenshotInsert"]',
    );
    expect(screenshotButton).toBeTruthy();
    screenshotButton?.click();
    const screenClippingButton = host.querySelector<HTMLButtonElement>(
      '[data-screenshot-insert="screen-clipping"]',
    );
    expect(screenClippingButton).toBeTruthy();
    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: screenClippingButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);
    await Promise.resolve();
    await Promise.resolve();

    const dialog = document.body.querySelector<HTMLElement>('.fc-tb__dlg');
    expect(dialog?.textContent).toContain('Screen Clipping');
    expect(dialog?.textContent).toContain('captureScreenClip');
    dialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await Promise.resolve();

    tb.dispose();
  });

  it('reports Current Sheet View screenshot export failures with a screenshot-specific detail', async () => {
    const toDataUrl = vi.spyOn(HTMLCanvasElement.prototype, 'toDataURL').mockReturnValue('');
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('insert');
    const screenshotButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="screenshotInsert"]',
    );
    expect(screenshotButton).toBeTruthy();
    screenshotButton?.click();
    const currentViewButton = host.querySelector<HTMLButtonElement>(
      '[data-screenshot-insert="current-view"]',
    );
    expect(currentViewButton).toBeTruthy();
    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: currentViewButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);
    await Promise.resolve();
    await Promise.resolve();

    const dialog = document.body.querySelector<HTMLElement>('.fc-tb__dlg');
    expect(dialog?.textContent).toContain('Current Sheet View');
    expect(dialog?.textContent).toContain('mounted grid canvas');
    dialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await Promise.resolve();

    toDataUrl.mockRestore();
    tb.dispose();
  });

  it('opens Recommended Charts report from primary click and keeps chart types secondary', async () => {
    seedNumber(sheet, 0, 0, 1);
    seedNumber(sheet, 1, 0, 2);
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('insert');

    const chartButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="chartInsert"]',
    );
    expect(chartButton).toBeTruthy();
    chartButton?.click();
    await Promise.resolve();
    expect(document.body.querySelector<HTMLElement>('.fc-tb__dlg')?.textContent).toContain(
      'Recommended Charts',
    );
    expect(document.body.querySelector<HTMLElement>('.fc-tb__dlg')?.textContent).toContain(
      'Persisted chart creation and editing',
    );
    expect(sheet.instance.store.getState().charts.charts).toEqual([]);
    document.body.querySelector<HTMLButtonElement>('.fc-tb__dlg .fc-fmtdlg__btn--primary')?.click();

    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 });
    tb.dropdownsApi?.openDynamicRibbonDropdown({
      command: 'chartInsert',
      menuId: 'menu-chart-insert',
    });
    const chartMenu = host.querySelector<HTMLDivElement>('#menu-chart-insert');
    expect(chartMenu?.classList.contains('fc-tb__menu--visual')).toBe(true);
    expect(chartMenu?.querySelectorAll('.fc-tb__visual-tile[data-chart-insert]').length).toBe(7);
    const recommendedButton = host.querySelector<HTMLButtonElement>(
      '[data-chart-insert="recommended"]',
    );
    expect(recommendedButton).toBeTruthy();
    const recommendedEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(recommendedEvent, 'target', { value: recommendedButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(recommendedEvent)).toBe(true);
    await Promise.resolve();
    expect(document.body.querySelector<HTMLElement>('.fc-tb__dlg')?.textContent).toContain(
      'Recommended Charts',
    );
    expect(sheet.instance.store.getState().charts.charts).toEqual([]);
    document.body.querySelector<HTMLButtonElement>('.fc-tb__dlg .fc-fmtdlg__btn--primary')?.click();

    tb.dropdownsApi?.openDynamicRibbonDropdown({
      command: 'chartInsert',
      menuId: 'menu-chart-insert',
    });
    const barButton = host.querySelector<HTMLButtonElement>('[data-chart-insert="bar"]');
    expect(barButton).toBeTruthy();
    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: barButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);

    expect(sheet.instance.store.getState().charts.charts).toMatchObject([
      { kind: 'bar', source: { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 } },
    ]);

    tb.dispose();
  });

  it('opens the PivotTable dialog from the Insert ribbon primary button', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('insert');

    const pivotButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="pivotTableInsert"]',
    );
    expect(pivotButton).toBeTruthy();
    pivotButton?.click();

    expect(document.body.textContent).toContain('Create PivotTable');
    document.body.querySelector<HTMLButtonElement>('.fc-tb__dlg .fc-fmtdlg__btn')?.click();
    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'pivotTableInsert', menuId: 'menu-pivot-table' },
      pivotButton as HTMLButtonElement,
    );
    expect(host.querySelectorAll('#menu-pivot-table .fc-tb__menu-item--iconic').length).toBe(5);
    expect(
      host.querySelector<HTMLButtonElement>('[data-pivot-table-action="dialog"]'),
    ).toBeTruthy();
    expect(
      host.querySelector<HTMLButtonElement>('[data-pivot-table-action="refresh"]'),
    ).toBeTruthy();
    const recommendedButton = host.querySelector<HTMLButtonElement>(
      '[data-pivot-table-action="recommended"]',
    );
    expect(recommendedButton).toBeTruthy();
    const recommendedEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(recommendedEvent, 'target', { value: recommendedButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(recommendedEvent)).toBe(true);
    expect(document.body.querySelector<HTMLElement>('.fc-tb__dlg')?.textContent).toContain(
      'Recommended PivotTables',
    );
    expect(document.body.querySelector<HTMLElement>('.fc-tb__dlg')?.textContent).toContain(
      'Creating or editing PivotTable definitions',
    );
    document.body.querySelector<HTMLButtonElement>('.fc-tb__dlg .fc-fmtdlg__btn--primary')?.click();

    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'pivotTableInsert', menuId: 'menu-pivot-table' },
      pivotButton as HTMLButtonElement,
    );
    const newSheetButton = host.querySelector<HTMLButtonElement>(
      '[data-pivot-table-action="new-sheet"]',
    );
    expect(newSheetButton).toBeTruthy();
    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: newSheetButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);
    expect(
      document.querySelector<HTMLInputElement>('input[name="fc-pivotdlg-destination"]:checked')
        ?.value,
    ).toBe('new');

    Array.from(document.querySelectorAll<HTMLButtonElement>('.fc-pivotdlg .fc-fmtdlg__btn'))
      .find((button) => button.textContent === 'Cancel')
      ?.click();
    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'pivotTableInsert', menuId: 'menu-pivot-table' },
      pivotButton as HTMLButtonElement,
    );
    const existingSheetButton = host.querySelector<HTMLButtonElement>(
      '[data-pivot-table-action="existing-sheet"]',
    );
    expect(existingSheetButton).toBeTruthy();
    const existingSheetEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(existingSheetEvent, 'target', { value: existingSheetButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(existingSheetEvent)).toBe(true);
    expect(
      document.querySelector<HTMLInputElement>('input[name="fc-pivotdlg-destination"]:checked')
        ?.value,
    ).toBe('existing');

    tb.dispose();
  });

  it('opens More Symbols from the Insert Symbol primary button', async () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('insert');

    const symbolButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="symbolInsert"]',
    );
    expect(symbolButton).toBeTruthy();
    expect(symbolButton?.dataset.ribbonActivation).toBe('splitPrimary');
    symbolButton?.click();

    await Promise.resolve();
    const dialog = document.body.querySelector<HTMLElement>('.fc-tb__dlg');
    expect(dialog?.textContent).toContain('More Symbols');
    const input = dialog?.querySelector<HTMLInputElement>('input');
    expect(input).toBeTruthy();
    if (!input) throw new Error('Expected More Symbols input.');
    input.value = 'Ω';
    dialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await Promise.resolve();
    await Promise.resolve();

    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
      kind: 'text',
      value: 'Ω',
    });

    tb.dispose();
  });

  it('opens the Insert Symbol secondary menu and inserts the selected symbol', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('insert');

    const symbolButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="symbolInsert"]',
    );
    expect(symbolButton).toBeTruthy();
    expect(
      tb.dropdownsApi?.dynamicDropdownSpecForButton(symbolButton as HTMLButtonElement),
    ).toEqual({
      command: 'symbolInsert',
      menuId: 'menu-symbol',
    });
    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'symbolInsert', menuId: 'menu-symbol' },
      symbolButton as HTMLButtonElement,
    );
    const menu = host.querySelector<HTMLDivElement>('#menu-symbol');
    expect(menu?.hidden).toBe(false);
    expect(menu?.classList.contains('fc-tb__menu--symbols')).toBe(true);
    expect(menu?.querySelectorAll('.fc-tb__symbol-grid').length).toBeGreaterThan(0);
    expect(menu?.querySelectorAll('button.fc-tb__menu-item[data-symbol]').length).toBe(0);
    expect(menu?.querySelectorAll('.fc-tb__menu-item--iconic').length).toBe(1);
    const piButton = Array.from(
      menu?.querySelectorAll<HTMLButtonElement>('[data-symbol]') ?? [],
    ).find((button) => button.dataset.symbol === 'π');
    expect(piButton).toBeTruthy();
    expect(piButton?.classList.contains('fc-tb__symbol-tile')).toBe(true);
    expect(piButton?.querySelector('.fc-tb__symbol-tile__glyph')?.textContent).toBe('π');
    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: piButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);
    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
      kind: 'text',
      value: 'π',
    });

    tb.dispose();
  });

  it('opens More Symbols from the Insert Symbol secondary menu and inserts custom text', async () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('insert');

    const symbolButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="symbolInsert"]',
    );
    expect(symbolButton).toBeTruthy();
    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'symbolInsert', menuId: 'menu-symbol' },
      symbolButton as HTMLButtonElement,
    );
    const moreButton = host.querySelector<HTMLButtonElement>('[data-symbol-action="more"]');
    expect(moreButton).toBeTruthy();
    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: moreButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);

    await Promise.resolve();
    const dialog = document.body.querySelector<HTMLElement>('.fc-tb__dlg');
    expect(dialog?.textContent).toContain('More Symbols');
    const input = dialog?.querySelector<HTMLInputElement>('input');
    expect(input).toBeTruthy();
    if (!input) throw new Error('Expected More Symbols input.');
    input.value = 'Ω';
    dialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await Promise.resolve();
    await Promise.resolve();

    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
      kind: 'text',
      value: 'Ω',
    });

    tb.dispose();
  });
});
