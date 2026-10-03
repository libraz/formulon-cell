import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { captureSnapshot } from '../../../../src/commands/clipboard/snapshot.js';
import { commentAt, setComment } from '../../../../src/commands/comment.js';
import { hyperlinkAt, setHyperlink } from '../../../../src/commands/hyperlinks.js';
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

  it('clears comments and hyperlinks through the Home Clear dropdown', () => {
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    setComment(sheet.instance.store, { sheet: 0, row: 0, col: 0 }, 'note', sheet.workbook);
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

    const clearButton = host.querySelector<HTMLButtonElement>(
      '.fc-tb__ribbon-group--editing [data-ribbon-command="clearFormat"]',
    );
    expect(clearButton).toBeTruthy();
    expect(clearButton?.dataset.ribbonActivation).toBe('dropdown');
    expect(clearButton?.dataset.ribbonMenuId).toBe('menu-clear');
    clearButton?.click();
    expect(host.querySelectorAll('#menu-clear .fc-tb__menu-item--iconic').length).toBe(7);
    const clearCommentsButton = host.querySelector<HTMLButtonElement>('[data-clear="comments"]');
    expect(clearCommentsButton).toBeTruthy();
    expect(clearCommentsButton?.getAttribute('aria-disabled')).toBe('false');
    const commentsEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(commentsEvent, 'target', { value: clearCommentsButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(commentsEvent)).toBe(true);
    expect(commentAt(sheet.instance.store.getState(), { sheet: 0, row: 0, col: 0 })).toBeNull();
    expect(hyperlinkAt(sheet.instance.store.getState(), { sheet: 0, row: 0, col: 0 })).toBe(
      'https://example.test',
    );

    clearButton?.click();
    expect(clearCommentsButton?.getAttribute('aria-disabled')).toBe('true');
    expect(clearCommentsButton?.getAttribute('aria-description')).toBe(
      'Nothing matching this clear option is selected.',
    );
    expect(clearCommentsButton?.dataset.menuDisabledReason).toBe(
      'Nothing matching this clear option is selected.',
    );
    const clearHyperlinksButton = host.querySelector<HTMLButtonElement>(
      '[data-clear="remove-hyperlinks"]',
    );
    expect(clearHyperlinksButton).toBeTruthy();
    expect(clearHyperlinksButton?.getAttribute('aria-disabled')).toBe('false');
    expect(clearHyperlinksButton?.getAttribute('aria-description')).toBeNull();
    expect(clearHyperlinksButton?.dataset.menuDisabledReason).toBeUndefined();
    const hyperlinksEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(hyperlinksEvent, 'target', { value: clearHyperlinksButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(hyperlinksEvent)).toBe(true);
    expect(hyperlinkAt(sheet.instance.store.getState(), { sheet: 0, row: 0, col: 0 })).toBeNull();
    clearButton?.click();
    expect(clearHyperlinksButton?.getAttribute('aria-disabled')).toBe('true');
    expect(clearHyperlinksButton?.dataset.menuDisabledReason).toBe(
      'Nothing matching this clear option is selected.',
    );

    tb.dispose();
  });

  it('does not enable Clear Formats for comment-only or empty format metadata', () => {
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    setComment(sheet.instance.store, { sheet: 0, row: 0, col: 0 }, 'note', sheet.workbook);
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    const clearButton = host.querySelector<HTMLButtonElement>(
      '.fc-tb__ribbon-group--editing [data-ribbon-command="clearFormat"]',
    );
    expect(clearButton).toBeTruthy();
    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'clearFormat', menuId: 'menu-clear' },
      clearButton as HTMLButtonElement,
    );
    const clearFormats = host.querySelector<HTMLButtonElement>('[data-clear="formats"]');
    const clearComments = host.querySelector<HTMLButtonElement>('[data-clear="comments"]');
    const clearAll = host.querySelector<HTMLButtonElement>('[data-clear="all"]');
    expect(clearFormats?.getAttribute('aria-disabled')).toBe('true');
    expect(clearComments?.getAttribute('aria-disabled')).toBe('false');
    expect(clearAll?.getAttribute('aria-disabled')).toBe('false');

    setComment(sheet.instance.store, { sheet: 0, row: 0, col: 0 }, '', sheet.workbook);
    sheet.instance.store.setState((s) => {
      const formats = new Map(s.format.formats);
      formats.set('0:0:0', {});
      return { ...s, format: { ...s.format, formats } };
    });
    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'clearFormat', menuId: 'menu-clear' },
      clearButton as HTMLButtonElement,
    );
    expect(clearFormats?.getAttribute('aria-disabled')).toBe('true');
    expect(clearComments?.getAttribute('aria-disabled')).toBe('true');
    expect(clearAll?.getAttribute('aria-disabled')).toBe('true');

    tb.dispose();
  });

  it('clears contents, formats, conditional rules, and all state through the Home Clear dropdown', () => {
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    const clearButton = host.querySelector<HTMLButtonElement>(
      '.fc-tb__ribbon-group--editing [data-ribbon-command="clearFormat"]',
    );
    expect(clearButton).toBeTruthy();
    const clickClear = (action: string): void => {
      tb.dropdownsApi?.openDynamicRibbonDropdown(
        { command: 'clearFormat', menuId: 'menu-clear' },
        clearButton as HTMLButtonElement,
      );
      const button = host.querySelector<HTMLButtonElement>(`[data-clear="${action}"]`);
      expect(button).toBeTruthy();
      expect(button?.disabled).toBe(false);
      const event = new MouseEvent('click', { bubbles: true });
      Object.defineProperty(event, 'target', { value: button });
      expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);
    };

    seedNumber(sheet, 0, 0, 42);
    clickClear('contents');
    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 0 }).kind).toBe('blank');

    mutators.setCellFormat(sheet.instance.store, { sheet: 0, row: 0, col: 0 }, { bold: true });
    clickClear('formats');
    expect(sheet.instance.store.getState().format.formats.get('0:0:0')).toBeUndefined();

    mutators.addConditionalRule(sheet.instance.store, {
      kind: 'data-bar',
      range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
      color: '#638ec6',
    });
    clickClear('conditional');
    expect(sheet.instance.store.getState().conditional.rules).toEqual([]);

    seedNumber(sheet, 0, 0, 7);
    mutators.setCellFormat(sheet.instance.store, { sheet: 0, row: 0, col: 0 }, { italic: true });
    setComment(sheet.instance.store, { sheet: 0, row: 0, col: 0 }, 'note', sheet.workbook);
    setHyperlink(
      sheet.instance.store,
      { sheet: 0, row: 0, col: 0 },
      'https://example.test',
      sheet.workbook,
    );
    mutators.addConditionalRule(sheet.instance.store, {
      kind: 'data-bar',
      range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
      color: '#638ec6',
    });
    clickClear('all');
    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 0 }).kind).toBe('blank');
    expect(sheet.instance.store.getState().format.formats.get('0:0:0')).toBeUndefined();
    expect(commentAt(sheet.instance.store.getState(), { sheet: 0, row: 0, col: 0 })).toBeNull();
    expect(hyperlinkAt(sheet.instance.store.getState(), { sheet: 0, row: 0, col: 0 })).toBeNull();
    expect(sheet.instance.store.getState().conditional.rules).toEqual([]);

    tb.dispose();
  });

  it('enables Clear Formats when the active empty cell only has pending input format', () => {
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 0, col: 0 });
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    const bold = host.querySelector<HTMLButtonElement>('[data-ribbon-command="bold"]');
    bold?.click();
    expect(sheet.instance.store.getState().ui.pendingFormat).toEqual({
      addr: { sheet: 0, row: 0, col: 0 },
      format: { bold: true },
    });

    const clearButton = host.querySelector<HTMLButtonElement>(
      '.fc-tb__ribbon-group--editing [data-ribbon-command="clearFormat"]',
    );
    expect(clearButton).toBeTruthy();
    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'clearFormat', menuId: 'menu-clear' },
      clearButton,
    );
    const clearFormats = host.querySelector<HTMLButtonElement>('[data-clear="formats"]');
    expect(clearFormats?.disabled).toBe(false);
    expect(clearFormats?.getAttribute('aria-disabled')).toBe('false');

    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: clearFormats });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);
    expect(sheet.instance.store.getState().ui.pendingFormat).toBeNull();
    expect(sheet.instance.store.getState().format.formats.get('0:0:0')).toBeUndefined();

    tb.dispose();
  });

  it('enables Clear All when the active empty cell only has pending input format', () => {
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 0, col: 0 });
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    host.querySelector<HTMLButtonElement>('[data-ribbon-command="bold"]')?.click();
    host.querySelector<HTMLButtonElement>('[data-ribbon-command="italic"]')?.click();
    expect(sheet.instance.store.getState().ui.pendingFormat).toEqual({
      addr: { sheet: 0, row: 0, col: 0 },
      format: { bold: true, italic: true },
    });

    const clearButton = host.querySelector<HTMLButtonElement>(
      '.fc-tb__ribbon-group--editing [data-ribbon-command="clearFormat"]',
    );
    expect(clearButton).toBeTruthy();
    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'clearFormat', menuId: 'menu-clear' },
      clearButton as HTMLButtonElement,
    );
    const clearAll = host.querySelector<HTMLButtonElement>('[data-clear="all"]');
    expect(clearAll?.disabled).toBe(false);
    expect(clearAll?.getAttribute('aria-disabled')).toBe('false');

    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: clearAll });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);
    expect(sheet.instance.store.getState().ui.pendingFormat).toBeNull();
    expect(sheet.instance.store.getState().format.formats.get('0:0:0')).toBeUndefined();
    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 0 }).kind).toBe('blank');

    tb.dispose();
  });

  it('clears pending Borders formatting through the shared Clear Formats dropdown', () => {
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 0, col: 0 });
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    const borders = host.querySelector<HTMLButtonElement>(
      '.fc-tb__ribbon-group--font [data-ribbon-command="borders"]',
    );
    expect(borders).toBeTruthy();
    borders?.click();
    const bottomBorder = host.querySelector<HTMLButtonElement>('[data-border-preset="bottom"]');
    expect(bottomBorder).toBeTruthy();
    bottomBorder?.click();
    expect(sheet.instance.store.getState().ui.pendingFormat).toEqual({
      addr: { sheet: 0, row: 0, col: 0 },
      format: { borders: { bottom: { style: 'thin' } } },
    });
    expect(sheet.instance.store.getState().format.formats.get('0:0:0')).toBeUndefined();

    const clearButton = host.querySelector<HTMLButtonElement>(
      '.fc-tb__ribbon-group--editing [data-ribbon-command="clearFormat"]',
    );
    expect(clearButton).toBeTruthy();
    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'clearFormat', menuId: 'menu-clear' },
      clearButton,
    );
    const clearFormats = host.querySelector<HTMLButtonElement>('[data-clear="formats"]');
    expect(clearFormats?.disabled).toBe(false);

    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: clearFormats });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);
    expect(sheet.instance.store.getState().ui.pendingFormat).toBeNull();
    expect(sheet.instance.store.getState().format.formats.get('0:0:0')).toBeUndefined();

    tb.dispose();
  });

  it('updates the Clear menu for huge selections by scanning materialized entries only', () => {
    seedText(sheet, 900_000, 0, 'far');
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 1048575, c1: 0 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    const clearButton = host.querySelector<HTMLButtonElement>(
      '.fc-tb__ribbon-group--editing [data-ribbon-command="clearFormat"]',
    );
    expect(clearButton).toBeTruthy();
    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'clearFormat', menuId: 'menu-clear' },
      clearButton,
    );

    const clearContents = host.querySelector<HTMLButtonElement>('[data-clear="contents"]');
    expect(clearContents?.disabled).toBe(false);
    expect(clearContents?.getAttribute('aria-disabled')).toBe('false');

    tb.dispose();
  });

  it('enables Clear Contents for materialized cells in an extra selection only', () => {
    seedText(sheet, 0, 1, 'extra');
    sheet.instance.store.setState((state) => ({
      ...state,
      selection: {
        ...state.selection,
        active: { sheet: 0, row: 0, col: 0 },
        anchor: { sheet: 0, row: 0, col: 0 },
        range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
        extraRanges: [{ sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 }],
      },
    }));
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    const clearButton = host.querySelector<HTMLButtonElement>(
      '.fc-tb__ribbon-group--editing [data-ribbon-command="clearFormat"]',
    );
    expect(clearButton).toBeTruthy();
    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'clearFormat', menuId: 'menu-clear' },
      clearButton,
    );

    const clearContents = host.querySelector<HTMLButtonElement>('[data-clear="contents"]');
    const clearAll = host.querySelector<HTMLButtonElement>('[data-clear="all"]');
    expect(clearContents?.disabled).toBe(false);
    expect(clearAll?.disabled).toBe(false);

    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: clearContents });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);
    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({ kind: 'blank' });

    tb.dispose();
  });

  it('opens the Fill Series dialog from the Fill dropdown and applies it', async () => {
    seedNumber(sheet, 0, 0, 7);
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 0 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    const fillButton = host.querySelector<HTMLButtonElement>('[data-ribbon-command="fillHome"]');
    expect(fillButton).toBeTruthy();
    fillButton?.click();
    expect(host.querySelectorAll('#menu-fill .fc-tb__menu-item--iconic').length).toBe(8);
    const fillDownButton = host.querySelector<HTMLButtonElement>('[data-fill="down"]');
    const fillRightButton = host.querySelector<HTMLButtonElement>('[data-fill="right"]');
    const fillGroupButton = host.querySelector<HTMLButtonElement>('[data-fill="group"]');
    const fillJustifyButton = host.querySelector<HTMLButtonElement>('[data-fill="justify"]');
    const flashFillButton = host.querySelector<HTMLButtonElement>('[data-fill="flash"]');
    expect(fillDownButton?.getAttribute('aria-disabled')).toBe('false');
    expect(fillRightButton?.getAttribute('aria-disabled')).toBe('true');
    expect(fillRightButton?.dataset.menuDisabledReason).toBe(
      'Select more than one column to fill left or right.',
    );
    expect(fillGroupButton?.disabled).toBe(true);
    expect(fillJustifyButton?.disabled).toBe(true);
    expect(flashFillButton).toBeTruthy();
    const seriesButton = host.querySelector<HTMLButtonElement>('[data-fill="series"]');
    expect(seriesButton).toBeTruthy();
    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: seriesButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);
    await new Promise((resolve) => requestAnimationFrame(resolve));

    const dialog = document.body.querySelector<HTMLElement>('.fc-tb__dlg');
    expect(dialog?.textContent).toContain('Series');
    dialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await Promise.resolve();

    expect(sheet.workbook.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({
      kind: 'number',
      value: 7,
    });
    expect(sheet.workbook.getValue({ sheet: 0, row: 2, col: 0 })).toEqual({
      kind: 'number',
      value: 7,
    });

    tb.dispose();
  });

  it('applies Fill dropdown directions through the shared default action', () => {
    seedNumber(sheet, 0, 0, 7);
    seedNumber(sheet, 0, 2, 9);
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 0 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    const fillButton = host.querySelector<HTMLButtonElement>('[data-ribbon-command="fillHome"]');
    expect(fillButton).toBeTruthy();
    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'fillHome', menuId: 'menu-fill' },
      fillButton as HTMLButtonElement,
    );
    const fillDownButton = host.querySelector<HTMLButtonElement>('[data-fill="down"]');
    expect(fillDownButton).toBeTruthy();
    expect(fillDownButton?.disabled).toBe(false);
    const downEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(downEvent, 'target', { value: fillDownButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(downEvent)).toBe(true);
    expect(sheet.workbook.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({
      kind: 'number',
      value: 7,
    });
    expect(sheet.workbook.getValue({ sheet: 0, row: 2, col: 0 })).toEqual({
      kind: 'number',
      value: 7,
    });

    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 2 });
    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'fillHome', menuId: 'menu-fill' },
      fillButton as HTMLButtonElement,
    );
    const fillLeftButton = host.querySelector<HTMLButtonElement>('[data-fill="left"]');
    expect(fillLeftButton).toBeTruthy();
    expect(fillLeftButton?.disabled).toBe(false);
    expect(fillLeftButton?.dataset.menuDisabledReason).toBeUndefined();
    const leftEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(leftEvent, 'target', { value: fillLeftButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(leftEvent)).toBe(true);
    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
      kind: 'number',
      value: 9,
    });
    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({
      kind: 'number',
      value: 9,
    });

    tb.dispose();
  });

  it('keeps the Paste secondary menu iconified under the shared shell', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    const pasteButton = host.querySelector<HTMLButtonElement>('[data-ribbon-command="paste"]');
    expect(pasteButton).toBeTruthy();

    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'paste', menuId: 'menu-paste' },
      pasteButton as HTMLButtonElement,
    );

    expect(host.querySelectorAll('#menu-paste .fc-tb__menu-item--iconic').length).toBe(8);
    expect(
      Array.from(host.querySelectorAll<HTMLButtonElement>('#menu-paste [data-paste-action]'))
        .filter((button) => !button.hidden)
        .map((button) => button.dataset.pasteAction),
    ).toEqual(['all', 'dialog']);
    expect(host.querySelector<HTMLElement>('#menu-paste .fc-tb__menu-sep')?.hidden).toBe(true);
    const pasteSpecial = host.querySelector<HTMLButtonElement>('[data-paste-action="dialog"]');
    const pasteAll = host.querySelector<HTMLButtonElement>('[data-paste-action="all"]');
    expect(pasteSpecial).toBeTruthy();
    expect(pasteAll?.getAttribute('aria-disabled')).toBe('false');
    expect(pasteSpecial?.disabled).toBe(true);
    expect(pasteSpecial?.getAttribute('aria-disabled')).toBe('true');
    expect(pasteSpecial?.getAttribute('aria-description')).toBe(
      'Copy or cut cells before using this paste option.',
    );
    expect(pasteSpecial?.dataset.menuDisabledReason).toBe(
      'Copy or cut cells before using this paste option.',
    );
    expect(pasteSpecial?.title).toContain('Copy or cut cells before using this paste option.');

    tb.dispose();
  });

  it('routes enabled Paste secondary actions through shared clipboard hooks', () => {
    seedNumber(sheet, 0, 0, 12);
    const snapshot = captureSnapshot(sheet.instance.store.getState(), {
      sheet: 0,
      r0: 0,
      c0: 0,
      r1: 0,
      c1: 0,
    });
    expect(snapshot).toBeTruthy();
    if (!snapshot) throw new Error('Expected clipboard snapshot.');
    Object.defineProperty(sheet.instance, 'clipboard', {
      configurable: true,
      value: {
        detach: vi.fn(),
        getSnapshot: () => snapshot,
        runShortcut: vi.fn(),
      },
    });
    const openPasteSpecial = vi
      .spyOn(sheet.instance, 'openPasteSpecial')
      .mockImplementation(() => undefined);
    const pasteSpecial = vi.spyOn(sheet.instance, 'pasteSpecial').mockReturnValue(true);
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    const pasteButton = host.querySelector<HTMLButtonElement>('[data-ribbon-command="paste"]');
    expect(pasteButton).toBeTruthy();

    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'paste', menuId: 'menu-paste' },
      pasteButton as HTMLButtonElement,
    );
    expect(
      Array.from(host.querySelectorAll<HTMLButtonElement>('#menu-paste [data-paste-action]')).some(
        (button) => button.hidden,
      ),
    ).toBe(false);
    const pasteSpecialButton = host.querySelector<HTMLButtonElement>(
      '[data-paste-action="dialog"]',
    );
    const pasteValuesButton = host.querySelector<HTMLButtonElement>('[data-paste-action="values"]');
    expect(pasteSpecialButton?.disabled).toBe(false);
    expect(pasteSpecialButton?.getAttribute('aria-description')).toBeNull();
    expect(pasteSpecialButton?.dataset.menuDisabledReason).toBeUndefined();
    expect(pasteValuesButton?.disabled).toBe(false);
    const dialogEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(dialogEvent, 'target', { value: pasteSpecialButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(dialogEvent)).toBe(true);
    expect(openPasteSpecial).toHaveBeenCalledTimes(1);

    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'paste', menuId: 'menu-paste' },
      pasteButton as HTMLButtonElement,
    );
    const valuesEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(valuesEvent, 'target', { value: pasteValuesButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(valuesEvent)).toBe(true);
    expect(pasteSpecial).toHaveBeenCalledWith({
      operation: 'none',
      skipBlanks: false,
      transpose: false,
      what: 'values',
    });

    tb.dispose();
  });

  it('routes Find & Select dropdown actions through shared dialogs, reports, and selection', async () => {
    seedNumber(sheet, 0, 0, 1);
    seedText(sheet, 1, 1, 'plain');
    sheet.workbook.setFormula({ sheet: 0, row: 2, col: 2 }, '=A1+1');
    sheet.instance.store.setState((state) => {
      const cells = new Map(state.data.cells);
      cells.set(addrKey({ sheet: 0, row: 2, col: 2 }), {
        value: { kind: 'number', value: 2 },
        formula: '=A1+1',
      });
      return { ...state, data: { ...state.data, cells } };
    });
    const openFindReplace = vi
      .spyOn(sheet.instance, 'openFindReplace')
      .mockImplementation(() => undefined);
    const openGoTo = vi.spyOn(sheet.instance, 'openGoTo').mockImplementation(() => undefined);
    const openGoToSpecial = vi
      .spyOn(sheet.instance, 'openGoToSpecial')
      .mockImplementation(() => undefined);
    const openWorkbookObjects = vi
      .spyOn(sheet.instance, 'openWorkbookObjects')
      .mockImplementation(() => undefined);
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    const findButton = host.querySelector<HTMLButtonElement>('[data-ribbon-command="findHome"]');
    expect(findButton).toBeTruthy();
    const clickFindSelect = async (action: string): Promise<void> => {
      tb.dropdownsApi?.openDynamicRibbonDropdown(
        { command: 'findHome', menuId: 'menu-find-select' },
        findButton as HTMLButtonElement,
      );
      const button = host.querySelector<HTMLButtonElement>(`[data-find-select="${action}"]`);
      expect(button).toBeTruthy();
      const event = new MouseEvent('click', { bubbles: true });
      Object.defineProperty(event, 'target', { value: button });
      expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);
      await Promise.resolve();
    };

    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'findHome', menuId: 'menu-find-select' },
      findButton as HTMLButtonElement,
    );
    expect(host.querySelectorAll('#menu-find-select .fc-tb__menu-item--iconic').length).toBe(11);

    await clickFindSelect('find');
    await clickFindSelect('replace');
    await clickFindSelect('go-to');
    await clickFindSelect('go-to-special');
    expect(openFindReplace).toHaveBeenNthCalledWith(1, 'find');
    expect(openFindReplace).toHaveBeenNthCalledWith(2, 'replace');
    expect(openGoTo).toHaveBeenCalledTimes(1);
    expect(openGoToSpecial).toHaveBeenCalledTimes(1);

    await clickFindSelect('object-select');
    await clickFindSelect('selection-pane');
    expect(openWorkbookObjects).toHaveBeenCalledTimes(2);

    await clickFindSelect('conditional-format');
    expect(document.body.querySelector<HTMLElement>('.fc-tb__dlg')?.textContent).toContain(
      'No matching cells were found.',
    );
    document.body.querySelector<HTMLButtonElement>('.fc-tb__dlg .fc-fmtdlg__btn--primary')?.click();

    await clickFindSelect('comments');
    expect(document.body.querySelector<HTMLElement>('.fc-tb__dlg')?.textContent).toContain(
      'No comments or notes were found.',
    );
    document.body.querySelector<HTMLButtonElement>('.fc-tb__dlg .fc-fmtdlg__btn--primary')?.click();

    await clickFindSelect('formulas');

    expect(sheet.instance.store.getState().selection.active).toEqual({
      sheet: 0,
      row: 2,
      col: 2,
    });

    tb.dispose();
  });
});
