import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { commentAt, setComment } from '../../../../src/commands/comment.js';
import { Spreadsheet } from '../../../../src/mount.js';
import { mutators } from '../../../../src/store/store.js';
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

  it('opens Protect Sheet from primary click and keeps protection options secondary', async () => {
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('review');

    const protectButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="protectReview"]',
    );
    expect(protectButton).toBeTruthy();
    expect(protectButton?.dataset.ribbonActivation).toBe('splitPrimary');
    protectButton?.click();
    expect(host.querySelector<HTMLDivElement>('#menu-protect-review')?.hidden).toBe(true);
    const protectDialog = document.body.querySelector<HTMLElement>('.fc-tb__dlg');
    expect(protectDialog?.textContent).toContain('Protect Sheet');
    expect(protectDialog?.textContent).toContain('Allow all users of this worksheet to:');
    for (const label of [
      'Format columns',
      'Format rows',
      'Insert columns',
      'Insert hyperlinks',
      'Delete columns',
      'Delete rows',
      'Use PivotTable reports',
      'Edit objects',
      'Edit scenarios',
    ]) {
      expect(protectDialog?.textContent).toContain(label);
    }
    const protectInputs = Array.from(
      protectDialog?.querySelectorAll<HTMLInputElement>('input') ?? [],
    );
    const passwordInput = protectInputs[0];
    const confirmInput = protectInputs[1];
    expect(passwordInput).toBeTruthy();
    expect(confirmInput).toBeTruthy();
    if (!passwordInput) throw new Error('Expected Protect Sheet password input.');
    if (!confirmInput) throw new Error('Expected Protect Sheet confirmation input.');
    passwordInput.value = 'pw';
    confirmInput.value = 'pw';
    const formatCellsOption = protectInputs.find((input) =>
      input.closest('label')?.textContent?.includes('Format cells'),
    );
    expect(formatCellsOption).toBeTruthy();
    if (!formatCellsOption) throw new Error('Expected Format cells permission option.');
    formatCellsOption.checked = true;
    protectDialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await Promise.resolve();
    expect(sheet.instance.store.getState().protection.protectedSheets.has(0)).toBe(true);
    expect(sheet.instance.store.getState().protection.protectedSheets.get(0)?.password).toBe('pw');
    expect(
      sheet.instance.store.getState().protection.protectedSheets.get(0)?.permissions,
    ).toMatchObject({
      formatCells: true,
      formatColumns: false,
      insertColumns: false,
      insertHyperlinks: false,
      deleteRows: false,
      selectLockedCells: true,
      selectUnlockedCells: true,
      pivotTables: false,
      objects: false,
      scenarios: false,
    });

    protectButton?.click();
    const unprotectDialog = document.body.querySelector<HTMLElement>('.fc-tb__dlg');
    expect(unprotectDialog?.textContent).toContain('Unprotect Sheet');
    const unprotectInput = unprotectDialog?.querySelector<HTMLInputElement>('input');
    expect(unprotectInput).toBeTruthy();
    if (!unprotectInput) throw new Error('Expected Unprotect Sheet password input.');
    unprotectInput.value = 'bad';
    unprotectDialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await Promise.resolve();
    expect(document.body.querySelector<HTMLElement>('.fc-tb__dlg')?.textContent).toContain(
      'The password is incorrect.',
    );
    document.body.querySelector<HTMLButtonElement>('.fc-tb__dlg .fc-fmtdlg__btn--primary')?.click();
    await Promise.resolve();
    expect(sheet.instance.store.getState().protection.protectedSheets.has(0)).toBe(true);

    protectButton?.click();
    const unprotectRetryDialog = document.body.querySelector<HTMLElement>('.fc-tb__dlg');
    const unprotectRetryInput = unprotectRetryDialog?.querySelector<HTMLInputElement>('input');
    expect(unprotectRetryInput).toBeTruthy();
    if (!unprotectRetryInput) throw new Error('Expected retry password input.');
    unprotectRetryInput.value = 'pw';
    unprotectRetryDialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await Promise.resolve();
    expect(sheet.instance.store.getState().protection.protectedSheets.has(0)).toBe(false);

    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'protectReview', menuId: 'menu-protect-review' },
      protectButton as HTMLButtonElement,
    );
    const reviewProtectMenu = host.querySelector<HTMLElement>('#menu-protect-review');
    expect(reviewProtectMenu?.querySelectorAll('.fc-tb__menu-item--iconic').length).toBe(8);
    expect(
      reviewProtectMenu?.querySelector<HTMLButtonElement>('[data-protect-action="protect-sheet"]')
        ?.disabled,
    ).toBe(false);
    expect(
      reviewProtectMenu?.querySelector<HTMLButtonElement>('[data-protect-action="unprotect-sheet"]')
        ?.disabled,
    ).toBe(true);
    const workbookButton = reviewProtectMenu?.querySelector<HTMLButtonElement>(
      '[data-protect-action="protect-workbook"]',
    );
    const unprotectWorkbookButton = reviewProtectMenu?.querySelector<HTMLButtonElement>(
      '[data-protect-action="unprotect-workbook"]',
    );
    const clearAllowedRangesButton = reviewProtectMenu?.querySelector<HTMLButtonElement>(
      '[data-protect-action="clear-allowed-edit-ranges"]',
    );
    expect(workbookButton).toBeTruthy();
    expect(workbookButton?.disabled).toBe(false);
    expect(unprotectWorkbookButton?.disabled).toBe(true);
    expect(unprotectWorkbookButton?.dataset.menuDisabledReason).toBe(
      'Workbook structure is not protected.',
    );
    expect(clearAllowedRangesButton?.disabled).toBe(true);
    expect(clearAllowedRangesButton?.dataset.menuDisabledReason).toBe(
      'There are no allowed edit ranges to clear.',
    );
    const workbookEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(workbookEvent, 'target', { value: workbookButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(workbookEvent)).toBe(true);
    expect(sheet.instance.store.getState().protection.workbookStructure).toEqual({});

    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'protectReview', menuId: 'menu-protect-review' },
      protectButton as HTMLButtonElement,
    );
    expect(workbookButton?.disabled).toBe(true);
    expect(workbookButton?.dataset.menuDisabledReason).toBe(
      'Workbook structure is already protected.',
    );
    expect(unprotectWorkbookButton?.disabled).toBe(false);
    expect(unprotectWorkbookButton?.dataset.menuDisabledReason).toBeUndefined();
    const allowRangeButton = reviewProtectMenu?.querySelector<HTMLButtonElement>(
      '[data-protect-action="allow-edit-ranges"]',
    );
    expect(allowRangeButton).toBeTruthy();
    const allowEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(allowEvent, 'target', { value: allowRangeButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(allowEvent)).toBe(true);
    await waitFor(() => Boolean(document.body.querySelector<HTMLElement>('.fc-tb__dlg')));
    const allowDialog = document.body.querySelector<HTMLElement>('.fc-tb__dlg');
    expect(allowDialog?.textContent).toContain('Range');
    const allowRangeInput = allowDialog?.querySelector<HTMLInputElement>('input');
    expect(allowRangeInput?.value).toBeTruthy();
    document.body.querySelector<HTMLButtonElement>('.fc-tb__dlg .fc-fmtdlg__btn--primary')?.click();
    await waitFor(() => sheet.instance.store.getState().protection.allowedEditRanges.length === 1);
    await waitFor(() =>
      Boolean(
        document.body.querySelector<HTMLElement>('.fc-tb__dlg')?.textContent?.includes('Allowed'),
      ),
    );
    document.body.querySelector<HTMLButtonElement>('.fc-tb__dlg .fc-fmtdlg__btn--primary')?.click();

    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'protectReview', menuId: 'menu-protect-review' },
      protectButton as HTMLButtonElement,
    );
    expect(clearAllowedRangesButton?.disabled).toBe(false);
    expect(clearAllowedRangesButton?.dataset.menuDisabledReason).toBeUndefined();
    const clearAllowedEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(clearAllowedEvent, 'target', { value: clearAllowedRangesButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(clearAllowedEvent)).toBe(true);
    expect(sheet.instance.store.getState().protection.allowedEditRanges).toEqual([]);
    expect(document.body.querySelector<HTMLElement>('.fc-tb__dlg')?.textContent).toContain(
      'Cleared allowed edit ranges',
    );
    document.body.querySelector<HTMLButtonElement>('.fc-tb__dlg .fc-fmtdlg__btn--primary')?.click();

    tb.dispose();
  });

  it('opens Protect Sheet from the View Protect primary button', async () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('view');

    const protectButton = host.querySelector<HTMLButtonElement>('[data-ribbon-command="protect"]');
    expect(protectButton).toBeTruthy();
    expect(protectButton?.dataset.ribbonActivation).toBe('splitPrimary');
    protectButton?.click();

    expect(host.querySelector<HTMLDivElement>('#menu-protect-view')?.hidden).toBe(true);
    const protectDialog = document.body.querySelector<HTMLElement>('.fc-tb__dlg');
    expect(protectDialog?.textContent).toContain('Protect Sheet');
    protectDialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await Promise.resolve();
    expect(sheet.instance.store.getState().protection.protectedSheets.has(0)).toBe(true);

    tb.dispose();
  });

  it('unprotects sheets protected with OOXML strong password hash metadata', async () => {
    mutators.setSheetProtected(sheet.instance.store, 0, true, {
      passwordHash: {
        algorithmName: 'SHA-512',
        hashValue:
          'auSsy34WkQnOLam+2c6zPtbZkU+N88uiaOoz1sI5q58RL3NaIY7C18AbhRNGb15+UxRtDZCp10kT11782Plh7A==',
        saltValue: 'Furur6jnDIFaQBhHQBXzFA==',
        spinCount: 0,
      },
    });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('review');

    const protectButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="protectReview"]',
    );
    expect(protectButton).toBeTruthy();
    protectButton?.click();

    let unprotectDialog = document.body.querySelector<HTMLElement>('.fc-tb__dlg');
    expect(unprotectDialog?.textContent).toContain('Unprotect Sheet');
    let input = unprotectDialog?.querySelector<HTMLInputElement>('input');
    expect(input).toBeTruthy();
    if (!input) throw new Error('Expected Unprotect Sheet password input.');
    input.value = 'wrong';
    unprotectDialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await waitFor(
      () =>
        document.body
          .querySelector<HTMLElement>('.fc-tb__dlg')
          ?.textContent?.includes('The password is incorrect.') === true,
    );
    expect(document.body.querySelector<HTMLElement>('.fc-tb__dlg')?.textContent).toContain(
      'The password is incorrect.',
    );
    document.body.querySelector<HTMLButtonElement>('.fc-tb__dlg .fc-fmtdlg__btn--primary')?.click();
    expect(sheet.instance.store.getState().protection.protectedSheets.has(0)).toBe(true);

    protectButton?.click();
    unprotectDialog = document.body.querySelector<HTMLElement>('.fc-tb__dlg');
    input = unprotectDialog?.querySelector<HTMLInputElement>('input');
    expect(input).toBeTruthy();
    if (!input) throw new Error('Expected retry password input.');
    input.value = 'password';
    unprotectDialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await waitFor(() => !sheet.instance.store.getState().protection.protectedSheets.has(0));

    tb.dispose();
  });

  it('keeps View Protect secondary menu scoped to sheet protection actions', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('view');

    const protectButton = host.querySelector<HTMLButtonElement>('[data-ribbon-command="protect"]');
    expect(protectButton).toBeTruthy();
    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'protect', menuId: 'menu-protect-view' },
      protectButton as HTMLButtonElement,
    );
    const menu = host.querySelector<HTMLDivElement>('#menu-protect-view');
    expect(menu?.querySelectorAll('.fc-tb__menu-item--iconic').length).toBe(2);
    const protectSheet = menu?.querySelector<HTMLButtonElement>(
      '[data-protect-action="protect-sheet"]',
    );
    const unprotectSheet = menu?.querySelector<HTMLButtonElement>(
      '[data-protect-action="unprotect-sheet"]',
    );
    expect(protectSheet).toBeTruthy();
    expect(unprotectSheet).toBeTruthy();
    expect(protectSheet?.disabled).toBe(false);
    expect(unprotectSheet?.disabled).toBe(true);
    expect(menu?.querySelector('[data-protect-action="protect-workbook"]')).toBeNull();
    expect(menu?.querySelector('[data-protect-action="allow-edit-ranges"]')).toBeNull();
    expect(menu?.querySelector('[data-protect-action="lock-cell"]')).toBeNull();

    mutators.setSheetProtected(sheet.instance.store, 0, true);
    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'protect', menuId: 'menu-protect-view' },
      protectButton as HTMLButtonElement,
    );
    expect(protectSheet?.disabled).toBe(true);
    expect(unprotectSheet?.disabled).toBe(false);

    tb.dispose();
  });

  it('deletes the active comment from primary click and keeps delete-all secondary', () => {
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 0, col: 0 });
    setComment(sheet.instance.store, { sheet: 0, row: 0, col: 0 }, 'note', sheet.workbook);
    setComment(sheet.instance.store, { sheet: 0, row: 1, col: 0 }, 'other note', sheet.workbook);
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('review');

    const commentsButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="deleteCommentReview"]',
    );
    expect(commentsButton).toBeTruthy();
    expect(commentsButton?.dataset.ribbonActivation).toBe('splitPrimary');
    commentsButton?.click();

    expect(host.querySelector<HTMLDivElement>('#menu-review-comments')?.hidden).toBe(true);
    expect(commentAt(sheet.instance.store.getState(), { sheet: 0, row: 0, col: 0 })).toBeNull();
    expect(commentAt(sheet.instance.store.getState(), { sheet: 0, row: 1, col: 0 })).toBe(
      'other note',
    );

    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'deleteCommentReview', menuId: 'menu-review-comments' },
      commentsButton as HTMLButtonElement,
    );
    expect(host.querySelectorAll('#menu-review-comments .fc-tb__menu-item--iconic').length).toBe(2);
    const deleteActiveButton = host.querySelector<HTMLButtonElement>(
      '[data-comment-action="delete-active"]',
    );
    expect(deleteActiveButton?.disabled).toBe(true);
    const deleteAllButton = host.querySelector<HTMLButtonElement>(
      '[data-comment-action="delete-all"]',
    );
    expect(deleteAllButton).toBeTruthy();
    expect(deleteAllButton?.disabled).toBe(false);
    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: deleteAllButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);

    expect(commentAt(sheet.instance.store.getState(), { sheet: 0, row: 1, col: 0 })).toBeNull();

    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'deleteCommentReview', menuId: 'menu-review-comments' },
      commentsButton as HTMLButtonElement,
    );
    expect(deleteActiveButton?.disabled).toBe(true);
    expect(deleteAllButton?.disabled).toBe(true);

    setComment(sheet.instance.store, { sheet: 0, row: 0, col: 0 }, 'fresh note', sheet.workbook);
    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'deleteCommentReview', menuId: 'menu-review-comments' },
      commentsButton as HTMLButtonElement,
    );
    expect(deleteActiveButton?.disabled).toBe(false);
    expect(deleteAllButton?.disabled).toBe(false);
    const deleteActiveEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(deleteActiveEvent, 'target', { value: deleteActiveButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(deleteActiveEvent)).toBe(true);
    expect(commentAt(sheet.instance.store.getState(), { sheet: 0, row: 0, col: 0 })).toBeNull();

    tb.dispose();
  });
});
