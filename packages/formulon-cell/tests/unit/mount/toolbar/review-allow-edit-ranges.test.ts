import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import {
  addAllowedEditRange,
  allowedEditRangesForSheet,
  setProtectedSheet,
  whenSheetPasswordsHashed,
} from '../../../../src/commands/protection.js';
import { createReviewDropdownDefaults } from '../../../../src/mount/dynamic-dropdown-defaults/review.js';
import { type MountedStubSheet, mountStubSheet } from '../../../test-utils/mount.js';
import { waitFor } from './fixtures.js';

const area = (sheet: number, r0: number, c0: number, r1: number, c1: number) => ({
  sheet,
  r0,
  c0,
  r1,
  c1,
});

describe('Protect menu and Allow Edit Ranges while the sheet is protected', () => {
  let sheet: MountedStubSheet;
  beforeEach(async () => {
    sheet = await mountStubSheet({ locale: 'en' });
  });
  afterEach(() => {
    sheet.dispose();
    document.body.innerHTML = '';
  });

  const menu = (): HTMLElement => {
    const el = document.createElement('div');
    for (const action of ['allow-edit-ranges', 'clear-allowed-edit-ranges']) {
      const button = document.createElement('button');
      button.dataset.protectAction = action;
      el.appendChild(button);
    }
    return el;
  };
  const button = (el: HTMLElement, action: string): HTMLButtonElement =>
    el.querySelector(`[data-protect-action="${action}"]`) as HTMLButtonElement;

  it('disables both items with a reason, and refuses the action itself', async () => {
    const { store } = sheet.instance;
    addAllowedEditRange(store, area(0, 0, 0, 0, 0), { title: 'R' });
    setProtectedSheet(store, 0, true);
    const defaults = createReviewDropdownDefaults(sheet.instance);
    const el = menu();
    defaults.updateProtectMenu(el);
    for (const action of ['allow-edit-ranges', 'clear-allowed-edit-ranges']) {
      expect(button(el, action).disabled, action).toBe(true);
      expect(button(el, action).dataset.menuDisabledReason, action).toBe(
        'Unprotect the sheet to change allowed edit ranges.',
      );
    }
    await defaults.applyProtectAction('allow-edit-ranges');
    await defaults.applyProtectAction('clear-allowed-edit-ranges');
    expect(document.querySelector('.fc-tb__dlg')).toBeNull();
    expect(allowedEditRangesForSheet(store.getState(), 0)).toHaveLength(1);
  });

  it('counts only the ranges of the active sheet for Clear', () => {
    const { store } = sheet.instance;
    addAllowedEditRange(store, area(1, 0, 0, 0, 0), { title: 'Elsewhere' });
    const el = menu();
    createReviewDropdownDefaults(sheet.instance).updateProtectMenu(el);
    expect(button(el, 'clear-allowed-edit-ranges').disabled).toBe(true);
    expect(button(el, 'allow-edit-ranges').disabled).toBe(false);
  });

  it('Clear All asks for the password of a guarded range and keeps it on a wrong one', async () => {
    const { store } = sheet.instance;
    addAllowedEditRange(store, area(0, 0, 0, 0, 0), { title: 'Guarded', password: 'pw' });
    await whenSheetPasswordsHashed();
    const defaults = createReviewDropdownDefaults(sheet.instance);
    void defaults.applyProtectAction('clear-allowed-edit-ranges');
    await waitFor(() => Boolean(document.querySelector('.fc-tb__dlg input')));
    (document.querySelector('.fc-tb__dlg input') as HTMLInputElement).value = 'bad';
    (document.querySelector('.fc-tb__dlg .fc-fmtdlg__btn--primary') as HTMLButtonElement).click();
    await waitFor(() =>
      (document.querySelector('.fc-tb__dlg')?.textContent ?? '').includes('incorrect'),
    );
    (document.querySelector('.fc-tb__dlg .fc-fmtdlg__btn--primary') as HTMLButtonElement).click();
    await waitFor(() => document.querySelector('.fc-tb__dlg') === null);
    expect(allowedEditRangesForSheet(store.getState(), 0)).toHaveLength(1);

    void defaults.applyProtectAction('clear-allowed-edit-ranges');
    await waitFor(() => Boolean(document.querySelector('.fc-tb__dlg input')));
    (document.querySelector('.fc-tb__dlg input') as HTMLInputElement).value = 'pw';
    (document.querySelector('.fc-tb__dlg .fc-fmtdlg__btn--primary') as HTMLButtonElement).click();
    await waitFor(() => allowedEditRangesForSheet(store.getState(), 0).length === 0);
  });
});

describe('Allow Edit Ranges manager commit buttons', () => {
  let sheet: MountedStubSheet;
  beforeEach(async () => {
    sheet = await mountStubSheet({ locale: 'en' });
  });
  afterEach(() => {
    sheet.dispose();
    document.body.innerHTML = '';
  });

  const dialogs = (): HTMLElement[] =>
    Array.from(document.querySelectorAll<HTMLElement>('.fc-tb__dlg'));
  const footerLabels = (): string[] =>
    Array.from(dialogs()[0]?.querySelectorAll<HTMLButtonElement>('.fc-fmtdlg__btn') ?? []).map(
      (b) => b.textContent ?? '',
    );
  const footerButton = (label: string): HTMLButtonElement => {
    const found = Array.from(
      dialogs()[0]?.querySelectorAll<HTMLButtonElement>('.fc-fmtdlg__btn') ?? [],
    ).find((b) => b.textContent === label);
    if (!found) throw new Error(`no ${label} button`);
    return found;
  };
  const titles = (): string[] =>
    allowedEditRangesForSheet(sheet.instance.store.getState(), 0).map((e) => e.title);
  /** Opens the manager; the result's `done` settles when it closes. */
  const open = async (): Promise<{ done: Promise<void> }> => {
    const done = Promise.resolve(
      createReviewDropdownDefaults(sheet.instance).applyProtectAction('allow-edit-ranges'),
    );
    await waitFor(() => dialogs().length === 1);
    return { done };
  };
  /** New... with the editor's prefilled selection, confirmed. */
  const addRange = async (): Promise<void> => {
    const before = titles().length;
    footerButton('New...').click();
    await waitFor(() => dialogs().length === 2);
    dialogs()[1]?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await waitFor(() => titles().length === before + 1 && dialogs().length === 1);
  };
  const deleteSelected = async (): Promise<void> => {
    const before = titles().length;
    footerButton('Delete').click();
    await waitFor(() => titles().length === before - 1);
  };

  it('lays out the Windows buttons with Permissions disabled', async () => {
    const { done } = await open();
    expect(footerLabels()).toEqual([
      'New...',
      'Modify...',
      'Delete',
      'Permissions...',
      'Protect Sheet...',
      'OK',
      'Cancel',
      'Apply',
    ]);
    expect(footerButton('Permissions...').disabled).toBe(true);
    expect(footerButton('Protect Sheet...').disabled).toBe(false);
    footerButton('OK').click();
    await done;
  });

  it('Cancel puts back every add and delete made since opening', async () => {
    addAllowedEditRange(sheet.instance.store, area(0, 4, 4, 4, 4), { title: 'Kept' });
    const { done } = await open();
    await deleteSelected();
    await addRange();
    expect(titles()).not.toContain('Kept');
    footerButton('Cancel').click();
    await done;
    expect(titles()).toEqual(['Kept']);
  });

  it('Esc discards like Cancel', async () => {
    const { done } = await open();
    await addRange();
    dialogs()[0]?.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    await done;
    expect(titles()).toEqual([]);
  });

  it('Apply keeps the changes so far, and Cancel then drops only the later ones', async () => {
    const { done } = await open();
    await addRange();
    footerButton('Apply').click();
    expect(dialogs()).toHaveLength(1);
    const applied = titles();
    await addRange();
    footerButton('Cancel').click();
    await done;
    expect(titles()).toEqual(applied);
  });

  it('OK keeps the changes', async () => {
    const { done } = await open();
    await addRange();
    footerButton('OK').click();
    await done;
    expect(titles()).toHaveLength(1);
  });

  it('Protect Sheet... keeps the changes and opens Protect Sheet', async () => {
    const { done } = await open();
    await addRange();
    footerButton('Protect Sheet...').click();
    await waitFor(() => (dialogs()[0]?.textContent ?? '').includes('Protect Sheet'));
    expect(footerLabels()).not.toContain('New...');
    expect(titles()).toHaveLength(1);
    dialogs()[0]?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await done;
    expect(sheet.instance.store.getState().protection.protectedSheets.has(0)).toBe(true);
    expect(titles()).toHaveLength(1);
  });

  it('keeps the OK-only layout on the Mac chrome, where Esc keeps the changes', async () => {
    sheet.instance.host.dataset.fcPlatform = 'mac';
    const { done } = await open();
    expect(footerLabels()).toEqual(['New...', 'Modify...', 'Delete', 'OK']);
    await addRange();
    dialogs()[0]?.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    await done;
    expect(titles()).toHaveLength(1);
  });
});
