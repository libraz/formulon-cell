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
