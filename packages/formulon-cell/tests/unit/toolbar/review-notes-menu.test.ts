import { afterEach, describe, expect, it } from 'vitest';
import { setComment } from '../../../src/commands/comment.js';
import { setProtectedSheet } from '../../../src/commands/protection.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { en } from '../../../src/i18n/strings.js';
import { createReviewDropdownDefaults } from '../../../src/mount/dynamic-dropdown-defaults/review.js';
import { mutators } from '../../../src/store/store.js';
import { createReviewMenuFactories } from '../../../src/toolbar/ribbon/menus/review.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/mount.js';

const ACTIONS = [
  'new-note',
  'previous-note',
  'next-note',
  'toggle-note',
  'toggle-all-notes',
  'convert-notes',
] as const;

describe('Review > Notes menu', () => {
  let sheet: MountedStubSheet | null = null;
  afterEach(() => {
    sheet?.dispose();
    sheet = null;
    document.body.replaceChildren();
  });

  const mountMenu = async () => {
    sheet = await mountStubSheet({ locale: 'en', workbook: await WorkbookHandle.createDefault() });
    const { instance } = sheet;
    const menu = createReviewMenuFactories(en.ribbonMenu, en.ribbon).createReviewNotesMenu();
    document.body.append(menu);
    const defaults = createReviewDropdownDefaults(instance);
    const item = (action: (typeof ACTIONS)[number]) =>
      menu.querySelector<HTMLButtonElement>(
        `[data-comment-action="${action}"]`,
      ) as HTMLButtonElement;
    const refresh = () => defaults.updateReviewCommentsMenu(menu);
    const disabled = () => {
      refresh();
      return Object.fromEntries(ACTIONS.map((action) => [action, item(action).disabled]));
    };
    return { instance, menu, defaults, item, refresh, disabled };
  };

  it('lists the desktop rows in order', async () => {
    const { menu } = await mountMenu();
    expect(
      [...menu.querySelectorAll('[data-comment-action]')].map((button) => [
        (button as HTMLElement).dataset.commentAction,
        button.textContent,
      ]),
    ).toEqual([
      ['new-note', 'New Note'],
      ['previous-note', 'Previous Note'],
      ['next-note', 'Next Note'],
      ['toggle-note', 'Show/Hide Note'],
      ['toggle-all-notes', 'Show All Notes'],
      ['convert-notes', 'Convert to Comments'],
    ]);
  });

  it('reads Edit Note on a note cell, absent while that note is locked', async () => {
    const { instance, item, refresh } = await mountMenu();
    refresh();
    expect([item('new-note').textContent, item('new-note').hidden]).toEqual(['New Note', false]);

    setComment(instance.store, { sheet: 0, row: 0, col: 0 }, 'Check this', instance.workbook);
    refresh();
    expect([item('new-note').textContent, item('new-note').getAttribute('aria-label')]).toEqual([
      'Edit Note',
      'Edit Note',
    ]);

    setProtectedSheet(instance.store, 0, true);
    refresh();
    expect(item('new-note').hidden).toBe(true);

    setProtectedSheet(instance.store, 0, false);
    mutators.setActive(instance.store, { sheet: 0, row: 1, col: 0 });
    refresh();
    expect([item('new-note').textContent, item('new-note').hidden]).toEqual(['New Note', false]);
  });

  it('enables note commands once a note exists and Show/Hide on a note cell', async () => {
    const { instance, disabled } = await mountMenu();
    expect(disabled()).toEqual({
      'new-note': false,
      'previous-note': true,
      'next-note': true,
      'toggle-note': true,
      'toggle-all-notes': false,
      'convert-notes': true,
    });
    setComment(instance.store, { sheet: 0, row: 3, col: 3 }, 'Check this', instance.workbook);
    expect(disabled()).toMatchObject({
      'previous-note': false,
      'next-note': false,
      'toggle-note': true,
      'toggle-all-notes': false,
      'convert-notes': false,
    });
    mutators.setActive(instance.store, { sheet: 0, row: 3, col: 3 });
    expect(disabled()['toggle-note']).toBe(false);
  });

  it('toggles Show All Notes and shows it pressed', async () => {
    const { instance, defaults, item, refresh } = await mountMenu();
    setComment(instance.store, { sheet: 0, row: 3, col: 3 }, 'Check this', instance.workbook);
    refresh();
    expect(item('toggle-all-notes').getAttribute('aria-pressed')).toBe('false');
    defaults.applyReviewCommentAction('toggle-all-notes');
    expect(instance.store.getState().ui.showAllNotes).toBe(true);
    refresh();
    expect(item('toggle-all-notes').getAttribute('aria-pressed')).toBe('true');
  });

  it('converts the notes after the titled confirm on the default chrome', async () => {
    const { instance, defaults } = await mountMenu();
    instance.workbook.setCommentEntry(0, 2, 2, 'Ann', 'Check the total');
    defaults.applyReviewCommentAction('convert-notes');
    await new Promise((resolve) => setTimeout(resolve, 0));
    const dialog = document.querySelector<HTMLElement>('[role="alertdialog"]');
    expect(dialog?.querySelector('.fc-fmtdlg__header')?.textContent).toBe('Convert to Comments');
    [...(dialog?.querySelectorAll('button') ?? [])]
      .find((button) => button.textContent === 'Convert all notes')
      ?.click();
    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(instance.workbook.getComments(0)).toEqual([]);
    expect(instance.workbook.getThreadedComments(0).map((entry) => entry.text)).toEqual([
      'Check the total',
    ]);
  });
});
