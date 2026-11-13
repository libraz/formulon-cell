import { afterEach, describe, expect, it } from 'vitest';
import { listThreadedComments } from '../../../src/commands/comment.js';
import { setProtectedSheet } from '../../../src/commands/protection.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { mutators } from '../../../src/store/store.js';
import type { ApplyRibbonCommandDeps } from '../../../src/toolbar/ribbon/apply-ribbon-command.js';
import { dispatchMacRibbonCommand } from '../../../src/toolbar/ribbon/mac/dispatch.js';
import { macRibbonLabelForCommand } from '../../../src/toolbar/ribbon/mac/labels.js';
import { MAC_RIBBON_MENU_ITEMS } from '../../../src/toolbar/ribbon/mac/menu-manifest.js';
import { projectMacRibbonState } from '../../../src/toolbar/ribbon/mac/state-projection.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/mount.js';

const NOTE_IDS = [
  'mac.review.newNote',
  'mac.review.previousNote',
  'mac.review.nextNote',
  'mac.review.showHideNote',
  'mac.review.showNotes',
  'mac.review.convertNotes',
] as const;

describe('Mac Review > Notes menu', () => {
  let sheet: MountedStubSheet | undefined;

  afterEach(() => {
    sheet?.dispose();
    sheet = undefined;
  });

  it('lists the desktop rows in order', () => {
    const items = MAC_RIBBON_MENU_ITEMS['mac.review.notes'] ?? [];
    expect(items.map((item) => item.id)).toEqual([...NOTE_IDS]);
    expect(items.map((item) => macRibbonLabelForCommand(item.id, 'ja'))).toEqual([
      '新しいメモ',
      '前のメモ',
      '次のメモ',
      'メモの表示/非表示',
      'すべてのメモを表示',
      'コメントに変換',
    ]);
    expect(items.some((item) => item.disabled)).toBe(false);
  });

  it('disables navigation and Show/Hide until the book has notes and the active cell has one', async () => {
    sheet = await mountStubSheet({
      workbook: await WorkbookHandle.createDefault({ preferStub: true }),
      locale: 'ja',
    });
    const { instance } = sheet;
    const host = document.createElement('div');
    const buttons = new Map(
      NOTE_IDS.map((id) => {
        const button = document.createElement('button');
        button.dataset.ribbonCommand = id;
        host.appendChild(button);
        return [id, button] as const;
      }),
    );
    const disabled = () => {
      projectMacRibbonState(host, instance);
      return Object.fromEntries(NOTE_IDS.map((id) => [id, buttons.get(id)?.disabled === true]));
    };

    expect(disabled()).toMatchObject({
      'mac.review.newNote': false,
      'mac.review.previousNote': true,
      'mac.review.nextNote': true,
      'mac.review.showHideNote': true,
    });

    mutators.setCellFormat(instance.store, { sheet: 0, row: 4, col: 4 }, { comment: 'note text' });
    expect(disabled()).toMatchObject({
      'mac.review.previousNote': false,
      'mac.review.nextNote': false,
      'mac.review.showHideNote': true,
    });

    mutators.setActive(instance.store, { sheet: 0, row: 4, col: 4 });
    expect(disabled()['mac.review.showHideNote']).toBe(false);

    expect(buttons.get('mac.review.showNotes')?.getAttribute('aria-pressed')).toBe('false');
    mutators.setCellFormat(instance.store, { sheet: 0, row: 4, col: 4 }, { commentVisible: true });
    projectMacRibbonState(host, instance);
    expect(buttons.get('mac.review.showNotes')?.getAttribute('aria-pressed')).toBe('false');
    mutators.setShowAllNotes(instance.store, true);
    projectMacRibbonState(host, instance);
    expect(buttons.get('mac.review.showNotes')?.getAttribute('aria-pressed')).toBe('true');
  });

  it('turns New Note into Edit Note on a note cell, absent while that note is locked', async () => {
    sheet = await mountStubSheet({
      workbook: await WorkbookHandle.createDefault({ preferStub: true }),
      locale: 'ja',
    });
    const { instance } = sheet;
    const menu = document.createElement('div');
    menu.className = 'fc-tb__menu';
    const button = document.createElement('button');
    button.dataset.ribbonCommand = 'mac.review.newNote';
    button.textContent = '新しいメモ';
    menu.appendChild(button);
    const host = document.createElement('div');
    host.appendChild(menu);

    projectMacRibbonState(host, instance);
    expect([button.textContent, button.hidden]).toEqual(['新しいメモ', false]);

    mutators.setCellFormat(instance.store, { sheet: 0, row: 0, col: 0 }, { comment: 'note text' });
    projectMacRibbonState(host, instance);
    expect([button.textContent, button.hidden]).toEqual(['メモの編集', false]);

    // Protected with Edit objects off: the note can't be edited, so its row is gone.
    setProtectedSheet(instance.store, 0, true);
    projectMacRibbonState(host, instance);
    expect(button.hidden).toBe(true);

    setProtectedSheet(instance.store, 0, false);
    mutators.setActive(instance.store, { sheet: 0, row: 3, col: 3 });
    projectMacRibbonState(host, instance);
    expect([button.textContent, button.hidden]).toEqual(['新しいメモ', false]);
  });

  it('converts every note to a thread after the compact alert, and only once a note exists', async () => {
    sheet = await mountStubSheet({ workbook: await WorkbookHandle.createDefault(), locale: 'en' });
    const { instance } = sheet;
    instance.host.dataset.fcPlatform = 'mac';
    const host = document.createElement('div');
    const button = document.createElement('button');
    button.dataset.ribbonCommand = 'mac.review.convertNotes';
    host.appendChild(button);
    projectMacRibbonState(host, instance);
    expect(button.disabled).toBe(true);

    instance.workbook.setCommentEntry(0, 1, 1, 'Ann', 'Check the total');
    projectMacRibbonState(host, instance);
    expect(button.disabled).toBe(false);

    const deps = {
      inst: instance,
      runtime: { projectFormatToolbar: () => undefined, focusSheet: () => undefined },
    } as unknown as ApplyRibbonCommandDeps;
    const settle = () => new Promise((resolve) => setTimeout(resolve, 0));
    const answer = async (label: string) => {
      dispatchMacRibbonCommand('mac.review.convertNotes', deps, () => false);
      await settle();
      const alert = document.querySelector<HTMLElement>('[role="alertdialog"]');
      const text = [
        alert?.querySelector<HTMLElement>('.fc-fmtdlg__header')?.hidden,
        alert?.querySelector('.fc-tb__dlg__message')?.textContent,
      ];
      [...(alert?.querySelectorAll('button') ?? [])]
        .find((candidate) => candidate.textContent === label)
        ?.click();
      await settle();
      return text;
    };
    const message =
      'When converting notes to comments, any images and formatting inside them will be removed.\n\nDo you want to convert all notes to comments?';
    expect(await answer('Cancel')).toEqual([true, message]);
    expect(instance.workbook.getComments(0)).toHaveLength(1);
    expect(await answer('Convert all notes')).toEqual([true, message]);
    expect(instance.workbook.getComments(0)).toEqual([]);
    expect(
      listThreadedComments(instance.store, 0, instance.workbook).map((entry) => entry.text),
    ).toEqual(['Check the total']);
  });
});
