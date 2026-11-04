import { afterEach, describe, expect, it } from 'vitest';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { mutators } from '../../../src/store/store.js';
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

  it('lists the desktop rows in order, with Convert to Comments unavailable', () => {
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
    expect(items.find((item) => item.id === 'mac.review.convertNotes')?.disabled).toBe(true);
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
    expect(buttons.get('mac.review.showNotes')?.getAttribute('aria-pressed')).toBe('true');
  });
});
