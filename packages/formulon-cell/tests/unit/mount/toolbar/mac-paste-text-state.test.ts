import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { dictionaries } from '../../../../src/i18n/strings.js';
import { createClipboardDropdownDefaults } from '../../../../src/mount/dynamic-dropdown-defaults/clipboard.js';
import { createPasteMenu } from '../../../../src/toolbar/ribbon/menus/paste.js';
import { type MountedStubSheet, mountStubSheet } from '../../../test-utils/mount.js';

// Excel for Mac with only text on the clipboard.
describe('Mac Paste menu in the text state', () => {
  let sheet: MountedStubSheet;

  beforeEach(async () => {
    sheet = await mountStubSheet({ locale: 'ja' });
  });
  afterEach(() => {
    sheet.dispose();
    document.body.replaceChildren();
  });

  const visibleRows = (menu: HTMLElement): string[] =>
    Array.from(menu.children)
      .filter((el) => !(el as HTMLElement).hidden)
      .map((el) => {
        if (el.classList.contains('fc-tb__menu-sep')) return '---';
        const disabled = el.getAttribute('aria-disabled') === 'true' ? '[DIS]' : '';
        return `${disabled}${(el.textContent ?? '').trim()}`;
      });

  it('offers Keep Text Only, the text import wizard and Paste Special', () => {
    const menu = createPasteMenu(dictionaries.ja, 'ja', { mac: true });
    createClipboardDropdownDefaults(sheet.instance).updatePasteMenu(menu);
    expect(visibleRows(menu)).toEqual([
      '[DIS]貼り付けのオプション:',
      'テキストのみ保持',
      '---',
      '[DIS]テキスト ファイル ウィザードを使用する...',
      '形式を選択してペースト...',
    ]);
  });

  it('keeps the desktop menu as it was', () => {
    const menu = createPasteMenu(dictionaries.ja, 'ja');
    createClipboardDropdownDefaults(sheet.instance).updatePasteMenu(menu);
    expect(visibleRows(menu)).toEqual(['貼り付け', '形式を選択して貼り付け…']);
  });
});
