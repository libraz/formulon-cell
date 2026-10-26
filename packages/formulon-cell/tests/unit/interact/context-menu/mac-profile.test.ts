import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { noteVisibleAt } from '../../../../src/commands/comment.js';
import type { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { en } from '../../../../src/i18n/strings/en.js';
import { ja } from '../../../../src/i18n/strings/ja.js';
import { zh } from '../../../../src/i18n/strings/zh.js';
import {
  attachContextMenu,
  type ContextMenuHandle,
} from '../../../../src/interact/context-menu.js';
import { buildMacCellEntries } from '../../../../src/interact/context-menu-spec.js';
import {
  disposeOverlayPortal,
  overlayPortalFor,
  syncOverlayPortalTheme,
} from '../../../../src/interact/overlay-portal.js';
import { createSpreadsheetStore, type SpreadsheetStore } from '../../../../src/store/store.js';
import { fireContextMenu, item, newWb, setFormat, visibleMenu } from './fixtures.js';

const label = (id: string): string | null | undefined =>
  item(id)?.querySelector('.fc-ctxmenu__label')?.textContent;
const hint = (id: string): string | null | undefined =>
  item(id)?.querySelector('.fc-ctxmenu__hint')?.textContent;

describe('context menu on the Mac platform profile', () => {
  let host: HTMLElement;
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;
  let detach: ContextMenuHandle | null;

  const attach = (strings = ja): void => {
    detach = attachContextMenu({
      host,
      store,
      wb,
      strings,
      onEditComment: () => {},
      onDefineName: () => {},
      onInsertHyperlink: () => {},
      onOpenHyperlink: () => {},
      onToggleWatch: () => {},
    });
  };

  beforeEach(async () => {
    host = document.createElement('div');
    host.tabIndex = -1;
    host.classList.add('fc-host');
    host.dataset.fcPlatform = 'mac';
    document.body.appendChild(host);
    store = createSpreadsheetStore();
    wb = await newWb();
    detach = null;
  });

  afterEach(() => {
    detach?.();
    disposeOverlayPortal(host);
    document.body.replaceChildren();
    vi.restoreAllMocks();
  });

  it('mirrors the host platform onto the overlay portal', () => {
    attach();
    const portal = overlayPortalFor(host);
    expect(portal.dataset.fcPlatform).toBe('mac');
    host.dataset.fcPlatform = 'default';
    syncOverlayPortalTheme(host);
    expect(portal.dataset.fcPlatform).toBe('default');
    delete host.dataset.fcPlatform;
    syncOverlayPortalTheme(host);
    expect(portal.dataset.fcPlatform).toBeUndefined();
  });

  it('renders no mini toolbar and no Windows-only rows', () => {
    attach();
    fireContextMenu(host, 200, 70);
    expect(visibleMenu()).not.toBeNull();
    expect(document.querySelector('.fc-ctxmenu__mini')).toBeNull();
    expect(document.querySelector('.fc-ctxmenu__mini-btn')).toBeNull();
    expect(item('selectAll')).toBeNull();
    expect(item('toggleWatch')).toBeNull();
    expect(item('deleteComment')).toBeNull();
    expect(visibleMenu()?.closest<HTMLElement>('.fc-overlay-portal')?.dataset.fcPlatform).toBe(
      'mac',
    );
  });

  it('uses the Mac Japanese labels in the Mac order', () => {
    attach();
    fireContextMenu(host, 200, 70);
    const root = visibleMenu();
    const labels = Array.from(
      root?.querySelectorAll(':scope > .fc-ctxmenu__item .fc-ctxmenu__label') ?? [],
    ).map((node) => node.textContent);
    expect(labels).toEqual([
      'カット',
      'コピー',
      'ペースト',
      '形式を選択してペースト',
      '挿入...',
      '削除...',
      '数式と値のクリア',
      'フィルター',
      '並べ替え',
      '新しいメモ',
      'セルの書式設定...',
      'ふりがなの編集...',
      '名前の定義...',
      'ハイパーリンク...',
      'ハイパーリンクを開く',
    ]);
  });

  it('switches the note rows when the active cell has a note', () => {
    setFormat(store, 0, 0, { comment: 'memo' });
    attach();
    fireContextMenu(host, 200, 70);
    expect(label('insertComment')).toBe('メモの編集');
    expect(label('deleteComment')).toBe('メモの削除');
    expect(label('showHideNote')).toBe('メモの表示/非表示');
    item('showHideNote')?.click();
    expect(noteVisibleAt(store.getState(), { sheet: 0, row: 0, col: 0 })).toBe(true);
  });

  it('shows Mac shortcut glyphs and drops the Windows-only hints', () => {
    attach();
    fireContextMenu(host, 200, 70);
    expect(hint('cut')).toBe('⌘X');
    expect(hint('copy')).toBe('⌘C');
    expect(hint('paste')).toBe('⌘V');
    expect(hint('formatCells')).toBe('⌘1');
    expect(hint('insertHyperlink')).toBe('⌘K');
    expect(hint('clear')).toBe('');
    expect(hint('insertComment')).toBe('');
    const pasteSpecial = buildMacCellEntries(ja, false).find((e) => e.id === 'pasteSpecialMenu');
    const dialog =
      pasteSpecial?.kind === 'submenu'
        ? pasteSpecial.children.find((e) => e.id === 'pasteSpecial')
        : undefined;
    expect(dialog?.kind === 'item' ? dialog.hint : undefined).toBe('^⌘V');
  });

  it('keeps per-locale Mac wording', () => {
    attach(en);
    fireContextMenu(host, 200, 70);
    expect(label('clear')).toBe('Clear Contents');
    expect(label('insertHyperlink')).toBe('Hyperlink…');
    detach?.();
    disposeOverlayPortal(host);
    attach(zh);
    fireContextMenu(host, 200, 70);
    expect(label('copy')).toBe('拷贝');
  });

  it('starts a pointer-opened menu with no row focused and follows the pointer', () => {
    attach();
    fireContextMenu(host, 200, 70);
    const root = visibleMenu();
    expect(document.activeElement).toBe(root);
    const copy = item('copy');
    copy?.dispatchEvent(new MouseEvent('mouseenter'));
    expect(document.activeElement).toBe(copy);
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }));
    expect(document.activeElement).toBe(item('paste'));
  });

  it('leaves the default profile on the Windows menu', () => {
    host.dataset.fcPlatform = 'default';
    attach();
    fireContextMenu(host, 200, 70);
    expect(document.querySelector('.fc-ctxmenu__mini')).not.toBeNull();
    expect(label('cut')).toBe(ja.contextMenu.cut);
    expect(label('clear')).toBe(ja.contextMenu.clear);
    expect(hint('clear')).toBe('Del');
    expect(item('selectAll')).not.toBeNull();
    expect(item('toggleWatch')).not.toBeNull();
    expect(document.activeElement).toBe(item('cut'));
  });
});
