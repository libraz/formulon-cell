import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { noteVisibleAt } from '../../../../src/commands/comment.js';
import { setProtectedSheet } from '../../../../src/commands/protection.js';
import { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
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
import { fireContextMenu, item, setFormat, visibleMenu } from './fixtures.js';

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
      onPickFromList: () => {},
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
    // The real engine: New Comment needs threaded comments.
    wb = await WorkbookHandle.createDefault();
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
      '新しいコメント',
      'セルの書式設定...',
      'ドロップダウン リストから選択...',
      'ふりがなの表示',
      '名前の定義...',
      'ハイパーリンク...',
      'ハイパーリンクを開く',
    ]);
  });

  it('groups the cell menu as the desktop does, with the note items beside New Comment', () => {
    const groups = (): string[][] => {
      const out: string[][] = [[]];
      for (const node of visibleMenu()?.children ?? []) {
        if (node.classList.contains('fc-ctxmenu__sep')) out.push([]);
        const label = node.querySelector('.fc-ctxmenu__label')?.textContent;
        if (node.classList.contains('fc-ctxmenu__item') && label) out[out.length - 1]?.push(label);
      }
      return out.filter((group) => group.length > 0);
    };
    attach();
    fireContextMenu(host, 200, 70);
    expect(groups()).toEqual([
      ['カット', 'コピー', 'ペースト', '形式を選択してペースト'],
      ['挿入...', '削除...', '数式と値のクリア'],
      ['フィルター', '並べ替え'],
      ['新しいコメント'],
      ['セルの書式設定...', 'ドロップダウン リストから選択...', 'ふりがなの表示', '名前の定義...'],
      ['ハイパーリンク...', 'ハイパーリンクを開く'],
    ]);
    document.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
    setFormat(store, 0, 0, { comment: 'memo' });
    fireContextMenu(host, 200, 70);
    expect(groups()[3]).toEqual([
      '新しいコメント',
      'メモの編集',
      'メモの削除',
      'メモの表示/非表示',
    ]);
  });

  it('opens the cell list with band wording on the select-all corner and selects the sheet', () => {
    attach();
    store.setState((st) => ({
      ...st,
      selection: {
        ...st.selection,
        active: { sheet: 0, row: 3, col: 3 },
        anchor: { sheet: 0, row: 3, col: 3 },
        range: { sheet: 0, r0: 3, c0: 3, r1: 3, c1: 3 },
      },
    }));
    fireContextMenu(host, 10, 10); // corner
    const labels = Array.from(
      visibleMenu()?.querySelectorAll(':scope > .fc-ctxmenu__item .fc-ctxmenu__label') ?? [],
    ).map((node) => node.textContent);
    expect(labels).toEqual([
      'カット',
      'コピー',
      'ペースト',
      '形式を選択してペースト',
      '挿入',
      '削除',
      '数式と値のクリア',
      'フィルター',
      '並べ替え',
      '新しいコメント',
      'セルの書式設定...',
      'ドロップダウン リストから選択...',
      'ふりがなの表示',
      '名前の定義...',
      'ハイパーリンク...',
      'ハイパーリンクを開く',
    ]);
    const sel = store.getState().selection;
    expect(sel.active).toEqual({ sheet: 0, row: 0, col: 0 });
    expect(sel.range).toMatchObject({ r0: 0, c0: 0, r1: 1048575 });
  });

  it('lists the pasted-link and source-format rows in the Paste Special submenu', () => {
    const sub = buildMacCellEntries(ja, false).find((e) => e.id === 'pasteSpecialMenu');
    const labels =
      sub?.kind === 'submenu'
        ? sub.children.map((e) => (e.kind === 'sep' ? '-' : e.kind === 'item' ? e.label : '?'))
        : [];
    expect(labels).toEqual([
      'ペースト',
      '-',
      '数式',
      '数式と数値の書式',
      '元の書式を保持',
      '罫線なし',
      '元の列幅を保持',
      '行/列の入れ替え',
      '-',
      '値',
      '値と数値の書式',
      '-',
      '書式のペースト',
      'リンク貼り付け',
      '-',
      '形式を選択してペースト...',
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

  it("drops a note's edit and delete rows on a protected sheet without Edit objects", () => {
    setFormat(store, 0, 0, { comment: 'memo' });
    setProtectedSheet(store, 0, true);
    attach();
    fireContextMenu(host, 200, 70);
    expect(item('insertComment')).toBeNull();
    expect(item('deleteComment')).toBeNull();
    expect(label('showHideNote')).toBe('メモの表示/非表示');
    document.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
    setProtectedSheet(store, 0, true, { permissions: { objects: true } });
    fireContextMenu(host, 200, 70);
    expect(label('insertComment')).toBe('メモの編集');
    expect(label('deleteComment')).toBe('メモの削除');
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
    expect(hint('newComment')).toBe('');
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

  describe('header band menus', () => {
    const labelsOf = (): string[] =>
      Array.from(
        visibleMenu()?.querySelectorAll(':scope > .fc-ctxmenu__item .fc-ctxmenu__label') ?? [],
      ).map((node) => node.textContent ?? '');

    it('column menu: bare Insert/Delete, Hide and Unhide with their hints, no filter or sort', () => {
      attach();
      fireContextMenu(host, 60, 10);
      expect(visibleMenu()?.dataset.fcMenu).toBe('col');
      expect(labelsOf()).toEqual([
        'カット',
        'コピー',
        'ペースト',
        '形式を選択してペースト',
        '挿入',
        '削除',
        '数式と値のクリア',
        'セルの書式設定...',
        '列の幅...',
        '表示しない',
        '再表示',
      ]);
      expect(hint('colHide')).toBe('^0');
      expect(hint('colUnhide')).toBe('^⇧0');
      expect(item('colInsertRight')).toBeNull();
    });

    it('row menu: row height, no insert-below row, hints ^9 and ^⇧9', () => {
      attach();
      fireContextMenu(host, 10, 30);
      expect(visibleMenu()?.dataset.fcMenu).toBe('row');
      expect(labelsOf()).toEqual([
        'カット',
        'コピー',
        'ペースト',
        '形式を選択してペースト',
        '挿入',
        '削除',
        '数式と値のクリア',
        'セルの書式設定...',
        '行の高さ...',
        '表示しない',
        '再表示',
      ]);
      expect(hint('rowHide')).toBe('^9');
      expect(hint('rowUnhide')).toBe('^⇧9');
      expect(item('rowInsertBelow')).toBeNull();
    });

    it('Unhide stays enabled with nothing hidden in the selection', () => {
      attach();
      fireContextMenu(host, 60, 10);
      expect(item('colUnhide')?.disabled).toBe(false);
      expect(item('colUnhide')?.getAttribute('aria-disabled')).not.toBe('true');
    });

    it('a pending copy turns Insert into the copied-cells row without an ellipsis', () => {
      store.setState((s) => ({
        ...s,
        ui: { ...s.ui, copyRange: { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 } },
      }));
      attach();
      fireContextMenu(host, 60, 10);
      expect(label('insertCopiedCells')).toBe('コピーしたセルの挿入');
      expect(item('colInsertLeft')).toBeNull();
    });
  });

  it('a pending copy turns the cell menu Insert into the copied-cells row with an ellipsis', () => {
    store.setState((s) => ({
      ...s,
      ui: { ...s.ui, copyRange: { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 } },
    }));
    attach();
    fireContextMenu(host, 200, 70);
    expect(label('insertCopiedCells')).toBe('コピーしたセルの挿入...');
    expect(item('insertCells')).toBeNull();
  });

  describe('placement', () => {
    const stubMenuRect = (): void => {
      vi.stubGlobal('innerWidth', 3840);
      vi.stubGlobal('innerHeight', 1526);
      vi.spyOn(HTMLElement.prototype, 'getBoundingClientRect').mockImplementation(function (
        this: HTMLElement,
      ) {
        const isMenu = this.classList.contains('fc-ctxmenu');
        const w = isMenu ? 287 : 0;
        const h = isMenu ? 650 : 0;
        return { x: 0, y: 0, left: 0, top: 0, right: w, bottom: h, width: w, height: h } as DOMRect;
      });
    };
    const position = (): { x: number; y: number } => {
      const root = visibleMenu();
      return {
        x: Number.parseFloat(root?.style.left ?? ''),
        y: Number.parseFloat(root?.style.top ?? ''),
      };
    };

    afterEach(() => vi.unstubAllGlobals());

    it('opens one pixel right of and above the pointer', () => {
      stubMenuRect();
      attach();
      fireContextMenu(host, 155, 292);
      expect(position()).toEqual({ x: 156, y: 291 });
    });

    it('flips to the pointer left at the right edge and slides up at the bottom edge', () => {
      stubMenuRect();
      attach();
      fireContextMenu(host, 3800, 1280);
      expect(position()).toEqual({ x: 3800 - 287, y: 1526 - 650 - 6 });
    });
  });
});
