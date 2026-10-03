import { afterEach, beforeEach, describe, expect, it, type Mock, vi } from 'vitest';
import { History } from '../../../../src/commands/history.js';
import {
  InteractionController,
  registerInteractionController,
} from '../../../../src/commands/interaction-controller.js';
import type {
  CellBatchCommand,
  OperationIntent,
} from '../../../../src/commands/interaction-policy.js';
import { fixedFormPolicy, viewerPolicy } from '../../../../src/commands/interaction-policy.js';
import type { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { en } from '../../../../src/i18n/strings/en.js';
import {
  attachContextMenu,
  type ContextMenuHandle,
} from '../../../../src/interact/context-menu.js';
import type {
  ContextMenuContext,
  ContextMenuInteractionController,
  ContextMenuOptions,
} from '../../../../src/interact/context-menu-options.js';
import { buildCellEntries } from '../../../../src/interact/context-menu-spec.js';
import {
  attachNavigationPolicy,
  navigationBoundsFor,
} from '../../../../src/interact/navigation-policy.js';
import {
  disposeOverlayPortal,
  setOverlayOptions,
} from '../../../../src/interact/overlay-portal.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../../src/store/store.js';
import { fireContextMenu, item, miniItem, newWb, setRange, visibleMenu } from './fixtures.js';

describe('attachContextMenu', () => {
  let host: HTMLElement;
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;
  let detach: ContextMenuHandle | null;
  let onAfterCommit: Mock<() => void>;
  let onFormatDialog: Mock<() => void>;
  let unregisterController: (() => void) | null;

  beforeEach(async () => {
    host = document.createElement('div');
    host.tabIndex = -1;
    document.body.appendChild(host);
    store = createSpreadsheetStore();
    wb = await newWb();
    onAfterCommit = vi.fn<() => void>();
    onFormatDialog = vi.fn<() => void>();
    detach = null;
    unregisterController = null;
  });

  afterEach(() => {
    detach?.();
    disposeOverlayPortal(host);
    unregisterController?.();
    document.body.innerHTML = '';
    vi.restoreAllMocks();
  });

  it('keeps the Excel-compatible cell menu item sequence', () => {
    const ids = buildCellEntries(en)
      .filter((entry) => entry.kind !== 'sep')
      .map((entry) => entry.id);
    expect(ids).toEqual([
      'cut',
      'copy',
      'paste',
      'pasteSpecialMenu',
      'insertCells',
      'insertCopiedCells',
      'deleteCells',
      'clear',
      'filterMenu',
      'sortMenu',
      'insertComment',
      'deleteComment',
      'formatCells',
      'defineName',
      'insertHyperlink',
      'openHyperlink',
      'editPhonetic',
      'toggleWatch',
      'selectAll',
    ]);
  });

  describe('menu opening', () => {
    it('right-click on a cell opens the cell menu', () => {
      detach = attachContextMenu({ host, store, wb, onAfterCommit });
      const e = fireContextMenu(host, 200, 70);
      expect(e.defaultPrevented).toBe(true);
      const menu = visibleMenu();
      expect(menu).not.toBeNull();
      expect(menu?.getAttribute('aria-label')).toBe('コンテキスト メニュー');
      expect(item('copy')).not.toBeNull();
      expect(item('formatCells')).not.toBeNull();
      // No row/col-only items in cell menu.
      expect(item('rowInsertAbove')).toBeNull();
      expect(item('colInsertLeft')).toBeNull();
      expect(document.querySelector('.fc-ctxmenu__mini')).not.toBeNull();
      expect(miniItem('bold')).not.toBeNull();
    });

    it('clamps the root menu inside the viewport through shared positioning', () => {
      Object.defineProperty(window, 'innerWidth', { configurable: true, value: 320 });
      Object.defineProperty(window, 'innerHeight', { configurable: true, value: 180 });
      detach = attachContextMenu({ host, store, wb, onAfterCommit });
      const root = document.querySelector<HTMLElement>('.fc-ctxmenu');
      expect(root).toBeTruthy();
      if (root) {
        Object.defineProperty(root, 'offsetWidth', { configurable: true, value: 292 });
        Object.defineProperty(root, 'offsetHeight', { configurable: true, value: 160 });
      }

      fireContextMenu(host, 310, 170);

      expect(root?.style.left).toBe('24px');
      expect(root?.style.top).toBe('16px');
    });

    it('flips and clamps submenus through the shared context menu positioning', () => {
      Object.defineProperty(window, 'innerWidth', { configurable: true, value: 320 });
      Object.defineProperty(window, 'innerHeight', { configurable: true, value: 160 });
      detach = attachContextMenu({ host, store, wb, onAfterCommit });
      fireContextMenu(host, 200, 70);
      const trigger = document.querySelector<HTMLButtonElement>(
        '[data-fc-submenu="pasteSpecialMenu"]',
      );
      const sub = document.querySelector<HTMLElement>('.fc-ctxmenu__sub');
      expect(trigger).toBeTruthy();
      expect(sub).toBeTruthy();
      if (trigger) {
        trigger.getBoundingClientRect = vi.fn(
          () =>
            ({
              left: 250,
              right: 300,
              top: 140,
              bottom: 170,
              width: 50,
              height: 30,
            }) as DOMRect,
        );
      }
      if (sub) {
        Object.defineProperty(sub, 'offsetWidth', { configurable: true, value: 180 });
        Object.defineProperty(sub, 'offsetHeight', { configurable: true, value: 130 });
      }

      trigger?.dispatchEvent(new MouseEvent('mouseenter'));

      expect(sub?.style.display).toBe('block');
      expect(sub?.style.left).toBe('72px');
      expect(sub?.style.top).toBe('26px');
    });

    it('mini toolbar invokes cell formatting actions', () => {
      detach = attachContextMenu({ host, store, wb, onAfterCommit, onFormatDialog });
      fireContextMenu(host, 200, 70);
      miniItem('bold')?.click();

      expect(store.getState().ui.pendingFormat).toEqual({
        addr: { sheet: 0, row: 0, col: 0 },
        format: { bold: true },
      });
      expect(store.getState().format.formats.get('0:0:0')).toBeUndefined();
      expect(visibleMenu()).toBeNull();

      fireContextMenu(host, 200, 70);
      miniItem('formatCells')?.click();
      expect(onFormatDialog).toHaveBeenCalledTimes(1);
    });

    it('does not leave adjacent or trailing separators when optional entries are hidden', () => {
      detach = attachContextMenu({ host, store, wb, onAfterCommit });
      fireContextMenu(host, 200, 70);
      const menu = visibleMenu();
      expect(menu).not.toBeNull();
      const children = Array.from(menu?.children ?? []);
      expect(children[0]?.classList.contains('fc-ctxmenu__sep')).toBe(false);
      expect(children[children.length - 1]?.classList.contains('fc-ctxmenu__sep')).toBe(false);
      for (let i = 1; i < children.length; i += 1) {
        const prevSep = children[i - 1]?.classList.contains('fc-ctxmenu__sep');
        const curSep = children[i]?.classList.contains('fc-ctxmenu__sep');
        expect(prevSep && curSep).toBe(false);
      }
    });

    it('right-click on the row header opens the row menu and promotes selection', () => {
      detach = attachContextMenu({ host, store, wb, onAfterCommit });
      fireContextMenu(host, 10, 30); // row 0 header
      expect(item('rowInsertAbove')).not.toBeNull();
      expect(item('rowDelete')).not.toBeNull();
      const sel = store.getState().selection.range;
      expect(sel).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 16383 });
    });

    it('right-click on the col header opens the col menu and promotes selection', () => {
      detach = attachContextMenu({ host, store, wb, onAfterCommit });
      fireContextMenu(host, 60, 10); // col 0 header
      expect(item('colInsertLeft')).not.toBeNull();
      expect(item('colDelete')).not.toBeNull();
      const sel = store.getState().selection.range;
      expect(sel).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 1048575, c1: 0 });
    });

    it('does not promote selection when the header click falls inside an existing band', () => {
      detach = attachContextMenu({ host, store, wb, onAfterCommit });
      mutators.selectRow(store, 0);
      const before = store.getState().selection.range;
      fireContextMenu(host, 10, 30);
      // Range unchanged.
      expect(store.getState().selection.range).toEqual(before);
    });

    it('keeps a selected row when right-clicking near its top resize edge', () => {
      detach = attachContextMenu({ host, store, wb, onAfterCommit });
      mutators.selectRow(store, 1);
      fireContextMenu(host, 10, 59); // row 1 header, inside resize slack from row 0

      expect(item('rowInsertAbove')).not.toBeNull();
      expect(store.getState().selection.range).toEqual({
        sheet: 0,
        r0: 1,
        c0: 0,
        r1: 1,
        c1: 16383,
      });
    });

    it('uses grid coordinates for row header context hit-testing', () => {
      const grid = document.createElement('div');
      host.appendChild(grid);
      vi.spyOn(host, 'getBoundingClientRect').mockReturnValue({
        left: 0,
        top: 40,
        right: 800,
        bottom: 640,
        width: 800,
        height: 600,
        x: 0,
        y: 40,
        toJSON: () => ({}),
      });
      vi.spyOn(grid, 'getBoundingClientRect').mockReturnValue({
        left: 0,
        top: 200,
        right: 800,
        bottom: 640,
        width: 800,
        height: 440,
        x: 0,
        y: 200,
        toJSON: () => ({}),
      });
      detach = attachContextMenu({ host, grid, store, wb, onAfterCommit });
      mutators.selectRow(store, 1);
      fireContextMenu(host, 10, 240); // grid-local y=40 => row 1 header

      expect(item('rowInsertAbove')).not.toBeNull();
      expect(store.getState().selection.range).toEqual({
        sheet: 0,
        r0: 1,
        c0: 0,
        r1: 1,
        c1: 16383,
      });
    });

    it('right-click inside an existing row band opens the row menu without changing selection', () => {
      detach = attachContextMenu({ host, store, wb, onAfterCommit });
      setRange(store, 1, 0, 3, 16383);
      fireContextMenu(host, 200, 70); // row 1 cell area

      expect(item('rowInsertAbove')).not.toBeNull();
      expect(item('rowDelete')).not.toBeNull();
      expect(store.getState().selection.range).toEqual({
        sheet: 0,
        r0: 1,
        c0: 0,
        r1: 3,
        c1: 16383,
      });
    });

    it('right-click on the formula bar does not open the menu', () => {
      const formulaBar = document.createElement('div');
      formulaBar.className = 'fc-host__formulabar';
      host.appendChild(formulaBar);
      detach = attachContextMenu({ host, store, wb, onAfterCommit });
      const e = new MouseEvent('contextmenu', {
        clientX: 50,
        clientY: 50,
        bubbles: true,
        cancelable: true,
      });
      formulaBar.dispatchEvent(e);
      expect(visibleMenu()).toBeNull();
      expect(e.defaultPrevented).toBe(false);
    });

    it('right-click on the sheet bar does not open the cell menu', () => {
      const sheetBar = document.createElement('div');
      sheetBar.className = 'fc-host__sheetbar';
      host.appendChild(sheetBar);
      detach = attachContextMenu({ host, store, wb, onAfterCommit });
      const e = new MouseEvent('contextmenu', {
        clientX: 50,
        clientY: 50,
        bubbles: true,
        cancelable: true,
      });
      sheetBar.dispatchEvent(e);
      expect(visibleMenu()).toBeNull();
      expect(e.defaultPrevented).toBe(false);
    });

    it('mounts the menu into the host overlay portal so theme travels by cascade', () => {
      host.classList.add('fc-host');
      host.dataset.fcTheme = 'ink';
      detach = attachContextMenu({ host, store, wb, onAfterCommit });

      fireContextMenu(host, 200, 70);
      const menu = visibleMenu();
      const portal = menu?.closest<HTMLElement>('.fc-overlay-portal');
      // The menu lives in the host's themed overlay portal, which carries the
      // host's data-fc-theme — no per-open inline token copying anymore.
      expect(portal).not.toBeNull();
      expect(portal?.dataset.fcTheme).toBe('ink');
      expect(menu?.style.getPropertyValue('--fc-bg-elev')).toBe('');
    });

    it('closes the open menu and rebuilds labels from the next string dictionary', () => {
      detach = attachContextMenu({ host, store, wb, onAfterCommit });
      fireContextMenu(host, 200, 70);
      expect(item('copy')?.textContent).toContain('コピー');

      detach.setStrings(en);

      expect(visibleMenu()).toBeNull();
      fireContextMenu(host, 200, 70);
      expect(item('copy')?.textContent).toContain('Copy');
      expect(item('formatCells')?.textContent).toContain('Format Cells');
    });
  });

  describe('composed menu contract', () => {
    it('selects built-ins and supports reordered, nested host actions', () => {
      const customAction = vi.fn<(context: ContextMenuContext) => void>();
      const options: ContextMenuOptions = {
        mode: 'builtIn',
        items: ['copy', 'clear'],
        transform: (ctx) => [
          { id: 'app.custom', label: 'Custom', action: customAction },
          ...ctx.defaultItems.filter((entry) => entry.id === 'copy'),
          {
            id: 'app.more',
            label: 'More',
            children: [{ id: 'app.nested', label: 'Nested', action: customAction }],
          },
        ],
      };
      detach = attachContextMenu({ host, store, wb, options });

      const event = fireContextMenu(host, 200, 70);
      expect(event.defaultPrevented).toBe(true);
      expect(item('copy')).not.toBeNull();
      expect(item('clear')).toBeNull();
      expect(item('app.custom')).not.toBeNull();
      expect(item('cut')).toBeNull();

      item('app.custom')?.click();
      expect(customAction).toHaveBeenCalledTimes(1);
      expect(customAction.mock.calls[0]?.[0].defaultItems.map((entry) => entry.id)).toEqual([
        'copy',
        'sep1',
        'clear',
      ]);

      fireContextMenu(host, 200, 70);
      document
        .querySelector<HTMLButtonElement>('[data-fc-submenu="app.more"]')
        ?.dispatchEvent(new MouseEvent('mouseenter'));
      document
        .querySelector<HTMLButtonElement>('.fc-ctxmenu__sub [data-fc-action="app.nested"]')
        ?.click();
      expect(customAction).toHaveBeenCalledTimes(2);
    });

    it('does not let a custom item spoof a built-in ID', () => {
      const action = vi.fn();
      detach = attachContextMenu({
        host,
        store,
        wb,
        options: {
          mode: 'builtIn',
          transform: () => [{ id: 'clear', label: 'Spoof', action }],
        },
      });
      fireContextMenu(host, 200, 70);
      expect(item('clear')).toBeNull();
      expect(action).not.toHaveBeenCalled();
    });

    it('routes custom commands with the context-menu origin and refreshes policy state', () => {
      let allowed = true;
      let refresh = (): void => {};
      const execute = vi.fn<(command: CellBatchCommand) => void>();
      const controller: ContextMenuInteractionController = {
        canExecute: (_intent: OperationIntent) =>
          allowed
            ? { allowed: true }
            : { allowed: false, code: 'operationDenied', reason: 'Blocked by host policy.' },
        execute,
        subscribe: (listener) => {
          refresh = listener;
          return () => {};
        },
      };
      const command: CellBatchCommand = {
        type: 'cellBatch',
        operation: 'clear',
        origin: 'instanceApi',
        changes: [],
      };
      detach = attachContextMenu({
        host,
        store,
        wb,
        interactionController: controller,
        options: {
          mode: 'builtIn',
          items: [],
          transform: () => [
            { id: 'app.command', label: 'Command', command },
            { id: 'app.locked', label: 'Always disabled', command, disabled: true },
          ],
        },
      });

      fireContextMenu(host, 200, 70);
      expect(item('app.command')?.disabled).toBe(false);
      expect(item('app.locked')?.disabled).toBe(true);
      item('app.command')?.click();
      expect(execute).toHaveBeenCalledWith({ ...command, origin: 'contextMenu' });

      fireContextMenu(host, 200, 70);
      allowed = false;
      refresh();
      expect(item('app.command')?.disabled).toBe(true);
      expect(item('app.locked')?.disabled).toBe(true);
      item('app.command')?.click();
      expect(execute).toHaveBeenCalledTimes(1);
      allowed = true;
      refresh();
      expect(item('app.command')?.disabled).toBe(false);
      expect(item('app.locked')?.disabled).toBe(true);
    });

    it('passes mouse and keyboard contexts to a host-owned menu', () => {
      const opens: ContextMenuContext[] = [];
      detach = attachContextMenu({
        host,
        store,
        wb,
        options: {
          mode: 'host',
          onOpen: (context) => opens.push(context),
        },
      });
      setRange(store, 1, 1, 3, 3);
      const mouse = fireContextMenu(host, 200, 70);
      expect(mouse.defaultPrevented).toBe(true);
      expect(visibleMenu()).toBeNull();
      expect(opens[0]?.event).toBeInstanceOf(MouseEvent);
      expect(opens[0]?.kind).toBe('cell');
      expect(opens[0]?.selection).toEqual({ sheet: 0, r0: 1, c0: 1, r1: 3, c1: 3 });

      host.focus();
      host.dispatchEvent(new KeyboardEvent('keydown', { key: 'ContextMenu', bubbles: true }));
      expect(opens).toHaveLength(2);
      expect(opens[1]?.event).toBeInstanceOf(KeyboardEvent);
      expect(opens[1]?.cell).toEqual(store.getState().selection.active);
    });

    it('refreshes a resolver overlay root before opening from the keyboard', () => {
      host.classList.add('fc-host');
      const firstRoot = document.createElement('section');
      const secondRoot = document.createElement('section');
      document.body.append(firstRoot, secondRoot);
      let currentRoot = firstRoot;
      setOverlayOptions(host, { root: () => currentRoot });
      detach = attachContextMenu({ host, store, wb, options: { mode: 'builtIn' } });
      expect(firstRoot.querySelector('.fc-ctxmenu')).not.toBeNull();

      currentRoot = secondRoot;
      host.focus();
      host.dispatchEvent(new KeyboardEvent('keydown', { key: 'ContextMenu', bubbles: true }));

      expect(firstRoot.querySelector('.fc-ctxmenu')).toBeNull();
      expect(secondRoot.querySelector('.fc-ctxmenu')).not.toBeNull();
    });

    it('leaves the browser menu and selection untouched when disabled', () => {
      detach = attachContextMenu({ host, store, wb, options: { mode: 'disabled' } });
      const before = store.getState().selection;
      const event = fireContextMenu(host, 200, 70);
      expect(event.defaultPrevented).toBe(false);
      expect(visibleMenu()).toBeNull();
      expect(store.getState().selection).toEqual(before);
    });
  });

  describe('restricted menu policy', () => {
    const registerPolicy = (policy: Parameters<typeof fixedFormPolicy>[0] | 'viewer') => {
      const history = new History();
      const controller = new InteractionController({
        store,
        history,
        getWb: () => wb,
      });
      unregisterController = registerInteractionController(store, controller);
      controller.setPolicy(policy === 'viewer' ? viewerPolicy() : fixedFormPolicy(policy));
      return controller;
    };

    it('disables clear, cut, and row insertion while keeping copy/paste visible', () => {
      const controller = registerPolicy({
        ranges: [{ sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 }],
      });
      const formPolicy = fixedFormPolicy({
        ranges: [{ sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 }],
      });
      controller.setPolicy({
        ...formPolicy,
        operations: { ...formPolicy.operations, clear: false },
      });
      setRange(store, 1, 1, 1, 1);
      detach = attachContextMenu({ host, store, wb, options: { mode: 'builtIn' } });
      fireContextMenu(host, 100, 50);
      expect(item('cut')?.disabled).toBe(true);
      expect(item('clear')?.disabled).toBe(true);
      expect(item('copy')?.disabled).toBe(false);
      expect(item('paste')?.disabled).toBe(false);

      controller.setPolicy(formPolicy);
      expect(item('clear')?.disabled).toBe(false);

      document.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
      fireContextMenu(host, 10, 30);
      expect(item('rowInsertAbove')?.disabled).toBe(true);
    });

    it('does not move selection while selection permission is denied', () => {
      const controller = registerPolicy('viewer');
      controller.setPolicy({ ...viewerPolicy(), selection: false });
      setRange(store, 0, 0, 0, 0);
      detach = attachContextMenu({ host, store, wb, options: { mode: 'builtIn' } });
      fireContextMenu(host, 200, 70);
      expect(store.getState().selection.range).toEqual({
        sheet: 0,
        r0: 0,
        c0: 0,
        r1: 0,
        c1: 0,
      });
    });

    it('preflights a legacy paste matrix against a bounded viewport', async () => {
      const bounds = {
        sheet: 0,
        r0: 0,
        c0: 0,
        r1: 0,
        c1: 0,
      };
      const navigation = attachNavigationPolicy(store, () => wb, { range: bounds });
      const history = new History();
      const controller = new InteractionController({
        store,
        history,
        getWb: () => wb,
        getBounds: () => navigationBoundsFor(store),
      });
      unregisterController = registerInteractionController(store, controller);
      vi.spyOn(navigator.clipboard, 'readText').mockResolvedValue('inside\toutside');
      setRange(store, 0, 0, 0, 0);
      detach = attachContextMenu({ host, store, wb, options: { mode: 'builtIn' } });
      fireContextMenu(host, 200, 70);
      item('paste')?.click();
      await Promise.resolve();
      await Promise.resolve();

      expect(wb.getValue({ sheet: 0, row: 0, col: 0 }).kind).toBe('blank');
      expect(wb.getValue({ sheet: 0, row: 0, col: 1 }).kind).toBe('blank');
      navigation.dispose();
    });
  });

  describe('dismissal', () => {
    it('clicking outside the menu hides it', () => {
      detach = attachContextMenu({ host, store, wb, onAfterCommit });
      fireContextMenu(host, 200, 70);
      expect(visibleMenu()).not.toBeNull();
      document.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
      expect(visibleMenu()).toBeNull();
    });

    it('does not carry the cell mini toolbar into row or column menus after dismissal', () => {
      detach = attachContextMenu({ host, store, wb, onAfterCommit });
      fireContextMenu(host, 200, 70);
      expect(document.querySelector('.fc-ctxmenu__mini')).not.toBeNull();

      document.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
      expect(visibleMenu()).toBeNull();

      fireContextMenu(host, 10, 30);
      expect(visibleMenu()).not.toBeNull();
      expect(item('rowInsertAbove')).not.toBeNull();
      expect(document.querySelector('.fc-ctxmenu__mini')).toBeNull();

      document.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
      fireContextMenu(host, 60, 10);
      expect(item('colInsertLeft')).not.toBeNull();
      expect(document.querySelector('.fc-ctxmenu__mini')).toBeNull();
    });

    it('right-clicking outside the menu hides it without opening another menu', () => {
      detach = attachContextMenu({ host, store, wb, onAfterCommit });
      fireContextMenu(host, 200, 70);
      expect(visibleMenu()).not.toBeNull();
      document.body.dispatchEvent(
        new MouseEvent('contextmenu', { bubbles: true, cancelable: true }),
      );
      expect(visibleMenu()).toBeNull();
    });

    it('Escape hides the menu', () => {
      detach = attachContextMenu({ host, store, wb, onAfterCommit });
      host.focus();
      fireContextMenu(host, 200, 70);
      document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', cancelable: true }));
      expect(visibleMenu()).toBeNull();
      expect(document.activeElement).toBe(host);
    });

    it('Arrow keys move focus through enabled menu items and Enter invokes the focused item', () => {
      detach = attachContextMenu({ host, store, wb, onAfterCommit, onFormatDialog });
      fireContextMenu(host, 200, 70);

      expect(document.activeElement).toBe(item('cut'));
      document.dispatchEvent(new KeyboardEvent('keydown', { key: 'End', cancelable: true }));
      expect(document.activeElement).toBe(item('selectAll'));
      document.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowUp', cancelable: true }));
      expect(document.activeElement).not.toBe(item('selectAll'));
      item('formatCells')?.focus();
      document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', cancelable: true }));

      expect(onFormatDialog).toHaveBeenCalledTimes(1);
      expect(visibleMenu()).toBeNull();
    });

    it('scroll hides the menu', () => {
      detach = attachContextMenu({ host, store, wb, onAfterCommit });
      fireContextMenu(host, 200, 70);
      window.dispatchEvent(new Event('scroll'));
      expect(visibleMenu()).toBeNull();
    });

    it('clicking inside the menu does not dismiss it', () => {
      detach = attachContextMenu({ host, store, wb, onAfterCommit });
      fireContextMenu(host, 200, 70);
      const root = document.querySelector<HTMLElement>('.fc-ctxmenu');
      if (root) {
        const e = new MouseEvent('mousedown', { bubbles: true });
        root.dispatchEvent(e);
      }
      expect(visibleMenu()).not.toBeNull();
    });
  });

  describe('row and column action availability', () => {
    it('disables row unhide when the selected row band has no hidden rows', () => {
      detach = attachContextMenu({ host, store, wb, onAfterCommit, strings: en });
      fireContextMenu(host, 10, 30);

      expect(item('rowUnhide')?.disabled).toBe(true);
      expect(item('rowUnhide')?.dataset.disabledReason).toBe(
        'There are no hidden rows in the selected row band.',
      );
      expect(item('rowUnhide')?.getAttribute('aria-description')).toBe(
        'There are no hidden rows in the selected row band.',
      );
    });

    it('enables row unhide when the selected row band contains hidden rows', () => {
      setRange(store, 0, 0, 2, 16383);
      store.setState((s) => ({
        ...s,
        layout: { ...s.layout, hiddenRows: new Set([1]) },
      }));
      detach = attachContextMenu({ host, store, wb, onAfterCommit, strings: en });
      fireContextMenu(host, 10, 30);

      expect(item('rowUnhide')?.disabled).toBe(false);
      expect(item('rowUnhide')?.dataset.disabledReason).toBeUndefined();
      expect(item('rowUnhide')?.getAttribute('aria-description')).toBeNull();
    });
  });
});
