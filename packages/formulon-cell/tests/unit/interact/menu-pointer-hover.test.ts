import { afterEach, describe, expect, it, vi } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { ja } from '../../../src/i18n/strings.js';
import { attachContextMenu, type ContextMenuHandle } from '../../../src/interact/context-menu.js';
import { createObjectMenu } from '../../../src/interact/object-menu.js';
import { disposeOverlayPortal } from '../../../src/interact/overlay-portal.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';
import { newWb } from './context-menu/fixtures.js';

// A menu opened from the keyboard under a resting pointer keeps keyboard focus: an item takes
// focus when the pointer moves onto it, not when it appears under the pointer (mouseenter alone).
describe('menu hover focus follows pointer movement', () => {
  let detach: ContextMenuHandle | null = null;
  afterEach(() => {
    detach?.();
    detach = null;
    document.body.replaceChildren();
    vi.restoreAllMocks();
  });

  const openCellMenuFromKeyboard = async () => {
    const host = document.createElement('div');
    host.tabIndex = -1;
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    const wb = await newWb();
    detach = attachContextMenu({
      strings: ja,
      host,
      store,
      wb,
      history: new History(),
      onAfterCommit: () => undefined,
    });
    host.focus();
    host.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'F10', shiftKey: true, bubbles: true, cancelable: true }),
    );
    const root = document.querySelector<HTMLElement>('.fc-ctxmenu--managed');
    if (!root) throw new Error('no cell menu');
    return { host, root };
  };
  const enabledItems = (root: HTMLElement): HTMLButtonElement[] =>
    [...root.querySelectorAll<HTMLButtonElement>(':scope > .fc-ctxmenu__item')].filter(
      (item) => !item.disabled && item.getAttribute('aria-disabled') !== 'true',
    );

  it('keeps cell-menu focus when an item appears under a resting pointer', async () => {
    const { host, root } = await openCellMenuFromKeyboard();
    const focused = document.activeElement;
    const item = enabledItems(root).find((candidate) => candidate !== focused);
    if (!item) throw new Error('no second item');
    item.dispatchEvent(new MouseEvent('mouseenter'));
    expect(document.activeElement).toBe(focused);
    item.dispatchEvent(new MouseEvent('mousemove', { bubbles: true }));
    expect(document.activeElement).toBe(item);
    disposeOverlayPortal(host);
  });

  it('opens no submenu under a resting pointer', async () => {
    const { host } = await openCellMenuFromKeyboard();
    const trigger = document.querySelector<HTMLButtonElement>('[data-fc-submenu]');
    const sub = document.querySelector<HTMLElement>('.fc-ctxmenu__sub');
    if (!trigger || !sub) throw new Error('no submenu');
    trigger.dispatchEvent(new MouseEvent('mouseenter'));
    expect(sub.style.display).not.toBe('block');
    trigger.dispatchEvent(new MouseEvent('mousemove', { bubbles: true }));
    expect(sub.style.display).toBe('block');
    disposeOverlayPortal(host);
  });

  it('keeps object-menu focus when an item appears under a resting pointer', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const menu = createObjectMenu(host);
    menu.open(
      10,
      10,
      [
        { id: 'sendBack', label: 'Send to Back', run: vi.fn() },
        { id: 'bringToFront', label: 'Bring to Front', run: vi.fn() },
      ],
      { label: 'Chart 1', restoreFocusTo: host },
    );
    const [first, second] = document.querySelectorAll<HTMLButtonElement>(
      '.fc-ctxmenu--object .fc-ctxmenu__item',
    );
    first?.focus();
    second?.dispatchEvent(new MouseEvent('mouseenter'));
    expect(document.activeElement).toBe(first);
    second?.dispatchEvent(new MouseEvent('mousemove', { bubbles: true }));
    expect(document.activeElement).toBe(second);
    menu.detach();
  });
});
