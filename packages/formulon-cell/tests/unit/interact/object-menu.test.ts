import { afterEach, describe, expect, it, vi } from 'vitest';
import { createObjectMenu, type ObjectMenuEntry } from '../../../src/interact/object-menu.js';

const activeItems = (root: ParentNode): HTMLButtonElement[] =>
  Array.from(root.querySelectorAll<HTMLButtonElement>('.fc-ctxmenu__item--active'));

const keydown = (target: EventTarget, key: string): void => {
  target.dispatchEvent(new KeyboardEvent('keydown', { key, bubbles: true, cancelable: true }));
};

describe('createObjectMenu', () => {
  afterEach(() => {
    document.body.replaceChildren();
  });

  const setup = (
    entries: readonly ObjectMenuEntry[] = [
      { id: 'bringToFront', label: 'Bring to Front', run: vi.fn() },
    ],
  ) => {
    const host = document.createElement('div');
    const opener = document.createElement('button');
    const outside = document.createElement('input');
    document.body.append(host, opener, outside);
    const menu = createObjectMenu(host);
    menu.open(10, 10, entries, {
      label: 'Chart 1',
      restoreFocusTo: opener,
    });
    const root = document.querySelector<HTMLElement>('.fc-ctxmenu--object');
    if (!root) throw new Error('Object menu was not mounted');
    return { menu, root, opener, outside };
  };

  it('closes when focus leaves the menu without stealing outside focus', () => {
    const { menu, outside } = setup();
    expect(menu.isOpen()).toBe(true);
    outside.focus();
    expect(menu.isOpen()).toBe(false);
    expect(document.activeElement).toBe(outside);
  });

  it('stays open while focus remains inside the menu', () => {
    const { menu, root } = setup();
    root.querySelector<HTMLButtonElement>('.fc-ctxmenu__item')?.focus();
    expect(menu.isOpen()).toBe(true);
    menu.detach();
  });

  it('moves focus to a hovered enabled row and keeps exactly that row active', () => {
    const { root } = setup([
      { id: 'sendBack', label: 'Send to Back', run: vi.fn() },
      { id: 'bringToFront', label: 'Bring to Front', run: vi.fn() },
    ]);
    const [first, second] = root.querySelectorAll<HTMLButtonElement>('.fc-ctxmenu__item');

    second?.dispatchEvent(new MouseEvent('mouseenter'));

    expect(document.activeElement).toBe(second);
    expect(activeItems(root)).toEqual([second]);
    expect(first?.classList.contains('fc-ctxmenu__item--active')).toBe(false);
  });

  it('syncs active state for programmatic focus and keyboard navigation', () => {
    const { root } = setup([
      { id: 'sendBack', label: 'Send to Back', run: vi.fn() },
      { id: 'bringToFront', label: 'Bring to Front', run: vi.fn() },
      { id: 'group', label: 'Group', run: vi.fn() },
    ]);
    const items = root.querySelectorAll<HTMLButtonElement>('.fc-ctxmenu__item');

    items[1]?.focus();
    expect(activeItems(root)).toEqual([items[1]]);
    keydown(items[1] as HTMLButtonElement, 'Home');
    expect(document.activeElement).toBe(items[0]);
    expect(activeItems(root)).toEqual([items[0]]);
    keydown(items[0] as HTMLButtonElement, 'End');
    expect(document.activeElement).toBe(items[2]);
    expect(activeItems(root)).toEqual([items[2]]);
  });

  it('clears active state when the root panel itself receives focus', () => {
    const { root } = setup();
    root.querySelector<HTMLButtonElement>('.fc-ctxmenu__item')?.focus();
    expect(activeItems(root)).toHaveLength(1);

    root.focus();

    expect(document.activeElement).toBe(root);
    expect(activeItems(root)).toHaveLength(0);
  });

  it('does not activate an item from Enter or Space after disabled-row hover', () => {
    const firstRun = vi.fn();
    const lastRun = vi.fn();
    const { root } = setup([
      { id: 'sendBack', label: 'Send to Back', run: firstRun },
      { id: 'disabled', label: 'Disabled action', run: vi.fn(), disabled: true },
      { id: 'bringToFront', label: 'Bring to Front', run: lastRun },
    ]);
    const items = root.querySelectorAll<HTMLButtonElement>('.fc-ctxmenu__item');
    const disabled = items[1];
    disabled?.dispatchEvent(new MouseEvent('mouseenter'));

    expect(document.activeElement).toBe(root);
    expect(activeItems(root)).toHaveLength(0);
    keydown(root, 'Enter');
    keydown(root, ' ');
    expect(firstRun).not.toHaveBeenCalled();
    expect(lastRun).not.toHaveBeenCalled();

    keydown(root, 'ArrowDown');
    expect(document.activeElement).toBe(items[0]);
    expect(activeItems(root)).toEqual([items[0]]);
    keydown(items[0] as HTMLButtonElement, 'Enter');
    expect(firstRun).toHaveBeenCalledOnce();
    expect(lastRun).not.toHaveBeenCalled();
  });

  it('clears the previous active row when reopening and preserves close-before-action', () => {
    const { menu, root, opener } = setup([
      { id: 'sendBack', label: 'Send to Back', run: vi.fn() },
      { id: 'bringToFront', label: 'Bring to Front', run: vi.fn() },
    ]);
    root.querySelectorAll<HTMLButtonElement>('.fc-ctxmenu__item')[1]?.focus();
    expect(activeItems(root)).toHaveLength(1);
    menu.close();

    const run = vi.fn(() => expect(menu.isOpen()).toBe(false));
    menu.open(10, 10, [{ id: 'group', label: 'Group', run }], {
      label: 'Chart 1',
      restoreFocusTo: opener,
    });
    const reopenedRoot = document.querySelector<HTMLElement>('.fc-ctxmenu--object');
    expect(reopenedRoot).not.toBeNull();
    expect(activeItems(reopenedRoot as HTMLElement)).toEqual([
      reopenedRoot?.querySelector('.fc-ctxmenu__item'),
    ]);

    reopenedRoot?.querySelector<HTMLButtonElement>('.fc-ctxmenu__item')?.click();
    expect(run).toHaveBeenCalledOnce();
    expect(menu.isOpen()).toBe(false);
    expect(document.activeElement).toBe(opener);
  });
});
