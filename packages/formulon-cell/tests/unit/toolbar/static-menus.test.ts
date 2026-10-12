import { afterEach, describe, expect, it } from 'vitest';
import {
  closeStaticRibbonMenus,
  hasOpenStaticRibbonMenu,
  toggleStaticRibbonSubmenu,
} from '../../../src/toolbar/ribbon/static-menus.js';

const rect = (left: number, top: number, width: number, height: number): DOMRect =>
  ({
    bottom: top + height,
    height,
    left,
    right: left + width,
    top,
    width,
    x: left,
    y: top,
    toJSON: () => ({}),
  }) as DOMRect;

const makeStaticMenu = (): {
  host: HTMLDivElement;
  button: HTMLButtonElement;
  menu: HTMLDivElement;
} => {
  const host = document.createElement('div');
  const button = document.createElement('button');
  const menu = document.createElement('div');
  button.dataset.ribbonMenuId = 'static-test-menu';
  button.getBoundingClientRect = () => rect(740, 100, 24, 24);
  menu.id = 'static-test-menu';
  menu.className = 'fc-tb__menu';
  menu.hidden = true;
  host.append(button, menu);
  document.body.appendChild(host);
  return { host, button, menu };
};

afterEach(() => {
  document.body.replaceChildren();
  Object.defineProperty(window, 'innerWidth', { configurable: true, value: initialInnerWidth });
  Object.defineProperty(window, 'innerHeight', { configurable: true, value: initialInnerHeight });
});

const initialInnerWidth = window.innerWidth;
const initialInnerHeight = window.innerHeight;

describe('toolbar/ribbon static menus', () => {
  it('opens a fixed, clamped menu and clears its geometry on close/reopen', () => {
    const { host, button, menu } = makeStaticMenu();
    Object.defineProperty(window, 'innerWidth', { configurable: true, value: 768 });
    Object.defineProperty(window, 'innerHeight', { configurable: true, value: 180 });

    expect(toggleStaticRibbonSubmenu(host, button)).toBe(true);
    expect(menu.hidden).toBe(false);
    expect(menu.style.position).toBe('fixed');
    expect(Number.parseFloat(menu.style.left)).toBeGreaterThanOrEqual(4);
    expect(Number.parseFloat(menu.style.top)).toBeGreaterThanOrEqual(4);
    expect(menu.style.overflowY).toBe('auto');
    expect(hasOpenStaticRibbonMenu(host)).toBe(true);

    expect(toggleStaticRibbonSubmenu(host, button)).toBe(true);
    expect(menu.hidden).toBe(true);
    expect(menu.style.position).toBe('');
    expect(menu.style.left).toBe('');
    expect(menu.style.maxHeight).toBe('');
    expect(hasOpenStaticRibbonMenu(host)).toBe(false);

    expect(toggleStaticRibbonSubmenu(host, button)).toBe(true);
    expect(menu.hidden).toBe(false);
    expect(menu.style.position).toBe('fixed');
    expect(menu.style.left).not.toBe('');
  });

  it('closes another open menu while retaining the requested exception', () => {
    const first = makeStaticMenu();
    const secondButton = document.createElement('button');
    const second = document.createElement('div');
    secondButton.dataset.ribbonMenuId = 'second-static-test-menu';
    secondButton.getBoundingClientRect = () => rect(20, 100, 24, 24);
    second.id = 'second-static-test-menu';
    second.className = 'fc-tb__menu';
    second.hidden = true;
    first.host.append(secondButton, second);

    toggleStaticRibbonSubmenu(first.host, first.button);
    toggleStaticRibbonSubmenu(first.host, secondButton);
    expect(first.menu.hidden).toBe(true);
    expect(second.hidden).toBe(false);

    closeStaticRibbonMenus(first.host, second);
    expect(second.hidden).toBe(false);
    expect(hasOpenStaticRibbonMenu(first.host)).toBe(true);
  });
});
