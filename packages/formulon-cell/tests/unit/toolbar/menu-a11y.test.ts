import { describe, expect, it, vi } from 'vitest';

import {
  focusMenuItem,
  handleMenuKeydown,
  prepareMenu,
  projectDisabledReason,
  projectDisabledState,
} from '../../../src/toolbar/menu-a11y.js';

const button = (text: string): HTMLButtonElement => {
  const el = document.createElement('button');
  el.type = 'button';
  el.textContent = text;
  return el;
};

describe('toolbar/menu-a11y', () => {
  it('projects and clears disabled reasons across title, aria, and dataset surfaces', () => {
    const el = button('Paste');

    projectDisabledReason(el, 'Clipboard is empty', {
      describedById: 'paste-help',
      datasetKey: 'disabledReason',
      titlePrefix: 'Paste',
    });

    expect(el.title).toBe('Paste\nClipboard is empty');
    expect(el.getAttribute('aria-describedby')).toBe('paste-help');
    expect(el.getAttribute('aria-description')).toBeNull();
    expect(el.dataset.disabledReason).toBe('Clipboard is empty');

    projectDisabledReason(el, null, {
      describedById: 'paste-help',
      datasetKey: 'disabledReason',
      titlePrefix: 'Paste',
    });

    expect(el.title).toBe('Paste');
    expect(el.hasAttribute('aria-describedby')).toBe(false);
    expect(el.hasAttribute('aria-description')).toBe(false);
    expect(el.dataset.disabledReason).toBeUndefined();
  });

  it('sets native disabled and aria-disabled while only exposing reasons for disabled controls', () => {
    const el = button('Sort');

    projectDisabledState(el, true, 'Select a range', { datasetKey: 'reason' });

    expect(el.disabled).toBe(true);
    expect(el.getAttribute('aria-disabled')).toBe('true');
    expect(el.getAttribute('aria-description')).toBe('Select a range');
    expect(el.dataset.reason).toBe('Select a range');

    projectDisabledState(el, false, 'Ignored while enabled', { datasetKey: 'reason' });

    expect(el.disabled).toBe(false);
    expect(el.getAttribute('aria-disabled')).toBe('false');
    expect(el.hasAttribute('aria-description')).toBe(false);
    expect(el.dataset.reason).toBeUndefined();
  });

  it('prepares menus and roves focus across enabled top-level buttons only', () => {
    const menu = document.createElement('div');
    const first = button('First');
    const disabled = button('Disabled');
    const ariaDisabled = button('Aria disabled');
    const last = button('Last');
    const submenu = document.createElement('div');
    submenu.setAttribute('role', 'menu');
    submenu.append(button('Nested'));
    disabled.disabled = true;
    ariaDisabled.setAttribute('aria-disabled', 'true');
    menu.append(first, disabled, ariaDisabled, last, submenu);
    document.body.appendChild(menu);

    try {
      prepareMenu(menu, 'More commands');
      focusMenuItem(menu, 'first');

      expect(menu.getAttribute('role')).toBe('menu');
      expect(menu.getAttribute('aria-label')).toBe('More commands');
      expect(first.getAttribute('role')).toBe('menuitem');
      expect(document.activeElement).toBe(first);

      handleMenuKeydown(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }), menu, {
        close: vi.fn(),
      });

      expect(document.activeElement).toBe(last);
      expect(last.tabIndex).toBe(0);
      expect(first.tabIndex).toBe(-1);
    } finally {
      menu.remove();
    }
  });

  it('activates the focused item and restores focus on escape', () => {
    const menu = document.createElement('div');
    const trigger = button('Open');
    const item = button('Command');
    const onClick = vi.fn();
    const close = vi.fn();
    item.addEventListener('click', onClick);
    menu.append(item);
    document.body.append(trigger, menu);

    try {
      prepareMenu(menu);
      focusMenuItem(menu, 0);

      handleMenuKeydown(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }), menu, {
        close,
      });

      expect(onClick).toHaveBeenCalledTimes(1);
      expect(close).not.toHaveBeenCalled();

      handleMenuKeydown(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }), menu, {
        close,
        restoreFocusTo: trigger,
      });

      expect(close).toHaveBeenCalledWith(true);
      expect(document.activeElement).toBe(trigger);
    } finally {
      trigger.remove();
      menu.remove();
    }
  });
});
