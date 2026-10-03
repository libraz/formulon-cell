import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it, vi } from 'vitest';
import {
  focusMenuItem,
  handleMenuKeydown,
  prepareMenu,
  projectDisabledReason,
  projectDisabledState,
} from '../../../../src/toolbar/menu-a11y.js';
import { createMenu, menuIconButton } from '../../../../src/toolbar/ribbon/menus/general.js';
import { disabledStateAuditDirs, mountDir, root, sourceFilesUnder } from './fixtures.js';

describe('toolbar/ribbon menu primitives', () => {
  it('keeps menu disabled state projection centralized in dropdown defaults', () => {
    const defaultsSource = readFileSync(join(mountDir, 'dynamic-dropdowns-defaults.ts'), 'utf8');

    expect(defaultsSource).toContain('const setMenuControlDisabled');
    expect(defaultsSource).toContain('projectDisabledState(button, disabled');
    expect(defaultsSource.match(/\.disabled\s*=/g) ?? []).toHaveLength(0);
    expect(defaultsSource.match(/setAttribute\('aria-disabled'/g) ?? []).toHaveLength(0);
    expect(defaultsSource).not.toContain("setAttribute('aria-description'");
    expect(defaultsSource).not.toContain('dataset.menuDisabledReason');
  });

  it('keeps disabled control state mutation centralized in menu-a11y', () => {
    const allowedFiles = new Set(['src/toolbar/menu-a11y.ts']);
    const files = disabledStateAuditDirs.flatMap(sourceFilesUnder);
    const violations: string[] = [];

    for (const file of files) {
      if (allowedFiles.has(file)) continue;
      const source = readFileSync(join(root, file), 'utf8');
      const lines = source.split('\n');
      lines.forEach((line, index) => {
        if (/\.disabled\s*=(?!=)/.test(line) || /setAttribute\((['"])aria-disabled\1/.test(line)) {
          violations.push(`${file}:${index + 1}: ${line.trim()}`);
        }
      });
    }

    expect(violations).toEqual([]);
  });

  it('keeps shared menu a11y disabled checks aligned with aria-disabled', () => {
    const menuA11ySource = readFileSync(join(root, 'src/toolbar/menu-a11y.ts'), 'utf8');

    expect(menuA11ySource).toContain('!item.disabled');
    expect(menuA11ySource).toContain("item.getAttribute('aria-disabled') !== 'true'");
  });

  it('projects disabled reasons through the shared helper', () => {
    const button = document.createElement('button');
    projectDisabledReason(button, 'Unavailable', { datasetKey: 'menuDisabledReason' });
    expect(button.title).toBe('Unavailable');
    expect(button.getAttribute('aria-description')).toBe('Unavailable');
    expect(button.dataset.menuDisabledReason).toBe('Unavailable');

    projectDisabledReason(button, null, { datasetKey: 'menuDisabledReason' });
    expect(button.title).toBe('');
    expect(button.getAttribute('aria-description')).toBeNull();
    expect(button.dataset.menuDisabledReason).toBeUndefined();

    projectDisabledReason(button, 'Coming soon', {
      datasetKey: 'ribbonDisabledReason',
      titlePrefix: 'Automate',
    });
    expect(button.title).toBe('Automate\nComing soon');
    expect(button.getAttribute('aria-description')).toBe('Coming soon');
    expect(button.dataset.ribbonDisabledReason).toBe('Coming soon');
    projectDisabledReason(button, null, {
      datasetKey: 'ribbonDisabledReason',
      titlePrefix: 'Automate',
    });
    expect(button.title).toBe('Automate');
    expect(button.dataset.ribbonDisabledReason).toBeUndefined();

    const input = document.createElement('input');
    projectDisabledReason(input, 'Read only', {
      ariaDescription: false,
      describedById: 'readonly-note',
    });
    expect(input.title).toBe('Read only');
    expect(input.getAttribute('aria-describedby')).toBe('readonly-note');
    expect(input.getAttribute('aria-description')).toBeNull();
    projectDisabledReason(input, null, {
      ariaDescription: false,
      describedById: 'readonly-note',
    });
    expect(input.title).toBe('');
    expect(input.getAttribute('aria-describedby')).toBeNull();
  });

  it('projects disabled control state through the shared helper', () => {
    const button = document.createElement('button');
    projectDisabledState(button, true, 'Unavailable', {
      datasetKey: 'disabledReason',
      titlePrefix: 'Insert Function',
    });

    expect(button.disabled).toBe(true);
    expect(button.getAttribute('aria-disabled')).toBe('true');
    expect(button.title).toBe('Insert Function\nUnavailable');
    expect(button.dataset.disabledReason).toBe('Unavailable');

    projectDisabledState(button, false, 'Unavailable', {
      datasetKey: 'disabledReason',
      titlePrefix: 'Insert Function',
    });

    expect(button.disabled).toBe(false);
    expect(button.getAttribute('aria-disabled')).toBe('false');
    expect(button.title).toBe('Insert Function');
    expect(button.dataset.disabledReason).toBeUndefined();
  });

  it('skips aria-disabled menu buttons during shared roving focus', () => {
    const menu = createMenu('menu-a11y-disabled-test');
    const ariaDisabled = menuIconButton('Disabled', 'clear', 'formats', 'clear-formats');
    const enabled = menuIconButton('Enabled', 'clear', 'contents', 'clear-contents');
    ariaDisabled.setAttribute('aria-disabled', 'true');
    menu.append(ariaDisabled, enabled);
    document.body.appendChild(menu);
    prepareMenu(menu);

    focusMenuItem(menu);
    expect(document.activeElement).toBe(enabled);
    expect(ariaDisabled.tabIndex).toBe(-1);
    expect(enabled.tabIndex).toBe(0);

    const key = new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true });
    const clickDisabled = vi.fn();
    const clickEnabled = vi.fn();
    ariaDisabled.addEventListener('click', clickDisabled);
    enabled.addEventListener('click', clickEnabled);
    Object.defineProperty(key, 'target', { value: menu });
    handleMenuKeydown(key, menu, { close: vi.fn() });
    expect(clickDisabled).not.toHaveBeenCalled();
    expect(clickEnabled).toHaveBeenCalledTimes(1);
    menu.remove();
  });
});
