import { afterEach, describe, expect, it, vi } from 'vitest';
import { chooseCompatibilityVersion } from '../../../src/commands/compatibility-version.js';
import { History } from '../../../src/commands/history.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createFormulasDropdownDefaults } from '../../../src/mount/dynamic-dropdown-defaults/formulas.js';
import { toolbarMenuText } from '../../../src/toolbar/menu-text.js';
import { createFormulasMenuFactories } from '../../../src/toolbar/ribbon/menus/formulas.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/mount.js';

describe('Compatibility Version on a book stamped above 3', () => {
  let sheet: MountedStubSheet | undefined;

  afterEach(() => {
    sheet?.dispose();
    sheet = undefined;
  });

  it('lists the stamp as a ticked Unsupported entry and keeps every choice off', async () => {
    sheet = await mountStubSheet({ workbook: await WorkbookHandle.createDefault(), locale: 'en' });
    const { instance } = sheet;
    vi.spyOn(instance.workbook, 'compatibilityVersion').mockReturnValue(4);
    const menu = createFormulasMenuFactories(toolbarMenuText('en'), 'en').createCalcOptionsMenu();
    createFormulasDropdownDefaults(instance).updateCalcOptionsMenu(menu);
    const rows = [...menu.querySelectorAll<HTMLButtonElement>('[data-calc-option^="compat-"]')]
      .filter((button) => button.dataset.calcOption !== 'compat-learn-more')
      .map((button) => [
        button.textContent?.trim(),
        button.hidden,
        button.getAttribute('aria-checked'),
        button.disabled,
      ]);
    expect(rows).toEqual([
      ['Version 1', false, 'false', true],
      ['Version 2 (Recommended)', false, 'false', true],
      ['Version 4 (Unsupported)', false, 'true', true],
    ]);
  });

  it('refuses a change it could not undo', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      vi.spyOn(wb, 'compatibilityVersion').mockReturnValue(4);
      const set = vi.spyOn(wb, 'setCompatibilityVersion');
      const history = new History();
      expect(chooseCompatibilityVersion(wb, history, 2)).toBe(false);
      expect(set).not.toHaveBeenCalled();
      expect(history.canUndo()).toBe(false);
    } finally {
      wb.dispose();
    }
  });
});
