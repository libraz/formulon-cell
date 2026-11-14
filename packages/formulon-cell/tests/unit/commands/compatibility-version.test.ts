import { afterEach, describe, expect, it } from 'vitest';
import { addrKey } from '../../../src/engine/address.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createFormulasDropdownDefaults } from '../../../src/mount/dynamic-dropdown-defaults/formulas.js';
import { toolbarMenuText } from '../../../src/toolbar/menu-text.js';
import type { ApplyRibbonCommandDeps } from '../../../src/toolbar/ribbon/apply-ribbon-command.js';
import { dispatchMacRibbonCommand } from '../../../src/toolbar/ribbon/mac/dispatch.js';
import { createFormulasMenuFactories } from '../../../src/toolbar/ribbon/menus/formulas.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/mount.js';

const A1 = { sheet: 0, row: 0, col: 0 };

describe('Calculation Options > Compatibility Version', () => {
  let sheet: MountedStubSheet | undefined;

  afterEach(() => {
    sheet?.dispose();
    sheet = undefined;
  });

  const setup = async () => {
    sheet = await mountStubSheet({ workbook: await WorkbookHandle.createDefault(), locale: 'en' });
    const { instance } = sheet;
    const menu = createFormulasMenuFactories(toolbarMenuText('en'), 'en').createCalcOptionsMenu();
    const defaults = createFormulasDropdownDefaults(instance);
    const ticked = () => {
      defaults.updateCalcOptionsMenu(menu);
      return [...menu.querySelectorAll<HTMLElement>('[data-calc-option^="compat-"]')]
        .filter((button) => button.getAttribute('aria-checked') === 'true')
        .map((button) => button.dataset.calcOption);
    };
    return { instance, defaults, ticked };
  };

  it('sets the version as one undo step and ticks the engine value', async () => {
    const { instance, defaults, ticked } = await setup();
    const wb = instance.workbook;
    expect([wb.compatibilityVersion(), ticked()]).toEqual([2, ['compat-2']]);

    defaults.applyCalcOptionAction('compat-1');
    expect([wb.compatibilityVersion(), ticked()]).toEqual([1, ['compat-1']]);
    // The held value records nothing, so one undo goes straight back to 2.
    defaults.applyCalcOptionAction('compat-1');
    expect(instance.history.undo()).toBe(true);
    expect([wb.compatibilityVersion(), ticked()]).toEqual([2, ['compat-2']]);
    expect(instance.history.canUndo()).toBe(false);
    expect(instance.history.redo()).toBe(true);
    expect(wb.compatibilityVersion()).toBe(1);

    const reloaded = await WorkbookHandle.loadBytes(wb.save());
    try {
      expect(reloaded.compatibilityVersion()).toBe(1);
    } finally {
      reloaded.dispose();
    }
  });

  it('recalculates the formulas whose text unit changed', async () => {
    const { instance, defaults } = await setup();
    // Version 1 counts UTF-16 units, later versions code points.
    instance.workbook.setText({ sheet: 0, row: 0, col: 1 }, 'a\u{1F600}b');
    instance.workbook.setFormula(A1, '=LEN(B1)');
    // The grid takes recalculated values in a microtask after the recalc.
    const len = async () => {
      await Promise.resolve();
      return instance.store.getState().data.cells.get(addrKey(A1))?.value;
    };
    const before = await len();
    expect(before).toEqual({ kind: 'number', value: 3 });
    defaults.applyCalcOptionAction('compat-1');
    expect(await len()).toEqual({ kind: 'number', value: 4 });
    instance.history.undo();
    expect(await len()).toEqual(before);
  });

  it('runs from the Mac menu leaf too', async () => {
    const { instance } = await setup();
    const deps = {
      inst: instance,
      runtime: { projectFormatToolbar: () => undefined, focusSheet: () => undefined },
    } as unknown as ApplyRibbonCommandDeps;
    dispatchMacRibbonCommand('mac.formulas.calc.compatVersion.3', deps, () => false);
    expect(instance.workbook.compatibilityVersion()).toBe(3);
  });
});
