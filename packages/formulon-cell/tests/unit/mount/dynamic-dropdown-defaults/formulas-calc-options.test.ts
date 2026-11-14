import { afterEach, describe, expect, it, vi } from 'vitest';
import { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { createFormulasDropdownDefaults } from '../../../../src/mount/dynamic-dropdown-defaults/formulas.js';
import { mutators } from '../../../../src/store/store.js';
import { toolbarMenuText } from '../../../../src/toolbar/menu-text.js';
import { createFormulasMenuFactories } from '../../../../src/toolbar/ribbon/menus/formulas.js';
import { type MountedStubSheet, mountStubSheet } from '../../../test-utils/mount.js';

const VERSIONS = ['compat-1', 'compat-2', 'compat-3'] as const;

describe('default profile Calculation Options and Error Checking menus', () => {
  let sheet: MountedStubSheet | undefined;

  afterEach(() => {
    vi.restoreAllMocks();
    sheet?.dispose();
    sheet = undefined;
  });

  const setup = async () => {
    sheet = await mountStubSheet({
      workbook: await WorkbookHandle.createDefault({ preferStub: true }),
      locale: 'en',
    });
    const factories = createFormulasMenuFactories(toolbarMenuText('en'), 'en');
    const calc = factories.createCalcOptionsMenu();
    const errorCheck = factories.createErrorCheckingMenu();
    const defaults = createFormulasDropdownDefaults(sheet.instance);
    const only = (menu: HTMLElement, selector: string): HTMLButtonElement => {
      const found = menu.querySelector<HTMLButtonElement>(selector);
      if (!found) throw new Error(`${selector} was not rendered`);
      return found;
    };
    return { instance: sheet.instance, calc, errorCheck, defaults, only };
  };

  it('places the items after the calc modes and Trace Error, the lists as format submenus', async () => {
    const { calc, errorCheck, only } = await setup();
    expect(
      [...calc.querySelectorAll<HTMLElement>('[data-calc-option], [data-format-submenu]')].map(
        (el) => el.dataset.calcOption ?? `submenu:${el.dataset.formatSubmenu}`,
      ),
    ).toEqual([
      'auto',
      'auto-no-table',
      'manual',
      'format-stale',
      'submenu:compat-version',
      ...VERSIONS,
      'compat-learn-more',
      'calculate-now',
      'calculate-sheet',
      'iterative',
    ]);
    expect(only(calc, '[data-calc-option="format-stale"]').getAttribute('role')).toBe(
      'menuitemcheckbox',
    );
    const panel = calc.querySelector<HTMLElement>('#menu-calc-options-compat-version');
    expect(panel?.dataset.formatPanel).toBe('compat-version');
    expect(
      VERSIONS.map((value) => {
        const button = only(calc, `[data-calc-option="${value}"]`);
        return [button.textContent, button.getAttribute('role')];
      }),
    ).toEqual([
      ['Version 1', 'menuitemradio'],
      ['Version 2 (Recommended)', 'menuitemradio'],
      ['Version 3 (Unsupported)', 'menuitemradio'],
    ]);
    expect(
      [
        ...errorCheck.querySelectorAll<HTMLElement>(
          '[data-formula-audit-action], [data-format-submenu]',
        ),
      ].map((el) => el.dataset.formulaAuditAction ?? `submenu:${el.dataset.formatSubmenu}`),
    ).toEqual(['error-checking', 'trace-error', 'submenu:circular', 'ignore-error']);
    expect(only(errorCheck, '[data-format-submenu="circular"]').textContent).toBe(
      'Circular References',
    );
  });

  it('keeps the new rows disabled while the engine lacks the methods', async () => {
    const { calc, errorCheck, defaults, only } = await setup();
    defaults.updateCalcOptionsMenu(calc);
    defaults.updateErrorCheckingMenu(errorCheck);
    expect(
      [
        only(calc, '[data-calc-option="format-stale"]'),
        only(calc, '[data-format-submenu="compat-version"]'),
        only(calc, '[data-calc-option="compat-1"]'),
        only(calc, '[data-calc-option="compat-learn-more"]'),
        only(errorCheck, '[data-format-submenu="circular"]'),
      ].map((button) => button.getAttribute('aria-disabled')),
    ).toEqual(['true', 'true', 'true', 'true', 'true']);
    expect(only(calc, '[data-calc-option="compat-3"]').hidden).toBe(true);
  });

  it('enables the rows the engine provides and projects the ticks', async () => {
    const { instance, calc, errorCheck, defaults, only } = await setup();
    vi.spyOn(instance.workbook, 'circularReferencesSupported').mockReturnValue(true);
    vi.spyOn(instance.workbook, 'circularReferences').mockReturnValue([{ row: 0, col: 0 }]);
    vi.spyOn(instance.workbook, 'staleValuesSupported').mockReturnValue(true);
    vi.spyOn(instance.workbook, 'compatibilityVersion').mockReturnValue(3);
    const mode = vi.spyOn(instance.workbook, 'calcMode').mockReturnValue(1);
    const state = (selector: string) => {
      const button = only(calc, selector);
      return [
        button.hidden,
        button.getAttribute('aria-disabled'),
        button.getAttribute('aria-checked'),
      ];
    };

    defaults.updateCalcOptionsMenu(calc);
    defaults.updateErrorCheckingMenu(errorCheck);
    expect(only(errorCheck, '[data-format-submenu="circular"]').getAttribute('aria-disabled')).toBe(
      'false',
    );
    expect(state('[data-calc-option="manual"]')[2]).toBe('true');
    expect(state('[data-calc-option="format-stale"]')).toEqual([false, 'false', 'true']);
    expect(VERSIONS.map((value) => state(`[data-calc-option="${value}"]`))).toEqual([
      [false, 'false', 'false'],
      [false, 'false', 'false'],
      [false, 'false', 'true'],
    ]);

    mode.mockReturnValue(0);
    mutators.setFormatStaleValues(instance.store, false);
    defaults.updateCalcOptionsMenu(calc);
    expect(state('[data-calc-option="format-stale"]')).toEqual([false, 'true', 'false']);
  });
});

describe('Error Checking > Circular References', () => {
  let sheet: MountedStubSheet | undefined;

  afterEach(() => {
    vi.restoreAllMocks();
    sheet?.dispose();
    sheet = undefined;
  });

  const setup = async () => {
    sheet = await mountStubSheet({
      workbook: await WorkbookHandle.createDefault({ preferStub: true }),
      locale: 'en',
    });
    const { instance } = sheet;
    instance.workbook.addSheet('My Sheet');
    const menu = createFormulasMenuFactories(toolbarMenuText('en'), 'en').createErrorCheckingMenu();
    const defaults = createFormulasDropdownDefaults(instance);
    vi.spyOn(instance.workbook, 'circularReferencesSupported').mockReturnValue(true);
    const cycles = new Map<number, { row: number; col: number }[]>();
    const list = vi
      .spyOn(instance.workbook, 'circularReferences')
      .mockImplementation((index) => cycles.get(index) ?? []);
    const open = () => {
      defaults.updateErrorCheckingMenu(menu);
      const trigger = menu.querySelector<HTMLButtonElement>('[data-format-submenu="circular"]');
      return {
        disabled: trigger?.getAttribute('aria-disabled'),
        items: [...menu.querySelectorAll<HTMLButtonElement>('[data-circular-reference]')].map(
          (button) => [
            button.textContent,
            button.dataset.circularReference,
            button.getAttribute('role'),
            button.getAttribute('aria-checked'),
          ],
        ),
      };
    };
    return { instance, defaults, cycles, list, open };
  };

  it('lists the active sheet in the engine order, ticking the active cell when listed', async () => {
    const { instance, cycles, open } = await setup();
    cycles.set(0, [
      { row: 2, col: 1 },
      { row: 0, col: 0 },
    ]);
    cycles.set(1, [{ row: 5, col: 5 }]);
    expect(open()).toEqual({
      disabled: 'false',
      items: [
        ['$B$3', '0:2:1', 'menuitemradio', 'false'],
        ['$A$1', '0:0:0', 'menuitemradio', 'true'],
      ],
    });
    mutators.setActive(instance.store, { sheet: 0, row: 9, col: 9 });
    expect(open().items.map((item) => item[3])).toEqual(['true', 'false']);
  });

  it('falls back to the first sheet in tab order that has members, sheet-qualified', async () => {
    const { cycles, open } = await setup();
    cycles.set(1, [{ row: 3, col: 2 }]);
    expect(open().items).toEqual([["'My Sheet'!$C$4", '1:3:2', 'menuitemradio', 'true']]);
  });

  it('is off with no cycle, under iterative calculation, and without the engine method', async () => {
    const { instance, cycles, list, open } = await setup();
    expect(open()).toEqual({ disabled: 'true', items: [] });
    cycles.set(0, [{ row: 0, col: 0 }]);
    vi.spyOn(instance.workbook, 'getIterative').mockReturnValue({
      enabled: true,
      maxIterations: 100,
      maxChange: 0.001,
    });
    list.mockClear();
    expect(open()).toEqual({ disabled: 'true', items: [] });
    expect(list).not.toHaveBeenCalled();
  });

  it('selects the entry, switching sheet when it lives elsewhere', async () => {
    const { instance, defaults } = await setup();
    defaults.selectCircularReference('1:3:2');
    const state = instance.store.getState();
    expect([state.data.sheetIndex, state.selection.active]).toEqual([
      1,
      { sheet: 1, row: 3, col: 2 },
    ]);
  });
});

describe('WorkbookHandle engine method detection', () => {
  it('reports the methods only when the engine objects carry them', async () => {
    const handle = await WorkbookHandle.createDefault({ preferStub: true });
    const engine = (handle as unknown as { wb: { sheet(index: number): object } }).wb;
    const sheetProto = Object.getPrototypeOf(engine.sheet(0)) as Record<string, unknown>;
    const bookProto = Object.getPrototypeOf(engine) as Record<string, unknown>;
    try {
      expect([
        handle.circularReferencesSupported(),
        handle.staleValuesSupported(),
        handle.compatibilityVersion(),
      ]).toEqual([false, false, null]);
      sheetProto.listCircularReferences = () => [];
      sheetProto.isCellStale = () => false;
      bookProto.getCompatibilityVersion = () => 3;
      expect([
        handle.circularReferencesSupported(),
        handle.staleValuesSupported(),
        handle.compatibilityVersion(),
      ]).toEqual([true, true, 3]);
    } finally {
      delete sheetProto.listCircularReferences;
      delete sheetProto.isCellStale;
      delete bookProto.getCompatibilityVersion;
      handle.dispose();
    }
  });
});
