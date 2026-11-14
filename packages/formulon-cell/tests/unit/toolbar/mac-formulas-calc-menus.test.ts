import { afterEach, describe, expect, it, vi } from 'vitest';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createDefaultDynamicDropdownsCtx } from '../../../src/mount/dynamic-dropdowns-defaults.js';
import { mutators } from '../../../src/store/store.js';
import type { ApplyRibbonCommandDeps } from '../../../src/toolbar/ribbon/apply-ribbon-command.js';
import { DYNAMIC_RIBBON_DROPDOWN_MENU_REFRESHERS } from '../../../src/toolbar/ribbon/dynamic-dropdowns.js';
import { dispatchMacRibbonCommand } from '../../../src/toolbar/ribbon/mac/dispatch.js';
import { macRibbonLabelForCommand } from '../../../src/toolbar/ribbon/mac/labels.js';
import {
  MAC_RIBBON_MENU_ITEMS,
  macMenuLeaves,
} from '../../../src/toolbar/ribbon/mac/menu-manifest.js';
import { createMacRibbonMenuFactory } from '../../../src/toolbar/ribbon/mac/menus.js';
import { projectMacRibbonState } from '../../../src/toolbar/ribbon/mac/state-projection.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/mount.js';

const VERSION_IDS = [
  'mac.formulas.calc.compatVersion.1',
  'mac.formulas.calc.compatVersion.2',
  'mac.formulas.calc.compatVersion.3',
] as const;

describe('Mac Formulas > Calculation Options and Error Checking', () => {
  let sheet: MountedStubSheet | undefined;

  afterEach(() => {
    vi.restoreAllMocks();
    sheet?.dispose();
    sheet = undefined;
  });

  const mount = async () => {
    sheet = await mountStubSheet({
      workbook: await WorkbookHandle.createDefault({ preferStub: true }),
      locale: 'en',
    });
    const host = document.createElement('div');
    const factory = createMacRibbonMenuFactory('en');
    for (const root of ['mac.formulas.calcOptions', 'mac.formulas.errorCheck']) {
      const menu = factory(root);
      if (menu) host.appendChild(menu);
    }
    const only = (selector: string): HTMLButtonElement => {
      const found = host.querySelector<HTMLButtonElement>(selector);
      if (!found) throw new Error(`${selector} was not rendered`);
      return found;
    };
    const instance = sheet.instance;
    // What opening Error Checking runs: the Mac menu id's refresher.
    const openErrorCheck = () => {
      const key = DYNAMIC_RIBBON_DROPDOWN_MENU_REFRESHERS['menu-mac-formulas-error-check'];
      const menu = host.querySelector<HTMLElement>('#menu-mac-formulas-error-check');
      if (!key || !menu) throw new Error('Error Checking has no open refresher');
      createDefaultDynamicDropdownsCtx(instance)[key]?.(menu);
    };
    return { instance, host, only, openErrorCheck };
  };

  it('lists the measured rows, with the two lists as submenus', () => {
    const calc = MAC_RIBBON_MENU_ITEMS['mac.formulas.calcOptions'] ?? [];
    expect(calc.map((item) => [item.id, item.check ?? null, !!item.submenu])).toEqual([
      ['mac.formulas.calc.auto', null, false],
      ['mac.formulas.calc.autoNoTable', null, false],
      ['mac.formulas.calc.manual', null, false],
      ['mac.formulas.calc.formatStale', 'checkbox', false],
      ['mac.formulas.calc.compatVersion', null, true],
    ]);
    expect(
      calc.at(-1)?.submenu?.map((item) => [item.id, item.check ?? null, !!item.disabled]),
    ).toEqual([
      [VERSION_IDS[0], 'radio', false],
      [VERSION_IDS[1], 'radio', false],
      [VERSION_IDS[2], 'radio', false],
      ['mac.formulas.calc.compatVersion.learnMore', null, true],
    ]);
    const errorCheck = MAC_RIBBON_MENU_ITEMS['mac.formulas.errorCheck'] ?? [];
    expect(errorCheck.map((item) => [item.id, item.submenu?.length ?? null])).toEqual([
      ['mac.formulas.errorCheck.run', null],
      ['mac.formulas.errorCheck.trace', null],
      ['mac.formulas.errorCheck.circular', 0],
    ]);
    expect(macMenuLeaves(calc).map((item) => item.id)).not.toContain(
      'mac.formulas.calc.compatVersion',
    );
  });

  it('composes the version labels from the per-locale templates', () => {
    expect(VERSION_IDS.map((id) => macRibbonLabelForCommand(id, 'en'))).toEqual([
      'Version 1',
      'Version 2 (Recommended)',
      'Version 3 (Unsupported)',
    ]);
    expect(VERSION_IDS.map((id) => macRibbonLabelForCommand(id, 'ja'))).toEqual([
      'バージョン 1',
      'バージョン 2 (推奨)',
      'バージョン 3 (サポートされていません)',
    ]);
    expect(macRibbonLabelForCommand(VERSION_IDS[1], 'ko')).toBe('버전 2(권장)');
    expect(macRibbonLabelForCommand('mac.formulas.errorCheck.circular', 'de')).toBe('Zirkelbezüge');
  });

  it('renders submenu rows in the format-submenu markup and checkable rows with aria-checked', async () => {
    const { host, only } = await mount();
    const trigger = only('[data-format-submenu="mac.formulas.calc.compatVersion"]');
    expect(trigger.dataset.ribbonCommand).toBeUndefined();
    expect(trigger.getAttribute('aria-haspopup')).toBe('menu');
    expect(trigger.getAttribute('aria-controls')).toBe(
      'menu-mac-formulas-calc-options-compat-version',
    );
    const panel = host.querySelector<HTMLElement>('#menu-mac-formulas-calc-options-compat-version');
    expect(panel?.dataset.formatPanel).toBe('mac.formulas.calc.compatVersion');
    expect(
      [...(panel?.querySelectorAll<HTMLElement>('[data-ribbon-command]') ?? [])].map((button) => [
        button.dataset.ribbonCommand,
        button.getAttribute('role'),
        button.getAttribute('aria-checked'),
      ]),
    ).toEqual([
      [VERSION_IDS[0], 'menuitemradio', 'false'],
      [VERSION_IDS[1], 'menuitemradio', 'false'],
      [VERSION_IDS[2], 'menuitemradio', 'false'],
      ['mac.formulas.calc.compatVersion.learnMore', 'menuitem', null],
    ]);
    const stale = only('[data-ribbon-command="mac.formulas.calc.formatStale"]');
    expect(stale.getAttribute('role')).toBe('menuitemcheckbox');
    expect(
      host.querySelector('#menu-mac-formulas-error-check-circular')?.getAttribute('role'),
    ).toBe('menu');
  });

  it('keeps every new row disabled while the engine lacks the methods', async () => {
    const { instance, host, only, openErrorCheck } = await mount();
    projectMacRibbonState(host, instance);
    openErrorCheck();
    expect(
      [
        '[data-format-submenu="mac.formulas.errorCheck.circular"]',
        '[data-format-submenu="mac.formulas.calc.compatVersion"]',
        '[data-ribbon-command="mac.formulas.calc.formatStale"]',
        ...VERSION_IDS.slice(0, 2).map((id) => `[data-ribbon-command="${id}"]`),
      ].map((selector) => only(selector).getAttribute('aria-disabled')),
    ).toEqual(['true', 'true', 'true', 'true', 'true']);
    expect(only(`[data-ribbon-command="${VERSION_IDS[2]}"]`).hidden).toBe(true);
  });

  it('enables the rows once the engine provides them and projects the ticks', async () => {
    const { instance, host, only, openErrorCheck } = await mount();
    vi.spyOn(instance.workbook, 'circularReferencesSupported').mockReturnValue(true);
    vi.spyOn(instance.workbook, 'circularReferences').mockReturnValue([{ row: 1, col: 1 }]);
    vi.spyOn(instance.workbook, 'staleValuesSupported').mockReturnValue(true);
    const version = vi.spyOn(instance.workbook, 'compatibilityVersion').mockReturnValue(2);
    const mode = vi.spyOn(instance.workbook, 'calcMode').mockReturnValue(1);
    const state = (id: string) => {
      const button = only(`[data-ribbon-command="${id}"]`);
      return [
        button.hidden,
        button.getAttribute('aria-disabled'),
        button.getAttribute('aria-checked'),
      ];
    };

    projectMacRibbonState(host, instance);
    openErrorCheck();
    expect(
      only('[data-format-submenu="mac.formulas.errorCheck.circular"]').getAttribute(
        'aria-disabled',
      ),
    ).toBe('false');
    const cell = only('#menu-mac-formulas-error-check-circular [data-circular-reference]');
    expect([
      cell.textContent,
      cell.classList.contains('fc-tb__menu-item--mac'),
      cell.getAttribute('aria-checked'),
    ]).toEqual(['$B$2', true, 'true']);
    expect(state('mac.formulas.calc.formatStale')).toEqual([false, 'false', 'true']);
    expect(VERSION_IDS.map(state)).toEqual([
      [false, 'false', 'false'],
      [false, 'false', 'true'],
      [true, 'false', 'false'],
    ]);

    version.mockReturnValue(3);
    projectMacRibbonState(host, instance);
    expect(state(VERSION_IDS[2])).toEqual([false, 'false', 'true']);

    // Under Automatic, Format Stale Values stays off without a reason.
    mode.mockReturnValue(0);
    mutators.setFormatStaleValues(instance.store, false);
    projectMacRibbonState(host, instance);
    const stale = only('[data-ribbon-command="mac.formulas.calc.formatStale"]');
    expect(state('mac.formulas.calc.formatStale')).toEqual([false, 'true', 'false']);
    expect(stale.dataset.ribbonDisabledReason).toBeUndefined();
  });

  it('toggles the Format Stale Values preference only when the engine reports stale values', async () => {
    const { instance } = await mount();
    const deps = {
      inst: instance,
      runtime: { projectFormatToolbar: () => undefined, focusSheet: () => undefined },
    } as unknown as ApplyRibbonCommandDeps;
    expect(instance.store.getState().ui.formatStaleValues).toBe(true);
    dispatchMacRibbonCommand('mac.formulas.calc.formatStale', deps, () => false);
    expect(instance.store.getState().ui.formatStaleValues).toBe(true);
    vi.spyOn(instance.workbook, 'staleValuesSupported').mockReturnValue(true);
    dispatchMacRibbonCommand('mac.formulas.calc.formatStale', deps, () => false);
    expect(instance.store.getState().ui.formatStaleValues).toBe(false);
  });
});
