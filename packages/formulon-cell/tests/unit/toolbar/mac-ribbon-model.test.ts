import { describe, expect, it, vi } from 'vitest';
import { COMMAND_IDS } from '../../../src/api/spreadsheet/command-ids.js';
import {
  FUNCTION_CATEGORY_NAMES,
  supportedFunctionNames,
} from '../../../src/commands/function-categories.js';
import { recordRecentFunction } from '../../../src/commands/function-history.js';
import { History } from '../../../src/commands/history.js';
import {
  InteractionController,
  registerInteractionController,
} from '../../../src/commands/interaction-controller.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createI18nController } from '../../../src/i18n/controller.js';
import { dictionaries } from '../../../src/i18n/strings.js';
import type { SpreadsheetInstance } from '../../../src/mount/types.js';
import { Spreadsheet } from '../../../src/mount.js';
import { createSpreadsheetStore, defaultPageSetup, mutators } from '../../../src/store/store.js';
import { ribbonActivationForCommand } from '../../../src/toolbar/ribbon/activation.js';
import type { ApplyRibbonCommandDeps } from '../../../src/toolbar/ribbon/apply-ribbon-command.js';
import {
  dispatchMacRibbonCommand,
  isMacRibbonCommandSupported,
} from '../../../src/toolbar/ribbon/mac/dispatch.js';
import {
  createMacRibbonMenuFactory,
  projectMacFunctionCategoryMenus,
  projectMacRecentMenu,
} from '../../../src/toolbar/ribbon/mac/menus.js';
import {
  MAC_RIBBON_COMMAND_IDS,
  MAC_RIBBON_DISABLED_COMMAND_IDS,
  MAC_RIBBON_MENU_COMMAND_SET,
  MAC_RIBBON_MENU_ITEMS,
  macMenuLeaves,
  macRibbonLabelForCommand,
  macRibbonMenuIdForCommand,
  projectMacRibbonState,
} from '../../../src/toolbar/ribbon/mac/model.js';
import type { RibbonRenderHelpers } from '../../../src/toolbar/ribbon/render-ribbon.js';
import { buildRibbonModel, EXCEL365_MAC_RIBBON_TABS } from '../../../src/toolbar/ribbon-model.js';
import { mountStubSheet } from '../../test-utils/mount.js';

const findCommand = (tabId: string, commandId: string) =>
  buildRibbonModel('en', { profile: 'excel365Mac' })
    .find((tab) => tab.id === tabId)
    ?.groups.flatMap((group) => group.commands)
    .find((candidate) => candidate.id === commandId);

const projectedButton = (host: HTMLElement, id: string): HTMLButtonElement => {
  const button = document.createElement('button');
  button.dataset.ribbonCommand = id;
  button.setAttribute('aria-label', macRibbonLabelForCommand(id, 'en'));
  button.title = macRibbonLabelForCommand(id, 'en');
  host.append(button);
  return button;
};

const testInstance = () => {
  const store = createSpreadsheetStore();
  const i18n = createI18nController({ locale: 'en' });
  const instance = {
    store,
    i18n,
    workbook: {
      calcMode: () => 'auto',
      getExternalLinks: () => [],
      getPivotTables: () => [],
      definedNames: function* () {},
    },
    isSheetProtected: () => false,
  } as unknown as SpreadsheetInstance;
  return { store, i18n, instance };
};

const mountedRibbonHelpers = (): RibbonRenderHelpers => ({
  createSelect: () => document.createElement('div'),
  createColor: () => document.createElement('div'),
  createIcon: () => null,
  makeSvg: (viewBox, pathData, className) => {
    const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
    svg.setAttribute('viewBox', viewBox);
    svg.setAttribute('d', pathData);
    svg.setAttribute('class', className);
    return svg;
  },
  chevronPath: 'M0 0',
});

const mountedFunctionLeaf = (host: HTMLElement, name: string): HTMLButtonElement | null =>
  host.querySelector<HTMLButtonElement>(`[data-ribbon-command="mac.function.${name}"]`);

describe('Mac Excel ribbon profile', () => {
  it('keeps Bar and Area inside native chart galleries and exposes supported shapes visually', () => {
    expect(findCommand('insert', 'mac.insert.chartBar')).toBeUndefined();
    expect(findCommand('insert', 'mac.insert.chartArea')).toBeUndefined();
    for (const root of [
      'mac.insert.chartColumn',
      'mac.insert.chartLine',
      'mac.insert.chartPie',
      'mac.insert.chartScatter',
    ]) {
      const gallery = createMacRibbonMenuFactory('en')(root);
      expect(gallery?.classList.contains('fc-tb__menu--visual')).toBe(true);
      const tiles = Array.from(gallery?.querySelectorAll('.fc-tb__visual-tile') ?? []);
      expect(tiles.length).toBeGreaterThan(0);
      for (const tile of tiles) {
        expect(tile.getAttribute('role')).toBe('menuitem');
        expect(tile.querySelector('svg')).not.toBeNull();
      }
    }
    const shapes = createMacRibbonMenuFactory('en')('mac.insert.shapes');
    expect(shapes?.querySelectorAll('.fc-tb__visual-tile').length).toBe(7);
    expect(MAC_RIBBON_MENU_ITEMS['mac.formulas.calcOptions']?.map((item) => item.id)).toEqual([
      'mac.formulas.calc.auto',
      'mac.formulas.calc.autoNoTable',
      'mac.formulas.calc.manual',
      'separator',
      'mac.formulas.calc.formatStale',
      'separator',
      'mac.formulas.calc.compatVersion',
    ]);
    expect(MAC_RIBBON_MENU_ITEMS['mac.formulas.autoSum']?.map((item) => item.id)).toEqual([
      'mac.autosum.SUM',
      'mac.autosum.AVERAGE',
      'mac.autosum.COUNT',
      'mac.autosum.MAX',
      'mac.autosum.MIN',
      'mac.formulas.category.all',
    ]);
  });

  it('projects only actual recent functions in MRU order and preserves unchanged menu focus', () => {
    const store = createSpreadsheetStore();
    const host = document.createElement('div');
    const menu = createMacRibbonMenuFactory('en')('mac.formulas.recent');
    if (!menu) throw new Error('Recent menu missing');
    menu.hidden = false;
    host.append(menu);
    document.body.append(host);
    projectMacRecentMenu(host, store, 'en');
    expect(menu.querySelectorAll('[data-ribbon-command]').length).toBe(1);
    recordRecentFunction(store, 'SUM');
    recordRecentFunction(store, 'COUNTIF');
    projectMacRecentMenu(host, store, 'en');
    expect(
      Array.from(menu.querySelectorAll<HTMLElement>('[data-ribbon-command]')).map(
        (button) => button.dataset.ribbonCommand,
      ),
    ).toEqual(['mac.function.COUNTIF', 'mac.function.SUM', 'mac.formulas.category.recent']);
    const first = menu.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="mac.function.COUNTIF"]',
    );
    first?.focus();
    projectMacRecentMenu(host, store, 'en');
    expect(menu.querySelector('[data-ribbon-command="mac.function.COUNTIF"]')).toBe(first);
    expect(document.activeElement).toBe(first);
    recordRecentFunction(store, 'AVERAGE');
    projectMacRecentMenu(host, store, 'en');
    expect((document.activeElement as HTMLElement).dataset.ribbonCommand).toBe(
      'mac.function.COUNTIF',
    );
    host.remove();
  });

  it('projects explicit function families from the live catalog and filters dynamic Recent leaves', () => {
    const store = createSpreadsheetStore();
    const host = document.createElement('div');
    for (const id of ['mac.formulas.math', 'mac.formulas.recent']) {
      const menu = createMacRibbonMenuFactory('en')(id);
      if (!menu) throw new Error(`missing ${id}`);
      menu.hidden = false;
      host.append(menu);
    }
    document.body.append(host);
    recordRecentFunction(store, 'ACOS', new Set(['ACOS']));
    projectMacFunctionCategoryMenus(host, store, 'en', new Set(['SUM', 'ACOS']));
    const math = host.querySelector<HTMLElement>('#menu-mac-formulas-math');
    expect(
      Array.from(math?.querySelectorAll<HTMLElement>('[data-ribbon-command]') ?? []).map(
        (button) => button.dataset.ribbonCommand,
      ),
    ).toEqual(['mac.function.ACOS', 'mac.function.SUM']);
    const acos = math?.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="mac.function.ACOS"]',
    );
    expect(acos?.getAttribute('role')).toBe('menuitem');
    expect(acos?.title).toBe('ACOS');
    expect(acos?.getAttribute('aria-label')).toBe('ACOS');
    math?.querySelector<HTMLButtonElement>('[data-ribbon-command="mac.function.SUM"]')?.focus();
    projectMacFunctionCategoryMenus(host, store, 'en', new Set(['SUM']));
    expect(math?.querySelector('[data-ribbon-command="mac.function.ACOS"]')).toBeNull();
    const replacement = math?.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="mac.function.SUM"]',
    );
    expect(replacement).not.toBeNull();
    expect(document.activeElement).toBe(replacement);
    projectMacRecentMenu(host, store, 'en', new Set(['ACOS']));
    expect(host.querySelector('[data-ribbon-command="mac.function.ACOS"]')).not.toBeNull();
    projectMacRecentMenu(host, store, 'en', new Set(['SUM']));
    expect(host.querySelector('[data-ribbon-command="mac.function.ACOS"]')).toBeNull();
    host.remove();
  });

  it('keeps class-3 availability stable across repeated projection and reader replacement', () => {
    const store = createSpreadsheetStore();
    const host = document.createElement('div');
    const menu = createMacRibbonMenuFactory('en')('mac.formulas.more');
    if (!menu) throw new Error('More Functions menu missing');
    host.append(menu);
    document.body.append(host);
    const metadataCalls: string[] = [];
    const reader = {
      functionMetadata: (name: string) => {
        metadataCalls.push(name);
        return {
          name,
          minArity: 1,
          maxArity: 1,
          ...(name === 'CUBEVALUE' ? { availability: 3 } : { availability: 2 }),
        };
      },
    };
    const liveNames = new Set(['CUBEVALUE', 'CUBESET', 'ACOS']);
    projectMacFunctionCategoryMenus(host, store, 'en', liveNames, {
      reader,
      unavailableReason: 'Unavailable in this calculation engine.',
    });
    const cube = host.querySelector<HTMLElement>('[data-function-category-panel="cube"]');
    const cubeValue = cube?.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="mac.function.CUBEVALUE"]',
    );
    const cubeSet = cube?.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="mac.function.CUBESET"]',
    );
    expect(cubeValue).not.toBeNull();
    expect(cubeValue?.disabled).toBe(true);
    expect(cubeValue?.dataset.functionUnavailable).toBe('true');
    expect(cubeValue?.getAttribute('aria-description')).toBe(
      'Unavailable in this calculation engine.',
    );
    expect(cubeValue?.title).toContain('Unavailable in this calculation engine.');
    expect(cubeSet?.disabled).toBe(false);

    const firstCallCount = metadataCalls.length;
    projectMacFunctionCategoryMenus(host, store, 'en', liveNames, {
      reader,
      unavailableReason: 'Unavailable in this calculation engine.',
    });
    expect(metadataCalls).toHaveLength(firstCallCount);
    expect(cubeValue?.getAttribute('aria-description')).toBe(
      'Unavailable in this calculation engine.',
    );

    const secondReader = {
      functionMetadata: (name: string) => ({
        name,
        minArity: 1,
        maxArity: 1,
        availability: name === 'CUBEVALUE' ? 2 : 3,
      }),
    };
    projectMacFunctionCategoryMenus(host, store, 'en', liveNames, {
      reader: secondReader,
      unavailableReason: 'Unavailable in this calculation engine.',
    });
    expect(
      cube?.querySelector<HTMLButtonElement>('[data-ribbon-command="mac.function.CUBEVALUE"]')
        ?.disabled,
    ).toBe(false);
    expect(
      cube?.querySelector<HTMLButtonElement>('[data-ribbon-command="mac.function.CUBESET"]')
        ?.disabled,
    ).toBe(true);
    host.remove();
  });

  it('renders More Functions as the native six-category submenu tree', () => {
    const menu = createMacRibbonMenuFactory('en')('mac.formulas.more');
    if (!menu) throw new Error('More Functions menu missing');
    const triggers = Array.from(menu.children).filter(
      (child): child is HTMLButtonElement =>
        child instanceof HTMLButtonElement && !!child.dataset.functionCategorySubmenu,
    );
    expect(triggers.map((button) => button.dataset.functionCategorySubmenu)).toEqual([
      'statistical',
      'engineering',
      'cube',
      'information',
      'compatibility',
      'web',
    ]);
    expect(triggers.map((button) => button.textContent?.replace('▶', '').trim())).toEqual([
      'Statistical',
      'Engineering',
      'Cube',
      'Information',
      'Compatibility',
      'Web',
    ]);
    const jaMenu = createMacRibbonMenuFactory('ja')('mac.formulas.more');
    expect(
      Array.from(
        jaMenu?.querySelectorAll<HTMLElement>('[data-function-category-submenu]') ?? [],
      ).map((button) => button.textContent?.replace('▶', '').trim()),
    ).toEqual(['統計', 'エンジニアリング', 'キューブ', '情報', '互換性', 'Web']);
    expect(triggers.every((button) => !button.dataset.ribbonCommand)).toBe(true);
    expect(triggers.every((button) => button.getAttribute('aria-haspopup') === 'menu')).toBe(true);
    expect(triggers.every((button) => button.getAttribute('aria-expanded') === 'false')).toBe(true);
    for (const trigger of triggers) {
      const panelId = trigger.getAttribute('aria-controls');
      expect(panelId).toBe(`menu-mac-formulas-more-${trigger.dataset.functionCategorySubmenu}`);
      const panel = menu.querySelector<HTMLElement>(`#${panelId}`);
      expect(panel?.dataset.functionCategoryPanel).toBe(trigger.dataset.functionCategorySubmenu);
      expect(panel?.getAttribute('role')).toBe('menu');
    }
    expect(menu.querySelectorAll('[data-function-category-panel]')).toHaveLength(6);
    const statistical = menu.querySelector<HTMLElement>(
      '[data-function-category-panel="statistical"]',
    );
    if (!statistical) throw new Error('Statistical More panel missing');
    expect(
      statistical?.querySelector('[data-ribbon-command="mac.function.COUNTIF"]'),
    ).not.toBeNull();
    expect(statistical?.querySelector('[data-ribbon-command="mac.function.ACOS"]')).toBeNull();

    const host = document.createElement('div');
    host.append(menu);
    document.body.append(host);
    projectMacFunctionCategoryMenus(
      host,
      createSpreadsheetStore(),
      'en',
      new Set(['COUNTIF', 'ACOS']),
    );
    expect(
      statistical?.querySelector('[data-ribbon-command="mac.function.COUNTIF"]'),
    ).not.toBeNull();
    expect(
      host.querySelector('[data-function-category-panel="engineering"] [data-ribbon-command]'),
    ).toBeNull();
    expect(
      host.querySelector('[data-function-category-panel="math"] [data-ribbon-command]'),
    ).toBeNull();
    menu.hidden = false;
    statistical.hidden = false;
    const countIf = statistical.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="mac.function.COUNTIF"]',
    );
    countIf?.focus();
    projectMacFunctionCategoryMenus(host, createSpreadsheetStore(), 'en', new Set(['AVERAGE']));
    expect(document.activeElement).toBe(
      statistical.querySelector('[data-ribbon-command="mac.function.AVERAGE"]'),
    );
    projectMacFunctionCategoryMenus(host, createSpreadsheetStore(), 'en', new Set());
    expect(menu.querySelectorAll('[data-function-category-panel]')).toHaveLength(6);
    expect(
      menu.querySelectorAll('[data-function-category-panel] [data-ribbon-command]'),
    ).toHaveLength(0);
    host.remove();
  });

  it('projects policy-denied live leaves as disabled and policy-allowed leaves as enabled', async () => {
    const store = createSpreadsheetStore();
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    const controller = new InteractionController({
      store,
      getWb: () => workbook,
      history: new History(),
    });
    const unregister = registerInteractionController(store, controller);
    const host = document.createElement('div');
    const menu = createMacRibbonMenuFactory('en')('mac.formulas.math');
    if (!menu) throw new Error('Math menu missing');
    host.append(menu);
    document.body.append(host);
    try {
      const disabledOf = (name: string): string | null | undefined =>
        menu
          .querySelector<HTMLButtonElement>(`[data-ribbon-command="mac.function.${name}"]`)
          ?.getAttribute('aria-disabled');
      controller.setPolicy({
        editable: [{ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }],
        operations: { formulaEdit: true },
        defaultOperation: 'deny',
        selection: true,
      });
      projectMacFunctionCategoryMenus(host, store, 'en', new Set(['ACOS', 'SUM']));
      expect(disabledOf('ACOS')).not.toBe('true');
      expect(disabledOf('SUM')).not.toBe('true');
      controller.setPolicy({
        editable: [{ sheet: 0, r0: 5, c0: 5, r1: 5, c1: 5 }],
        operations: { formulaEdit: true },
        defaultOperation: 'deny',
        selection: true,
      });
      projectMacFunctionCategoryMenus(host, store, 'en', new Set(['ACOS', 'SUM']));
      expect(disabledOf('ACOS')).toBe('true');
      expect(disabledOf('SUM')).toBe('true');
    } finally {
      unregister();
      controller.dispose();
      workbook.dispose();
      host.remove();
    }
  });

  it('reprojects mounted policy state without re-reading native availability metadata', async () => {
    const workbook = await WorkbookHandle.createDefault();
    expect(workbook.isStub).toBe(false);

    // Keep the real native handle and metadata implementation, while the
    // two-name response makes the mounted projection deterministic. The
    // availability assertions below are the raw native classes, not fixture
    // overlays or a fake workbook/controller.
    const nativeFunctionMetadata = workbook.functionMetadata.bind(workbook);
    const metadata = vi.spyOn(workbook, 'functionMetadata');
    expect(workbook.functionMetadata('ACOS')?.availability).toBe(0);
    expect(workbook.functionMetadata('CUBEVALUE')?.availability).toBe(3);
    metadata.mockImplementation((name, locale = 0) => nativeFunctionMetadata(name, locale));
    vi.spyOn(workbook, 'functionNames').mockReturnValue(['ACOS', 'CUBEVALUE']);

    const sheet = await mountStubSheet({ locale: 'en', workbook });
    const toolbarHost = document.createElement('div');
    document.body.append(toolbarHost);
    const toolbar = Spreadsheet.mountToolbar(toolbarHost, sheet.instance, {
      platform: 'mac',
      activeTab: 'formulas',
      helpers: mountedRibbonHelpers(),
    });
    // ACOS is supplied by the live catalog rather than the legacy static
    // signature table. Mark the deterministic fixture as a previously
    // successful insertion so restricted-command classification can exercise
    // the same dynamic-function path used by the mounted picker.
    recordRecentFunction(sheet.instance.store, 'ACOS', new Set(['ACOS', 'CUBEVALUE']));

    const leaf = (name: string): HTMLButtonElement | null => mountedFunctionLeaf(toolbarHost, name);
    const setAcosPolicy = (allowAcos: boolean): void => {
      sheet.instance.commands.setPolicy({
        editable: () => true,
        operations: { formulaEdit: true },
        defaultOperation: 'deny',
        restrict: ({ commandId }) => allowAcos || commandId !== 'mac.function.ACOS',
      });
    };

    try {
      const initialAcos = leaf('ACOS');
      const initialCube = leaf('CUBEVALUE');
      expect(initialAcos).not.toBeNull();
      expect(initialAcos?.getAttribute('aria-disabled')).not.toBe('true');
      expect(initialCube?.dataset.functionUnavailable).toBe('true');
      expect(initialCube?.getAttribute('aria-disabled')).toBe('true');
      const intrinsicDescription = initialCube?.getAttribute('aria-description');
      const intrinsicTitle = initialCube?.title;
      expect(intrinsicDescription).toBeTruthy();
      expect(intrinsicTitle).toContain(intrinsicDescription ?? '');

      const metadataCallsAfterMount = metadata.mock.calls.length;
      expect(metadataCallsAfterMount).toBeGreaterThan(0);

      setAcosPolicy(false);
      const deniedAcos = leaf('ACOS');
      const deniedCube = leaf('CUBEVALUE');
      expect(deniedAcos?.getAttribute('aria-disabled')).toBe('true');
      expect(deniedCube?.dataset.functionUnavailable).toBe('true');
      expect(deniedCube?.getAttribute('aria-description')).toBe(intrinsicDescription);
      expect(deniedCube?.title).toBe(intrinsicTitle);
      expect(metadata.mock.calls.length).toBe(metadataCallsAfterMount);

      setAcosPolicy(true);
      const allowedAcos = leaf('ACOS');
      const allowedCube = leaf('CUBEVALUE');
      expect(allowedAcos?.getAttribute('aria-disabled')).not.toBe('true');
      expect(allowedCube?.dataset.functionUnavailable).toBe('true');
      expect(allowedCube?.getAttribute('aria-description')).toBe(intrinsicDescription);
      expect(allowedCube?.title).toBe(intrinsicTitle);
      expect(metadata.mock.calls.length).toBe(metadataCallsAfterMount);

      sheet.instance.commands.setPolicy(undefined);
      const clearedAcos = leaf('ACOS');
      const clearedCube = leaf('CUBEVALUE');
      expect(clearedAcos?.getAttribute('aria-disabled')).not.toBe('true');
      expect(clearedCube?.dataset.functionUnavailable).toBe('true');
      expect(clearedCube?.getAttribute('aria-description')).toBe(intrinsicDescription);
      expect(clearedCube?.title).toBe(intrinsicTitle);
      expect(metadata.mock.calls.length).toBe(metadataCallsAfterMount);
    } finally {
      toolbar.dispose();
      toolbarHost.remove();
      sheet.dispose();
    }
  });

  it('exposes local workbook links and explains local gaps without requiring Microsoft 365', () => {
    expect(findCommand('data', 'mac.data.workbookLinks')?.disabled).not.toBe(true);
    for (const [tab, id] of [
      ['insert', 'mac.insert.timeline'],
      ['data', 'mac.data.powerQuery'],
      ['review', 'mac.review.translate'],
    ]) {
      const item = findCommand(tab ?? '', id ?? '');
      expect(item?.disabled).toBe(true);
      expect(item?.disabledReason).toContain('not implemented');
      expect(item?.disabledReason).not.toContain('Microsoft 365');
    }
  });

  it('explains unavailable local menu features in Japanese', () => {
    const menu = createMacRibbonMenuFactory('ja')('mac.data.whatIf');
    expect(
      Array.from(menu?.querySelectorAll<HTMLElement>('[data-ribbon-command]') ?? []).map(
        (item) => item.dataset.ribbonCommand,
      ),
    ).toEqual(['mac.data.scenarioManager', 'mac.data.goalSeek', 'mac.data.dataTable']);
    const scenario = menu?.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="mac.data.scenarioManager"]',
    );
    expect(scenario?.getAttribute('aria-disabled')).toBe('true');
    expect(scenario?.title).toContain('この機能は未実装です。');
    expect(scenario?.title).not.toContain('Microsoft 365');
    expect(scenario?.textContent).toBe('シナリオの登録と管理...');
    const shapes = createMacRibbonMenuFactory('ja')('mac.insert.shapes');
    const callout = shapes?.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="mac.insert.shape.callouts.0"]',
    );
    expect(callout?.textContent).toBe('四角形吹き出し');
    expect(callout?.getAttribute('aria-disabled')).toBe('true');
    expect(callout?.title).toContain('この機能は未実装です。');
    expect(
      Array.from(shapes?.querySelectorAll('.fc-tb__menu-heading') ?? []).map(
        (heading) => heading.textContent,
      ),
    ).toEqual([
      '基本図形',
      '線',
      '矢印',
      '吹き出し',
      '記号と特殊文字',
      'フローチャート',
      '星とリボン',
    ]);
  });

  it('renders the nine Office 365 Mac tabs in order without File or Help', () => {
    const model = buildRibbonModel('en', { profile: 'excel365Mac' });
    expect(model.map((tab) => tab.id)).toEqual(EXCEL365_MAC_RIBBON_TABS);
    expect(model.map((tab) => tab.label)).toEqual([
      'Home',
      'Insert',
      'Draw',
      'Page Layout',
      'Formulas',
      'Data',
      'Review',
      'View',
      'Automate',
    ]);
    expect(model.some((tab) => tab.id === 'file' || tab.id === 'help')).toBe(false);
  });

  it('keeps Mac commands explicit and gives every menu root a stable menu id', () => {
    const model = buildRibbonModel('en', { profile: 'excel365Mac' });
    const commands = model
      .flatMap((tab) => tab.groups)
      .flatMap((group) => group.commands)
      .filter((command) => command.id.startsWith('mac.'));
    expect(commands.length).toBeGreaterThan(80);
    expect(new Set(commands.map((command) => command.id)).size).toBe(commands.length);
    for (const command of commands) {
      expect(MAC_RIBBON_COMMAND_IDS).toContain(command.id);
      if (MAC_RIBBON_DISABLED_COMMAND_IDS.has(command.id)) {
        expect(command.disabled).toBe(true);
        expect(command.disabledReason).toBeTruthy();
      }
      if (MAC_RIBBON_MENU_COMMAND_SET.has(command.id)) {
        // A disabled menu root still draws its chevron; the disabled flag greys it out.
        expect(ribbonActivationForCommand(command.id)).toEqual({
          kind: 'dropdown',
          menuId: macRibbonMenuIdForCommand(command.id),
        });
        // More Functions builds its menu from the catalog and the WordArt gallery lists no entries.
        if (command.id !== 'mac.formulas.more' && command.id !== 'mac.insert.wordArt') {
          expect(MAC_RIBBON_MENU_ITEMS[command.id]?.length).toBeGreaterThan(0);
        }
      }
    }
  });

  it('uses Mac-specific Home order and compact, direct command groups', () => {
    const model = buildRibbonModel('en', { profile: 'excel365Mac' });
    const home = model.find((tab) => tab.id === 'home');
    const genericHome = buildRibbonModel('en').find((tab) => tab.id === 'home');
    const macFont = home?.groups.find((group) => group.variant === 'font');
    const genericFont = genericHome?.groups.find((group) => group.variant === 'font');
    const fontIds = macFont?.commands.map((command) => command.id) ?? [];
    expect(genericFont?.commands.some((command) => command.id === 'strike')).toBe(true);
    expect(fontIds).not.toContain('strike');
    expect(fontIds.indexOf('showPhonetic')).toBeGreaterThan(fontIds.indexOf('fontColor'));
    expect(home?.groups.at(-1)?.commands.map((command) => command.id)).toEqual(['mac.home.addins']);
    expect(findCommand('home', 'mac.home.addins')?.disabled).toBeFalsy();

    const data = model.find((tab) => tab.id === 'data');
    expect(
      data?.groups.find((group) => group.title === 'Sort & Filter')?.commands.map((c) => c.id),
    ).toEqual([
      'mac.data.sortAsc',
      'mac.data.sortDesc',
      'mac.data.sortCustom',
      'mac.data.filter',
      'mac.data.clear',
      'mac.data.reapply',
      'mac.data.advancedFilter',
    ]);
    expect(MAC_RIBBON_MENU_COMMAND_SET.has('mac.data.sortFilter')).toBe(true);

    const page = model.find((tab) => tab.id === 'pageLayout');
    const sheetOptions = page?.groups.find((group) => group.title === 'Sheet Options');
    expect(sheetOptions?.variant).toBe('checks');
    expect(sheetOptions?.commands.map((command) => command.id)).toEqual([
      'mac.page.gridlinesHeading',
      'mac.page.showGridlines',
      'mac.page.printGridlines',
      'mac.page.headingsHeading',
      'mac.page.showHeadings',
      'mac.page.printHeadings',
    ]);
    expect(sheetOptions?.commands.every((command) => command.icon === undefined)).toBe(true);

    const view = model.find((tab) => tab.id === 'view');
    expect(
      view?.groups.find((group) => group.title === 'Sheet View')?.commands.map((c) => c.id),
    ).toEqual([
      'sheetViewSelect',
      'mac.view.sheetViewKeep',
      'mac.view.sheetViewExit',
      'mac.view.sheetViewNew',
      'mac.view.sheetViewOptions',
    ]);
    expect(
      view?.groups.find((group) => group.title === 'Workbook Views')?.commands.map((c) => c.id),
    ).toEqual([
      'mac.view.standard',
      'mac.view.pageBreakPreview',
      'mac.view.pageLayout',
      'mac.view.customViews',
    ]);
    expect(
      view?.groups.find((group) => group.title === 'Zoom')?.commands.map((c) => c.id),
    ).toContain('mac.view.zoom100');

    const automate = model.find((tab) => tab.id === 'automate');
    expect(
      automate?.groups.find((group) => group.title === 'Office Scripts')?.commands.map((c) => c.id),
    ).toEqual(['mac.automate.newScript', 'mac.automate.showScripts']);
    expect(
      automate?.groups.find((group) => group.title === 'Office Scripts Gallery')?.variant,
    ).toBe('gallery');
    expect(
      automate?.groups
        .find((group) => group.title === 'Office Scripts Gallery')
        ?.commands.map((c) => c.id),
    ).toEqual([
      'mac.automate.allRowsColumns',
      'mac.automate.freezeSelection',
      'mac.automate.makeSubtable',
      'mac.automate.removeHyperlinks',
      'mac.automate.countEmptyRows',
      'mac.automate.tableToJson',
      'mac.automate.newPivotTable',
    ]);
    expect(findCommand('automate', 'mac.automate.newPivotTable')?.disabled).not.toBe(true);
  });

  it('keeps Home alignment actions labeled and wide on Mac without changing their ids', () => {
    const locales = ['en', 'ja', 'de', 'fr', 'zh', 'ko', 'th'] as const;
    for (const locale of locales) {
      const macModel = buildRibbonModel(locale, { profile: 'excel365Mac' });
      const genericModel = buildRibbonModel(locale);
      const macAlignment = macModel
        .find((tab) => tab.id === 'home')
        ?.groups.find((group) => group.variant === 'alignment');
      const genericAlignment = genericModel
        .find((tab) => tab.id === 'home')
        ?.groups.find((group) => group.variant === 'alignment');
      const macById = new Map(macAlignment?.commands.map((command) => [command.id, command]));
      const genericById = new Map(
        genericAlignment?.commands.map((command) => [command.id, command]),
      );

      expect(macById.get('wrap')).toMatchObject({ kind: 'wide', layout: 'stacked' });
      expect(macById.get('merge')).toMatchObject({ kind: 'wide', layout: 'stacked' });
      expect(macById.get('merge')?.label).toBe(dictionaries[locale].ribbon.mergeAndCenter);
      expect(macById.get('merge')?.title).toBe(dictionaries[locale].ribbon.mergeAndCenter);
      expect(macById.get('wrap')?.label).toBe(
        locale === 'ja' ? '折り返して全体を表示する' : genericById.get('wrap')?.label,
      );
      expect(macById.get('wrap')?.title).toBe(
        locale === 'ja' ? '折り返して全体を表示する' : genericById.get('wrap')?.title,
      );
      expect(macAlignment?.commands.map((command) => command.id)).toEqual(
        genericAlignment?.commands.map((command) => command.id),
      );
      expect(genericById.get('wrap')).toMatchObject({ kind: 'button' });
      expect(genericById.get('merge')).toMatchObject({ kind: 'button' });
      expect(genericById.get('merge')?.label).toBe(dictionaries[locale].ribbon.mergeAndCenter);
      expect(genericById.get('merge')?.title).toBe(dictionaries[locale].ribbon.mergeAndCenter);
    }
  });

  it('lays out Mac Home clipboard, cells and editing like the desktop ribbon', () => {
    const locales = ['en', 'ja', 'de', 'fr', 'zh', 'ko', 'th'] as const;
    for (const locale of locales) {
      const homeOf = (profile?: 'excel365Mac') =>
        buildRibbonModel(locale, profile ? { profile } : {}).find((tab) => tab.id === 'home');
      const macHome = homeOf('excel365Mac');
      const genericHome = homeOf();
      const commandsOf = (home: typeof macHome) =>
        new Map(
          home?.groups.flatMap((group) => group.commands).map((command) => [command.id, command]),
        );
      const mac = commandsOf(macHome);
      const generic = commandsOf(genericHome);
      const macLabels = dictionaries[locale].mac.homeRibbon;

      expect(mac.get('paste')).toMatchObject({ kind: 'large', label: macLabels.paste });
      for (const [id, label] of [
        ['cut', macLabels.cut],
        ['copy', macLabels.copy],
        ['formatPainter', macLabels.formatPainter],
      ] as const) {
        expect(mac.get(id), `${locale} ${id}`).toMatchObject({ kind: 'wide', layout: 'stacked' });
        expect(mac.get(id)?.label).toBe(label);
        expect(generic.get(id)?.kind).toBe('button');
      }

      const macCells = macHome?.groups.find(
        (group) => group.title === dictionaries[locale].ribbon.cells,
      );
      expect(macCells?.variant).toBe('tiles');
      expect(genericHome?.groups.find((group) => group.variant === 'cells')).toBeDefined();
      for (const id of ['insertRows', 'deleteRows', 'formatCellsHome']) {
        expect(mac.get(id), `${locale} ${id}`).toMatchObject({ kind: 'wide', layout: undefined });
        expect(generic.get(id)?.layout).toBe('stacked');
      }

      for (const id of ['autosum', 'fillHome', 'clearFormat']) {
        expect(mac.get(id), `${locale} ${id}`).toMatchObject({ kind: 'wide', layout: 'stacked' });
        expect(mac.get(id)?.className ?? '').not.toContain('fc-tb__rb--mac-icon-only');
      }
      expect(mac.get('autosum')?.label).toBe(macLabels.autoSum);
      expect(mac.get('autosum')?.title).toBe(macLabels.autoSum);
      expect(mac.get('conditional')?.label).toBe(macLabels.conditional);
      expect(mac.get('cellStyles')?.label).toBe(macLabels.cellStyles);
      expect(generic.get('paste')?.label).toBe(dictionaries[locale].ribbon.paste);
      expect(generic.get('cellStyles')?.label).toBe(dictionaries[locale].ribbon.cellStyles);
    }
    expect(dictionaries.ja.mac.homeRibbon).toMatchObject({
      paste: 'ペースト',
      formatPainter: '書式',
      cellStyles: 'セルのスタイル',
      autoSum: 'オート SUM',
    });
    expect(dictionaries.en.mac.homeRibbon.formatPainter).toBe('Format');
  });

  it('projects Mac checked states and contextual eligibility without accumulating titles', () => {
    const { instance, i18n, store } = testInstance();
    const host = document.createElement('div');
    const ids = [
      'mac.page.showGridlines',
      'mac.page.printGridlines',
      'mac.page.showHeadings',
      'mac.page.printHeadings',
      'mac.data.clear',
      'mac.data.reapply',
      'mac.formulas.removeArrows.all',
      'mac.formulas.removeArrows.precedents',
      'mac.formulas.removeArrows.dependents',
      'mac.formulas.errorCheck.trace',
      'mac.data.validation.circleInvalid',
      'mac.data.validation.clearCircles',
      'mac.data.validation.clearRules',
    ];
    const buttons = new Map(ids.map((id) => [id, projectedButton(host, id)]));

    store.setState((state) => ({
      ...state,
      ui: { ...state.ui, showGridLines: false, showHeaders: true },
      pageSetup: {
        setupBySheet: new Map([
          [0, { ...defaultPageSetup(), showGridlines: true, showHeadings: false }],
        ]),
      },
    }));
    projectMacRibbonState(host, instance);
    expect(buttons.get('mac.page.showGridlines')?.getAttribute('aria-pressed')).toBe('false');
    expect(buttons.get('mac.page.printGridlines')?.getAttribute('aria-pressed')).toBe('true');
    expect(buttons.get('mac.page.showHeadings')?.getAttribute('aria-pressed')).toBe('true');
    expect(buttons.get('mac.page.printHeadings')?.getAttribute('aria-pressed')).toBe('false');

    const clear = buttons.get('mac.data.clear');
    projectMacRibbonState(host, instance);
    const disabledClearTitle = 'Clear\nThere is no active filter range.';
    expect(clear?.disabled).toBe(true);
    expect(clear?.title).toBe(disabledClearTitle);
    projectMacRibbonState(host, instance);
    expect(clear?.title).toBe(disabledClearTitle);
    expect(buttons.get('mac.data.reapply')?.disabled).toBe(true);
    expect(buttons.get('mac.formulas.removeArrows.all')?.disabled).toBe(true);
    expect(buttons.get('mac.formulas.errorCheck.trace')?.disabled).toBe(true);
    expect(buttons.get('mac.data.validation.circleInvalid')?.disabled).toBe(true);
    const range = { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 };
    store.setState((state) => ({
      ...state,
      ui: {
        ...state.ui,
        filterRanges: [range],
        filterCriteria: [{ range, byCol: 0, hiddenValues: ['hidden'] }],
      },
      traces: {
        items: [
          {
            kind: 'precedent',
            from: { sheet: 0, row: 0, col: 1 },
            to: { sheet: 0, row: 0, col: 0 },
          },
          {
            kind: 'dependent',
            from: { sheet: 0, row: 0, col: 0 },
            to: { sheet: 0, row: 0, col: 1 },
          },
        ],
      },
    }));
    mutators.setCell(
      store,
      { sheet: 0, row: 0, col: 0 },
      { kind: 'error', code: 1, text: '#DIV/0!' },
      '=1/0',
    );
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      {
        validation: { kind: 'whole', op: '>', a: 0 },
      },
    );
    store.setState((state) => ({
      ...state,
      errorIndicators: {
        ...state.errorIndicators,
        validationCircles: new Set(['0:0:0']),
      },
    }));
    projectMacRibbonState(host, instance);
    for (const id of ids.slice(4)) {
      expect(buttons.get(id)?.disabled, id).toBe(false);
      expect(buttons.get(id)?.getAttribute('aria-disabled'), id).toBe('false');
    }
    i18n.dispose();
  });

  it('disables commands whose target the document does not have', () => {
    const { instance, i18n } = testInstance();
    const host = document.createElement('div');
    const ids = [
      'mac.data.showDetail',
      'mac.data.hideDetail',
      'mac.data.workbookLinks',
      'mac.insert.slicer',
      'mac.formulas.useInFormula',
    ];
    const buttons = ids.map((id) => projectedButton(host, id));
    projectMacRibbonState(host, instance);
    expect(buttons.map((button) => button.disabled)).toEqual(ids.map(() => true));
    i18n.dispose();
  });

  it('uses Japanese labels while retaining the same action surface', () => {
    const model = buildRibbonModel('ja', { profile: 'excel365Mac' });
    const home = model.find((tab) => tab.id === 'home');
    const insert = model.find((tab) => tab.id === 'insert');
    expect(home?.label).toBe('ホーム');
    expect(insert?.label).toBe('挿入');
    expect(insert?.groups.some((group) => group.title === 'グラフ')).toBe(true);
    expect(
      model
        .flatMap((tab) => tab.groups)
        .flatMap((group) => group.commands)
        .filter((command) => command.id.startsWith('mac.'))
        .map((command) => command.id)
        .sort(),
    ).toEqual(
      buildRibbonModel('en', { profile: 'excel365Mac' })
        .flatMap((tab) => tab.groups)
        .flatMap((group) => group.commands)
        .filter((command) => command.id.startsWith('mac.'))
        .map((command) => command.id)
        .sort(),
    );
  });

  it('keeps every static Mac menu leaf in the command and activation catalog', () => {
    const unsupportedMenuLeaves: string[] = [];
    for (const [root, items] of Object.entries(MAC_RIBBON_MENU_ITEMS)) {
      expect(MAC_RIBBON_MENU_COMMAND_SET.has(root)).toBe(true);
      expect(macRibbonMenuIdForCommand(root)).toMatch(/^menu-mac-/);
      for (const item of macMenuLeaves(items).filter((candidate) => !candidate.heading)) {
        expect(MAC_RIBBON_COMMAND_IDS).toContain(item.id);
        const activation = ribbonActivationForCommand(item.id);
        const expectedKind = item.disabled
          ? 'disabled'
          : MAC_RIBBON_MENU_COMMAND_SET.has(item.id)
            ? 'dropdown'
            : 'primaryAction';
        expect(activation.kind).toBe(expectedKind);
        if (!item.disabled && !isMacRibbonCommandSupported(item.id)) {
          unsupportedMenuLeaves.push(item.id);
        }
      }
    }
    expect(unsupportedMenuLeaves).toEqual([]);
  });

  it('builds complete catalog-backed leaves and keeps Recently Used backed by session history', () => {
    const categories = ['financial', 'logical', 'text', 'datetime', 'lookup', 'math'] as const;
    for (const category of categories) {
      expect(
        MAC_RIBBON_MENU_ITEMS[
          `mac.formulas.${category === 'datetime' ? 'dateTime' : category}`
        ]?.map((item) => item.id),
      ).toEqual(supportedFunctionNames(category).map((name) => `mac.function.${name}`));
    }

    expect(MAC_RIBBON_MENU_COMMAND_SET.has('mac.formulas.recent')).toBe(true);
    expect(MAC_RIBBON_MENU_ITEMS['mac.formulas.recent']?.map((item) => item.id)).toEqual([
      'mac.formulas.category.recent',
    ]);
    expect(isMacRibbonCommandSupported('mac.formulas.recent')).toBe(true);
    expect(MAC_RIBBON_MENU_ITEMS['mac.formulas.more']).toEqual([]);
    expect(FUNCTION_CATEGORY_NAMES.statistical).toContain('COUNT');
    expect(FUNCTION_CATEGORY_NAMES.statistical).toContain('COUNTIF');
    expect(FUNCTION_CATEGORY_NAMES.math).not.toContain('COUNT');
    expect(FUNCTION_CATEGORY_NAMES.math).not.toContain('COUNTIF');
  });

  it('keeps backed Mac controls wired to the generic select dispatcher', () => {
    const model = buildRibbonModel('en', { profile: 'excel365Mac' });
    const command = (id: string) =>
      model
        .flatMap((tab) => tab.groups)
        .flatMap((group) => group.commands)
        .find((candidate) => candidate.id === id);
    expect(command('scaleWidth')).toMatchObject({ kind: 'select' });
    expect(command('scaleHeight')).toMatchObject({ kind: 'select' });
    expect(command('sheetViewSelect')).toMatchObject({ kind: 'select' });
    expect(command('mac.view.sheetViewNew')).toMatchObject({ disabled: undefined });
    expect(command('mac.view.sheetViewOptions')).toMatchObject({ disabled: undefined });
    expect(command('mac.data.showDetail')).toMatchObject({ disabled: undefined });
    expect(command('mac.data.hideDetail')).toMatchObject({ disabled: undefined });
  });

  it('distinguishes local gaps from cloud service requirements', () => {
    const model = buildRibbonModel('en', { profile: 'excel365Mac' });
    const command = (id: string) =>
      model
        .flatMap((tab) => tab.groups)
        .flatMap((group) => group.commands)
        .find((candidate) => candidate.id === id);
    for (const id of [
      'mac.insert.icons',
      'mac.insert.threeDModel',
      'mac.insert.smartArt',
      'mac.insert.checkBox',
      'mac.insert.pivotChart',
      'mac.insert.wordArt',
      'mac.insert.object',
      'mac.insert.equation',
      'mac.data.analysis',
      'mac.view.customViews',
      'mac.view.macros',
    ]) {
      expect(command(id)?.disabledReason).toBe('This feature is not implemented yet.');
    }
    expect(command('mac.insert.forms')?.disabledReason).toBe(
      'Requires a cloud service connection.',
    );
  });

  it('opens the chart galleries the host cannot draw with every row unavailable', () => {
    for (const [id, first] of [
      ['mac.insert.chartHierarchy', 'ツリーマップ'],
      ['mac.insert.chartStatistical', 'ヒストグラム'],
      ['mac.insert.chartWaterfall', 'ウォーターフォール'],
      ['mac.insert.chartCombo', '集合縦棒 - 折れ線'],
      ['mac.insert.chartMap', '塗り分けマップ'],
    ] as const) {
      const menu = createMacRibbonMenuFactory('ja')(id);
      const rows = Array.from(menu?.querySelectorAll<HTMLButtonElement>('[role=menuitem]') ?? []);
      expect(rows.length).toBeGreaterThan(0);
      expect(rows.every((row) => row.getAttribute('aria-disabled') === 'true')).toBe(true);
      expect(rows.some((row) => row.textContent === first)).toBe(true);
    }
  });

  it('does not expose an enabled leaf without a Mac dispatcher', () => {
    const enabledLeaves = buildRibbonModel('en', { profile: 'excel365Mac' })
      .flatMap((tab) => tab.groups)
      .flatMap((group) => group.commands)
      .filter(
        (command) =>
          command.id.startsWith('mac.') &&
          command.kind !== 'label' &&
          !command.disabled &&
          !MAC_RIBBON_MENU_COMMAND_SET.has(command.id),
      );
    const unsupported = enabledLeaves
      .filter((command) => !isMacRibbonCommandSupported(command.id))
      .map((command) => command.id);
    expect(unsupported).toEqual([]);
  });

  describe('Formulas tab layout', () => {
    const formulasGroups = () =>
      buildRibbonModel('en', { profile: 'excel365Mac' }).find((tab) => tab.id === 'formulas')
        ?.groups ?? [];
    const layoutOf = (groupTitle: string) =>
      formulasGroups()
        .find((group) => group.title === groupTitle)
        ?.commands.map((command) => [command.id, command.layout === 'stacked' ? 'small' : 'large']);

    it('draws each group as large tiles beside stacked rows in the drawn order', () => {
      expect(formulasGroups().map((group) => group.title)).toEqual([
        'Function Library',
        'Python (Preview)',
        'Defined Names',
        'Formula Auditing',
        'Calculation',
      ]);
      expect(layoutOf('Python (Preview)')).toEqual([
        ['mac.formulas.python', 'large'],
        ['mac.formulas.reset', 'small'],
        ['mac.formulas.editor', 'small'],
        ['mac.formulas.initialize', 'small'],
      ]);
      expect(layoutOf('Defined Names')).toEqual([
        ['mac.formulas.namesManager', 'large'],
        ['mac.formulas.defineName', 'small'],
        ['mac.formulas.useInFormula', 'small'],
        ['mac.formulas.createFromSelection', 'small'],
      ]);
      expect(layoutOf('Formula Auditing')).toEqual([
        ['mac.formulas.precedents', 'small'],
        ['mac.formulas.dependents', 'small'],
        ['mac.formulas.removeArrows', 'small'],
        ['mac.formulas.showFormulas', 'large'],
        ['mac.formulas.errorCheck', 'large'],
        ['mac.formulas.watch', 'large'],
      ]);
      expect(layoutOf('Calculation')).toEqual([
        ['mac.formulas.calcOptions', 'large'],
        ['mac.formulas.recalc', 'small'],
        ['mac.formulas.sheetRecalc', 'small'],
      ]);
    });

    it('renders the Python, names, auditing and calculation groups with the mixed layout', async () => {
      const sheet = await mountStubSheet({ locale: 'en' });
      const toolbarHost = document.createElement('div');
      document.body.append(toolbarHost);
      const toolbar = Spreadsheet.mountToolbar(toolbarHost, sheet.instance, {
        platform: 'mac',
        activeTab: 'formulas',
        helpers: mountedRibbonHelpers(),
      });
      try {
        const mixed = Array.from(
          toolbarHost.querySelectorAll<HTMLElement>(
            '[data-ribbon-panel="formulas"] .fc-tb__ribbon-group--mixed',
          ),
        ).map((group) => group.getAttribute('aria-label'));
        expect(mixed).toEqual([
          'Python (Preview)',
          'Defined Names',
          'Formula Auditing',
          'Calculation',
        ]);
        const reset = toolbarHost.querySelector<HTMLButtonElement>(
          '[data-ribbon-command="mac.formulas.reset"]',
        );
        expect(reset?.disabled || reset?.getAttribute('aria-disabled') === 'true').toBe(true);
        expect(reset?.querySelector('.fc-tb__rb-split-chevron')).not.toBeNull();
      } finally {
        toolbar.dispose();
        toolbarHost.remove();
        sheet.dispose();
      }
    });

    it('declares Reset, Define Name and Use in Formula as menu buttons with the drawn items', () => {
      const items = (root: string) =>
        MAC_RIBBON_MENU_ITEMS[root]?.map((item) => [item.id, item.label, !!item.disabled]);
      expect(items('mac.formulas.reset')).toEqual([
        ['mac.formulas.reset.clear', 'Reset', true],
        ['mac.formulas.reset.runtime', 'Reset Runtime', true],
        ['mac.formulas.reset.versionStable', 'Stable', true],
        ['mac.formulas.reset.versionLatest', 'Latest', true],
        ['mac.formulas.reset.versionLearnMore', 'Learn More', true],
      ]);
      expect(items('mac.formulas.defineName')).toEqual([
        ['mac.formulas.defineName.define', 'Define Name...', false],
        ['mac.formulas.defineName.apply', 'Apply...', true],
      ]);
      expect(items('mac.formulas.useInFormula')).toEqual([
        ['mac.formulas.useInFormula.paste', 'Paste Names...', false],
      ]);
      for (const root of [
        'mac.formulas.reset',
        'mac.formulas.defineName',
        'mac.formulas.useInFormula',
      ]) {
        expect(MAC_RIBBON_MENU_COMMAND_SET.has(root)).toBe(true);
        expect(ribbonActivationForCommand(root).kind).toBe('dropdown');
      }
      expect(isMacRibbonCommandSupported('mac.formulas.defineName.define')).toBe(true);
      expect(isMacRibbonCommandSupported('mac.formulas.useInFormula.paste')).toBe(true);
      expect(MAC_RIBBON_DISABLED_COMMAND_IDS.has('mac.formulas.defineName.apply')).toBe(true);
    });

    it('localizes the new menu items in every locale', () => {
      for (const locale of ['en', 'ja', 'de', 'fr', 'ko', 'th', 'zh'] as const) {
        for (const id of [
          'mac.formulas.reset.clear',
          'mac.formulas.reset.runtime',
          'mac.formulas.reset.versionStable',
          'mac.formulas.reset.versionLatest',
          'mac.formulas.reset.versionLearnMore',
          'mac.formulas.defineName.define',
          'mac.formulas.defineName.apply',
          'mac.formulas.useInFormula.paste',
        ]) {
          expect(macRibbonLabelForCommand(id, locale)).not.toBe(id);
        }
      }
      expect(macRibbonLabelForCommand('mac.formulas.defineName.apply', 'ja')).toBe('引用...');
    });

    it('leaves Ignore Error out of the Error Checking menu', () => {
      expect(MAC_RIBBON_MENU_ITEMS['mac.formulas.errorCheck']?.map((item) => item.id)).toEqual([
        'mac.formulas.errorCheck.run',
        'mac.formulas.errorCheck.trace',
        'mac.formulas.errorCheck.circular',
      ]);
      expect(isMacRibbonCommandSupported('mac.formulas.errorCheck.ignore')).toBe(false);
    });
  });

  describe('Insert tab layout', () => {
    const insertGroups = () =>
      buildRibbonModel('en', { profile: 'excel365Mac' }).find((tab) => tab.id === 'insert')
        ?.groups ?? [];
    const sizeOf = (command: { layout?: string; className?: string }) =>
      command.layout === 'stacked'
        ? 'small'
        : command.className?.includes('mac-icon-only')
          ? 'icon'
          : 'large';
    const layoutOf = (groupTitle: string) =>
      insertGroups()
        .find((group) => group.title === groupTitle)
        ?.commands.map((command) => [command.id, sizeOf(command)]);
    const items = (root: string) =>
      MAC_RIBBON_MENU_ITEMS[root]?.map((item) => [item.id, item.label, !!item.disabled]);

    it('keeps the drawn group order with Data from Picture as the third group', () => {
      expect(insertGroups().map((group) => group.title)).toEqual([
        'Tables',
        'Forms',
        'Data from Picture',
        'Illustrations',
        'Controls',
        'Charts',
        'Sparklines',
        'Filters',
        'Links',
        'Comments',
        'Text',
        'Symbols',
      ]);
      expect(layoutOf('Data from Picture')).toEqual([['mac.insert.pictureFromData', 'large']]);
    });

    it('draws Illustrations and Charts with full-height tiles beside stacked rows', () => {
      expect(layoutOf('Illustrations')).toEqual([
        ['mac.insert.photo', 'large'],
        ['mac.insert.shapes', 'large'],
        ['mac.insert.icons', 'large'],
        ['mac.insert.threeDModel', 'small'],
        ['mac.insert.smartArt', 'small'],
        ['mac.insert.screenshot', 'small'],
      ]);
      expect(layoutOf('Charts')).toEqual([
        ['mac.insert.recommendedChart', 'large'],
        ['mac.insert.chartColumn', 'icon'],
        ['mac.insert.chartLine', 'icon'],
        ['mac.insert.chartPie', 'icon'],
        ['mac.insert.chartHierarchy', 'icon'],
        ['mac.insert.chartStatistical', 'icon'],
        ['mac.insert.chartScatter', 'icon'],
        ['mac.insert.chartWaterfall', 'icon'],
        ['mac.insert.chartCombo', 'icon'],
        ['mac.insert.chartMap', 'large'],
        ['mac.insert.pivotChart', 'large'],
      ]);
    });

    it('renders Illustrations and Charts with the mixed layout and keeps menu chevrons', async () => {
      const sheet = await mountStubSheet({ locale: 'en' });
      const toolbarHost = document.createElement('div');
      document.body.append(toolbarHost);
      const toolbar = Spreadsheet.mountToolbar(toolbarHost, sheet.instance, {
        platform: 'mac',
        activeTab: 'insert',
        helpers: mountedRibbonHelpers(),
      });
      try {
        const mixed = Array.from(
          toolbarHost.querySelectorAll<HTMLElement>(
            '[data-ribbon-panel="insert"] .fc-tb__ribbon-group--mixed',
          ),
        ).map((group) => group.getAttribute('aria-label'));
        expect(mixed).toEqual(['Illustrations', 'Charts']);
        const button = (id: string) =>
          toolbarHost.querySelector<HTMLButtonElement>(`[data-ribbon-command="${id}"]`);
        for (const id of ['mac.insert.forms', 'mac.insert.threeDModel', 'mac.insert.wordArt']) {
          expect(button(id)?.querySelector('.fc-tb__rb-split-chevron'), id).not.toBeNull();
        }
        expect(
          button('mac.insert.pivotTable')?.querySelector('.fc-tb__rb-split-chevron'),
        ).toBeNull();
      } finally {
        toolbar.dispose();
        toolbarHost.remove();
        sheet.dispose();
      }
    });

    it('declares the drawn menu buttons with the desktop items', () => {
      expect(items('mac.insert.forms')).toEqual([
        ['mac.insert.forms.new', 'New Form', true],
        ['mac.insert.forms.preview', 'Preview Form', true],
        ['mac.insert.forms.edit', 'Edit Form', true],
        ['mac.insert.forms.send', 'Send Form', true],
      ]);
      expect(items('mac.insert.pictureFromData')).toEqual([
        ['mac.insert.pictureFromData.file', 'Picture From File...', true],
        ['mac.insert.pictureFromData.clipboard', 'Picture From Clipboard', true],
      ]);
      expect(items('mac.insert.photo')).toEqual([
        ['mac.insert.photo.browser', 'Photo Browser…', true],
        ['mac.insert.photo.file', 'Picture from File...', false],
        ['mac.insert.photo.stock', 'Stock Images...', true],
        ['mac.insert.photo.online', 'Online Pictures...', true],
      ]);
      expect(items('mac.insert.threeDModel')).toEqual([
        ['mac.insert.threeDModel.device', 'This Device...', true],
        ['mac.insert.threeDModel.stock', 'Stock 3D Models...', true],
      ]);
      expect(items('mac.insert.smartArt')?.map((item) => item[1])).toEqual([
        'List',
        'Process',
        'Cycle',
        'Hierarchy',
        'Relationship',
        'Matrix',
        'Pyramid',
        'Picture',
      ]);
      expect(items('mac.insert.screenshot')).toEqual([
        ['mac.insert.screenshot.clipping', 'Screen Clipping', false],
      ]);
      expect(items('mac.insert.sparkline')).toEqual([
        ['mac.insert.sparkline.line', 'Line', false],
        ['mac.insert.sparkline.column', 'Column', false],
        ['mac.insert.sparkline.winLoss', 'Win/Loss', false],
      ]);
      expect(items('mac.insert.textBox')).toEqual([
        ['mac.insert.textBox.horizontal', 'Horizontal Text Box', true],
        ['mac.insert.textBox.vertical', 'Vertical Text Box', true],
      ]);
      expect(items('mac.insert.wordArt')).toEqual([]);
      expect(items('mac.insert.equation')).toEqual([
        ['mac.insert.equation.insert', 'Equation', true],
      ]);
      for (const root of [
        'mac.insert.forms',
        'mac.insert.pictureFromData',
        'mac.insert.photo',
        'mac.insert.threeDModel',
        'mac.insert.smartArt',
        'mac.insert.screenshot',
        'mac.insert.sparkline',
        'mac.insert.textBox',
        'mac.insert.wordArt',
        'mac.insert.equation',
      ]) {
        expect(MAC_RIBBON_MENU_COMMAND_SET.has(root), root).toBe(true);
        expect(ribbonActivationForCommand(root).kind, root).toBe('dropdown');
      }
    });

    it('keeps the existing actions reachable from the menu items', () => {
      for (const id of [
        'mac.insert.photo.file',
        'mac.insert.screenshot.clipping',
        'mac.insert.sparkline.line',
      ]) {
        expect(isMacRibbonCommandSupported(id), id).toBe(true);
      }
    });

    it('keeps disabled controls disabled and PivotTable a plain button', () => {
      for (const id of [
        'mac.insert.forms',
        'mac.insert.pictureFromData',
        'mac.insert.icons',
        'mac.insert.threeDModel',
        'mac.insert.smartArt',
        'mac.insert.checkBox',
        'mac.insert.pivotChart',
        'mac.insert.textBox',
        'mac.insert.wordArt',
        'mac.insert.object',
        'mac.insert.equation',
      ]) {
        expect(findCommand('insert', id)?.disabled, id).toBe(true);
      }
      expect(findCommand('insert', 'mac.insert.symbol')?.disabled).toBeUndefined();
      expect(MAC_RIBBON_MENU_COMMAND_SET.has('mac.insert.pivotTable')).toBe(false);
      expect(MAC_RIBBON_MENU_ITEMS['mac.insert.pivotTable']).toBeUndefined();
      expect(macRibbonMenuIdForCommand('mac.insert.pivotTable')).toBeNull();
      expect(ribbonActivationForCommand('mac.insert.pivotTable')).toEqual({
        kind: 'primaryAction',
      });
    });

    it('titles the picture group Data from Picture and localizes the new items', () => {
      const titles: Record<string, string> = {
        en: 'Data from Picture',
        ja: '画像からのデータ',
        de: 'Daten aus Bild',
        fr: 'Données de l’image',
        ko: '그림의 데이터',
        th: 'ข้อมูลจากรูปภาพ',
        zh: '来自图片的数据',
      };
      for (const [locale, title] of Object.entries(titles)) {
        const insert = buildRibbonModel(locale as 'en', { profile: 'excel365Mac' }).find(
          (tab) => tab.id === 'insert',
        );
        expect(insert?.groups[2]?.title, locale).toBe(title);
      }
      for (const locale of ['en', 'ja', 'de', 'fr', 'ko', 'th', 'zh'] as const) {
        for (const root of [
          'mac.insert.forms',
          'mac.insert.pictureFromData',
          'mac.insert.photo',
          'mac.insert.threeDModel',
          'mac.insert.smartArt',
          'mac.insert.screenshot',
          'mac.insert.sparkline',
          'mac.insert.textBox',
          'mac.insert.equation',
        ]) {
          for (const item of MAC_RIBBON_MENU_ITEMS[root] ?? []) {
            if (item.id === root) continue;
            expect(macRibbonLabelForCommand(item.id, locale), `${locale} ${item.id}`).not.toBe(
              item.id,
            );
          }
        }
      }
      expect(macRibbonLabelForCommand('mac.insert.photo.browser', 'ja')).toBe('写真ブラウザー…');
      expect(macRibbonLabelForCommand('mac.insert.smartArt.hierarchy', 'ja')).toBe('階層構造');
    });
  });

  describe('Draw tab layout', () => {
    const drawGroups = (lang: 'en' | 'ja' = 'en') =>
      buildRibbonModel(lang, { profile: 'excel365Mac' }).find((tab) => tab.id === 'draw')?.groups ??
      [];
    const sizeOf = (command: { layout?: string; className?: string }) =>
      command.layout === 'stacked'
        ? 'small'
        : command.className?.includes('mac-icon-only')
          ? 'icon'
          : 'large';
    const layoutOf = (groupTitle: string) =>
      drawGroups()
        .find((group) => group.title === groupTitle)
        ?.commands.map((command) => [command.id, sizeOf(command)]);
    const items = (root: string) =>
      MAC_RIBBON_MENU_ITEMS[root]?.map((item) => [item.id, item.label, !!item.disabled]);

    it('draws Tools as three large tiles, Add beside the pen strip and Trackpad as a large tile', () => {
      expect(drawGroups().map((group) => group.title)).toEqual(['Tools', 'Pens', 'Mode']);
      expect(layoutOf('Tools')).toEqual([
        ['mac.draw.toggle', 'large'],
        ['mac.draw.eraser', 'large'],
        ['mac.draw.lasso', 'large'],
      ]);
      expect(layoutOf('Pens')).toEqual([
        ['mac.draw.penBlack', 'icon'],
        ['mac.draw.penRed', 'icon'],
        ['mac.draw.pencil', 'icon'],
        ['mac.draw.highlighter', 'icon'],
        ['mac.draw.add', 'large'],
      ]);
      expect(layoutOf('Mode')).toEqual([['mac.draw.trackpad', 'large']]);
      expect(findCommand('draw', 'mac.draw.lasso')?.disabled).toBe(true);
      expect(findCommand('draw', 'mac.draw.lasso')?.disabledReason).toBeTruthy();
    });

    it('names each pen by its description and leaves the pen buttons unlabelled', () => {
      const pens = drawGroups('ja').find((group) => group.title === 'ペン')?.commands ?? [];
      expect(pens.slice(0, 4).map((command) => command.title)).toEqual([
        'ペン: 黒、0.35 mm',
        'ペン: 赤、0.35 mm',
        '鉛筆書き: 灰色、0.5 mm',
        '蛍光ペン: 黄、6 mm',
      ]);
      expect(pens.slice(0, 4).map((command) => command.label)).toEqual(
        pens.slice(0, 4).map((command) => command.title),
      );
      expect(macRibbonLabelForCommand('mac.draw.penBlack', 'en')).toBe('Pen: Black, 0.35 mm');
      expect(macRibbonLabelForCommand('mac.draw.pencil', 'en')).toBe('Pencil: Gray, 0.5 mm');
      expect(macRibbonLabelForCommand('mac.draw.penRed', 'fr')).toBe('Stylo : Rouge, 0.35 mm');
      expect(macRibbonLabelForCommand('mac.draw.highlighter', 'zh')).toBe('荧光笔: 黄色，6 毫米');
    });

    it('declares Add as a menu of three pen kinds that cannot be added', () => {
      expect(MAC_RIBBON_MENU_COMMAND_SET.has('mac.draw.add')).toBe(true);
      expect(findCommand('draw', 'mac.draw.add')?.disabled).toBeFalsy();
      expect(items('mac.draw.add')).toEqual([
        ['mac.draw.add.pen', 'Pen', true],
        ['mac.draw.add.highlighter', 'Highlighter', true],
        ['mac.draw.add.pencil', 'Pencil', true],
      ]);
      for (const item of MAC_RIBBON_MENU_ITEMS['mac.draw.add'] ?? []) {
        expect(item.disabledReason, item.id).toBeTruthy();
      }
      expect(macRibbonLabelForCommand('mac.draw.add.highlighter', 'ja')).toBe('蛍光ペン');
    });

    it('opens the Eraser menu with the four eraser variants and only the stroke eraser enabled', () => {
      expect(items('mac.draw.eraser')).toEqual([
        ['mac.draw.eraser.stroke', 'Stroke Eraser', false],
        ['mac.draw.eraser.small', 'Small Eraser', true],
        ['mac.draw.eraser.medium', 'Medium Eraser', true],
        ['mac.draw.eraser.segment', 'Segment Eraser', true],
      ]);
      expect(macRibbonLabelForCommand('mac.draw.eraser.segment', 'ja')).toBe('セグメント消しゴム');
      expect(macRibbonLabelForCommand('mac.draw.eraser.stroke', 'ja')).toBe(
        '消しゴム (ストローク)',
      );
      expect(isMacRibbonCommandSupported('mac.draw.eraser.stroke')).toBe(true);
      expect(isMacRibbonCommandSupported('mac.draw.eraser.small')).toBe(false);
    });

    it('renders the Draw groups with the mixed layout, toggles for Draw, Lasso and Trackpad, and unlabelled pens', async () => {
      const sheet = await mountStubSheet({ locale: 'en' });
      const toolbarHost = document.createElement('div');
      document.body.append(toolbarHost);
      const toolbar = Spreadsheet.mountToolbar(toolbarHost, sheet.instance, {
        platform: 'mac',
        activeTab: 'draw',
        helpers: {
          ...mountedRibbonHelpers(),
          createIcon: () => document.createElementNS('http://www.w3.org/2000/svg', 'svg'),
        },
      });
      try {
        projectMacRibbonState(toolbarHost, sheet.instance);
        const mixed = Array.from(
          toolbarHost.querySelectorAll<HTMLElement>(
            '[data-ribbon-panel="draw"] .fc-tb__ribbon-group--mixed',
          ),
        ).map((group) => group.getAttribute('aria-label'));
        expect(mixed).toEqual(['Tools', 'Pens', 'Mode']);
        const button = (id: string) =>
          toolbarHost.querySelector<HTMLButtonElement>(`[data-ribbon-command="${id}"]`);
        for (const id of ['mac.draw.toggle', 'mac.draw.lasso', 'mac.draw.trackpad']) {
          expect(button(id)?.getAttribute('aria-pressed'), id).toBe('false');
        }
        for (const id of ['mac.draw.eraser', 'mac.draw.add', 'mac.draw.lasso']) {
          expect(button(id)?.querySelector('.fc-tb__rb-split-chevron') !== null, id).toBe(
            id !== 'mac.draw.lasso',
          );
        }
        const pen = button('mac.draw.penBlack');
        expect(pen?.getAttribute('aria-label')).toBe('Pen: Black, 0.35 mm');
        expect(pen?.textContent).toBe('');
      } finally {
        toolbar.dispose();
        toolbarHost.remove();
        sheet.dispose();
      }
    });
  });

  describe('Page Layout tab layout', () => {
    const pageGroups = () =>
      buildRibbonModel('en', { profile: 'excel365Mac' }).find((tab) => tab.id === 'pageLayout')
        ?.groups ?? [];
    const sizeOf = (command: { layout?: string }) =>
      command.layout === 'stacked' ? 'small' : 'large';
    const layoutOf = (groupTitle: string) =>
      pageGroups()
        .find((group) => group.title === groupTitle)
        ?.commands.map((command) => [command.id, sizeOf(command)]);
    const items = (root: string) =>
      MAC_RIBBON_MENU_ITEMS[root]?.map((item) => [item.id, item.label, !!item.disabled]);

    it('draws Page Setup as eight large tiles and Themes as a large tile beside two small rows', () => {
      expect(pageGroups().map((group) => group.title)).toEqual([
        'Themes',
        'Page Setup',
        'Scale to Fit',
        'Sheet Options',
      ]);
      expect(layoutOf('Themes')).toEqual([
        ['mac.page.theme', 'large'],
        ['mac.page.color', 'small'],
        ['mac.page.font', 'small'],
      ]);
      expect(layoutOf('Page Setup')).toEqual([
        ['mac.page.margins', 'large'],
        ['mac.page.orientation', 'large'],
        ['mac.page.size', 'large'],
        ['mac.page.printArea', 'large'],
        ['mac.page.pageBreaks', 'large'],
        ['mac.page.background', 'large'],
        ['mac.page.printTitles', 'large'],
        ['mac.page.pageSetup', 'large'],
      ]);
    });

    it('keeps the Scale to Fit combos as selects', () => {
      const scale = pageGroups().find((group) => group.title === 'Scale to Fit');
      expect(scale?.commands.map((command) => [command.id, command.kind])).toEqual([
        ['scaleWidth', 'select'],
        ['scaleHeight', 'select'],
      ]);
    });

    it('lists the desktop theme commands after the working themes, disabled with a reason', () => {
      expect(items('mac.page.theme')).toEqual([
        ['mac.page.theme.office', 'Office', false],
        ['mac.page.theme.gray', 'Grayscale', false],
        ['mac.page.theme.enableUpdates', 'Enable Content Updates', true],
        ['mac.page.theme.browse', 'Browse for Themes...', true],
        ['mac.page.theme.save', 'Save Current Theme...', true],
      ]);
      for (const item of MAC_RIBBON_MENU_ITEMS['mac.page.theme'] ?? []) {
        if (item.disabled) expect(item.disabledReason, item.id).toBeTruthy();
      }
      expect(items('mac.page.color')?.map(([id]) => id)).toEqual([
        'mac.page.color.office',
        'mac.page.color.gray',
      ]);
      expect(macRibbonLabelForCommand('mac.page.theme.browse', 'ja')).toBe('テーマを参照...');
      expect(macRibbonLabelForCommand('mac.page.theme.save', 'fr')).toBe(
        'Enregistrer le thème actif...',
      );
      expect(macRibbonLabelForCommand('mac.page.theme.enableUpdates', 'zh')).toBe('启用内容更新');
    });

    it('gives the View tab button its own label while the tab keeps the Page Layout wording', () => {
      const viewLabels = {
        en: 'Page Layout',
        ja: 'ページ レイアウト',
        de: 'Seitenlayout',
        fr: 'Mode Page',
        zh: '页面布局',
        ko: '페이지 레이아웃',
        th: 'เค้าโครงหน้ากระดาษ',
      } as const;
      const tabLabels = {
        en: 'Page Layout',
        ja: 'ページ レイアウト',
        de: 'Seitenlayout',
        fr: 'Mise en page',
        zh: '页面布局',
        ko: '페이지 레이아웃',
        th: 'เค้าโครงหน้ากระดาษ',
      } as const;
      for (const locale of Object.keys(viewLabels) as (keyof typeof viewLabels)[]) {
        const model = buildRibbonModel(locale, { profile: 'excel365Mac' });
        const button = model
          .find((tab) => tab.id === 'view')
          ?.groups.flatMap((group) => group.commands)
          .find((command) => command.id === 'mac.view.pageLayout');
        expect(button?.label, locale).toBe(viewLabels[locale]);
        expect(button?.title, locale).toBe(viewLabels[locale]);
        expect(model.find((tab) => tab.id === 'pageLayout')?.label, locale).toBe(tabLabels[locale]);
      }
    });

    it('renders Themes and Page Setup with the mixed layout and keeps the theme menu chevrons', async () => {
      const sheet = await mountStubSheet({ locale: 'en' });
      const toolbarHost = document.createElement('div');
      document.body.append(toolbarHost);
      const toolbar = Spreadsheet.mountToolbar(toolbarHost, sheet.instance, {
        platform: 'mac',
        activeTab: 'pageLayout',
        helpers: {
          ...mountedRibbonHelpers(),
          createIcon: () => document.createElementNS('http://www.w3.org/2000/svg', 'svg'),
        },
      });
      try {
        projectMacRibbonState(toolbarHost, sheet.instance);
        const mixed = Array.from(
          toolbarHost.querySelectorAll<HTMLElement>(
            '[data-ribbon-panel="pageLayout"] .fc-tb__ribbon-group--mixed',
          ),
        ).map((group) => group.getAttribute('aria-label'));
        expect(mixed).toEqual(['Themes', 'Page Setup']);
        const button = (id: string) =>
          toolbarHost.querySelector<HTMLButtonElement>(`[data-ribbon-command="${id}"]`);
        for (const id of ['mac.page.theme', 'mac.page.color', 'mac.page.font']) {
          expect(button(id)?.querySelector('.fc-tb__rb-split-chevron') !== null, id).toBe(true);
        }
      } finally {
        toolbar.dispose();
        toolbarHost.remove();
        sheet.dispose();
      }
    });
  });

  describe('Review tab layout', () => {
    const reviewGroups = () =>
      buildRibbonModel('en', { profile: 'excel365Mac' }).find((tab) => tab.id === 'review')
        ?.groups ?? [];
    const sizeOf = (command: { layout?: string }) =>
      command.layout === 'stacked' ? 'small' : 'large';
    const items = (root: string) =>
      MAC_RIBBON_MENU_ITEMS[root]?.map((item) => [item.id, item.label, !!item.disabled]);
    const commandIds = [
      'mac.review.spelling',
      'mac.review.thesaurus',
      'mac.review.stats',
      'mac.review.accessibility',
      'mac.review.translate',
      'mac.review.showChanges',
      'mac.review.newComment',
      'mac.review.deleteComment',
      'mac.review.previousComment',
      'mac.review.nextComment',
      'mac.review.showComments',
      'mac.review.notes',
      'mac.review.protectSheet',
      'mac.review.protectWorkbook',
      'mac.review.alwaysReadOnly',
      'mac.review.hideInk',
    ];

    it('draws all sixteen controls as large tiles in the desktop order', () => {
      const commands = reviewGroups().flatMap((group) => group.commands);
      expect(commands.map((command) => [command.id, sizeOf(command)])).toEqual(
        commandIds.map((id) => [id, 'large']),
      );
      expect(reviewGroups().map((group) => group.commands.length)).toEqual([
        3, 1, 1, 1, 5, 1, 2, 1, 1,
      ]);
    });

    it('keeps Thesaurus, Translate, Show Changes, Always Open Read-Only and Hide Ink disabled with a reason', () => {
      for (const id of [
        'mac.review.thesaurus',
        'mac.review.translate',
        'mac.review.showChanges',
        'mac.review.alwaysReadOnly',
        'mac.review.hideInk',
      ]) {
        expect(findCommand('review', id)?.disabled, id).toBe(true);
        expect(findCommand('review', id)?.disabledReason, id).toBeTruthy();
      }
    });

    it('opens Check Accessibility as a menu of five items, Alt Text and Help disabled', () => {
      expect(MAC_RIBBON_MENU_COMMAND_SET.has('mac.review.accessibility')).toBe(true);
      expect(findCommand('review', 'mac.review.accessibility')?.disabled).toBeFalsy();
      expect(items('mac.review.accessibility')).toEqual([
        ['mac.review.accessibility.check', 'Check Accessibility', false],
        ['mac.review.accessibility.altText', 'Alt Text', true],
        ['mac.review.accessibility.formatTable', 'Format as Table', false],
        ['mac.review.accessibility.unmerge', 'Unmerge Cells', false],
        ['mac.review.accessibility.help', 'Accessibility Help', true],
      ]);
      for (const item of MAC_RIBBON_MENU_ITEMS['mac.review.accessibility'] ?? []) {
        expect(!!item.disabledReason, item.id).toBe(!!item.disabled);
      }
      expect(isMacRibbonCommandSupported('mac.review.accessibility.check')).toBe(true);
      expect(macRibbonLabelForCommand('mac.review.accessibility.unmerge', 'ja')).toBe(
        'セル結合の解除',
      );
    });

    it('opens Hide Ink as a menu of three items, all disabled with a reason', () => {
      expect(MAC_RIBBON_MENU_COMMAND_SET.has('mac.review.hideInk')).toBe(true);
      expect(items('mac.review.hideInk')).toEqual([
        ['mac.review.hideInk.hide', 'Hide Ink', true],
        ['mac.review.hideInk.deleteSheet', 'Delete All Ink on Sheet', true],
        ['mac.review.hideInk.deleteWorkbook', 'Delete All Ink in Workbook', true],
      ]);
      for (const item of MAC_RIBBON_MENU_ITEMS['mac.review.hideInk'] ?? []) {
        expect(item.disabledReason, item.id).toBeTruthy();
      }
      expect(macRibbonLabelForCommand('mac.review.hideInk.deleteSheet', 'ja')).toBe(
        'シート内のすべてのインクを削除',
      );
    });

    it('labels the menu items in all seven locales', () => {
      for (const locale of ['en', 'ja', 'de', 'fr', 'ko', 'th', 'zh'] as const) {
        for (const root of ['mac.review.accessibility', 'mac.review.hideInk']) {
          for (const item of MAC_RIBBON_MENU_ITEMS[root] ?? []) {
            expect(macRibbonLabelForCommand(item.id, locale), `${locale} ${item.id}`).not.toBe(
              item.id,
            );
          }
        }
      }
      expect(macRibbonLabelForCommand('mac.review.accessibility.altText', 'fr')).toBe('Texte Alt');
      expect(macRibbonLabelForCommand('mac.review.hideInk.hide', 'zh')).toBe('隐藏墨迹');
    });

    it('titles the Protect Sheet group Changes and the read-only group Protect in every language', () => {
      const titles: Record<string, [string, string]> = {
        en: ['Changes', 'Protect'],
        ja: ['変更箇所', '保護'],
        de: ['Änderungen', 'Schützen'],
        fr: ['Modifications', 'Protéger'],
        ko: ['변경 내용', '보호'],
        th: ['การเปลี่ยนแปลง', 'ป้องกัน'],
        zh: ['更改', '保护'],
      };
      for (const [locale, [changes, protect]] of Object.entries(titles)) {
        const groups = buildRibbonModel(locale as 'en', { profile: 'excel365Mac' }).find(
          (tab) => tab.id === 'review',
        )?.groups;
        expect(groups?.[6]?.title, locale).toBe(changes);
        expect(groups?.[6]?.commands.map((command) => command.id)).toEqual([
          'mac.review.protectSheet',
          'mac.review.protectWorkbook',
        ]);
        expect(groups?.[7]?.title, locale).toBe(protect);
        expect(groups?.[7]?.commands.map((command) => command.id)).toEqual([
          'mac.review.alwaysReadOnly',
        ]);
      }
    });

    it('renders four toggles pressed only where the model holds the state, and the menu chevrons', async () => {
      const sheet = await mountStubSheet({ locale: 'en' });
      const toolbarHost = document.createElement('div');
      document.body.append(toolbarHost);
      const toolbar = Spreadsheet.mountToolbar(toolbarHost, sheet.instance, {
        platform: 'mac',
        activeTab: 'review',
        helpers: {
          ...mountedRibbonHelpers(),
          createIcon: () => document.createElementNS('http://www.w3.org/2000/svg', 'svg'),
        },
      });
      try {
        projectMacRibbonState(toolbarHost, sheet.instance);
        const button = (id: string) =>
          toolbarHost.querySelector<HTMLButtonElement>(`[data-ribbon-command="${id}"]`);
        const toggles = [
          'mac.review.showChanges',
          'mac.review.showComments',
          'mac.review.protectWorkbook',
          'mac.review.alwaysReadOnly',
        ];
        for (const id of toggles) {
          expect(button(id)?.getAttribute('aria-pressed'), id).toBe('false');
        }
        expect(button('mac.review.spelling')?.hasAttribute('aria-pressed')).toBe(false);
        for (const id of ['mac.review.accessibility', 'mac.review.hideInk', 'mac.review.notes']) {
          expect(button(id)?.querySelector('.fc-tb__rb-split-chevron') !== null, id).toBe(true);
        }
        expect(button('mac.review.spelling')?.querySelector('.fc-tb__rb-split-chevron')).toBeNull();
        mutators.setWorkbookStructureProtected(sheet.instance.store, true);
        projectMacRibbonState(toolbarHost, sheet.instance);
        expect(button('mac.review.protectWorkbook')?.getAttribute('aria-pressed')).toBe('true');
        for (const id of toggles.filter((id) => id !== 'mac.review.protectWorkbook')) {
          expect(button(id)?.getAttribute('aria-pressed'), id).toBe('false');
        }
      } finally {
        toolbar.dispose();
        toolbarHost.remove();
        sheet.dispose();
      }
    });

    it('keeps Show Comments enabled on a sheet without comments and opens an empty list', async () => {
      const sheet = await mountStubSheet({
        locale: 'en',
        workbook: await WorkbookHandle.createDefault(),
      });
      const { instance } = sheet;
      expect(instance.workbook.capabilities.threadedComments).toBe(true);
      try {
        const host = document.createElement('div');
        const button = projectedButton(host, 'mac.review.showComments');
        projectMacRibbonState(host, instance);
        expect(button.disabled).toBe(false);
        expect(instance.store.getState().threadedComments.bySheet.get(0) ?? []).toEqual([]);

        const deps = {
          inst: instance,
          runtime: { projectFormatToolbar: () => undefined },
        } as unknown as ApplyRibbonCommandDeps;
        expect(dispatchMacRibbonCommand('mac.review.showComments', deps, () => false)).toBe(true);
        // The list opens with its empty notice and no rows.
        const list = document.querySelector('.fc-tb__dlg__list');
        expect(list?.textContent).toBe(instance.i18n.strings.reviewReports.noIssues);
        expect(list?.querySelectorAll('strong').length).toBe(0);
      } finally {
        document.body.innerHTML = '';
        sheet.dispose();
      }
    });
  });

  describe('Automate tab layout', () => {
    const toolsGroup = () =>
      buildRibbonModel('en', { profile: 'excel365Mac' })
        .find((tab) => tab.id === 'automate')
        ?.groups.find((group) => group.title === 'Office Scripts');
    const items = (root: string) =>
      MAC_RIBBON_MENU_ITEMS[root]?.map((item) => [item.id, item.label, !!item.disabled]);

    it('draws New Script and View Scripts as large menu buttons and no Script Gallery button', () => {
      const commands = toolsGroup()?.commands ?? [];
      expect(commands.map((command) => [command.id, command.layout])).toEqual([
        ['mac.automate.newScript', undefined],
        ['mac.automate.showScripts', undefined],
      ]);
      for (const root of ['mac.automate.newScript', 'mac.automate.showScripts']) {
        expect(MAC_RIBBON_MENU_COMMAND_SET.has(root), root).toBe(true);
        expect(ribbonActivationForCommand(root), root).toEqual({
          kind: 'dropdown',
          menuId: macRibbonMenuIdForCommand(root),
        });
      }
      expect(MAC_RIBBON_MENU_COMMAND_SET.has('mac.automate.gallery')).toBe(false);
      expect(MAC_RIBBON_MENU_ITEMS['mac.automate.gallery']).toBeUndefined();
    });

    it('keeps New Script disabled with its reason and every item of its menu disabled', () => {
      expect(findCommand('automate', 'mac.automate.newScript')?.disabled).toBe(true);
      expect(findCommand('automate', 'mac.automate.newScript')?.disabledReason).toBeTruthy();
      expect(items('mac.automate.newScript')).toEqual([
        ['mac.automate.newScript.editor', 'Create in Code Editor', true],
        ['mac.automate.newScript.recording', 'Create from Recording', true],
      ]);
      for (const item of MAC_RIBBON_MENU_ITEMS['mac.automate.newScript'] ?? []) {
        expect(item.disabledReason, item.id).toBeTruthy();
      }
    });

    it('opens View Scripts as a menu where only Samples is enabled', () => {
      expect(findCommand('automate', 'mac.automate.showScripts')?.disabled).toBeFalsy();
      expect(items('mac.automate.showScripts')).toEqual([
        ['mac.automate.showScripts.recent', 'Recent Scripts', true],
        ['mac.automate.showScripts.workbook', 'Scripts from this Workbook', true],
        ['mac.automate.showScripts.samples', 'Samples', false],
      ]);
      for (const item of MAC_RIBBON_MENU_ITEMS['mac.automate.showScripts'] ?? []) {
        expect(!!item.disabledReason, item.id).toBe(!!item.disabled);
      }
    });

    it('labels the menu items in all seven locales', () => {
      for (const locale of ['en', 'ja', 'de', 'fr', 'ko', 'th', 'zh'] as const) {
        for (const root of ['mac.automate.newScript', 'mac.automate.showScripts']) {
          for (const item of MAC_RIBBON_MENU_ITEMS[root] ?? []) {
            expect(macRibbonLabelForCommand(item.id, locale), `${locale} ${item.id}`).not.toBe(
              item.id,
            );
          }
        }
      }
      expect(macRibbonLabelForCommand('mac.automate.showScripts.samples', 'ja')).toBe('サンプル');
      expect(macRibbonLabelForCommand('mac.automate.newScript.editor', 'de')).toBe(
        'Im Code-Editor erstellen',
      );
    });

    it('carries no Script Gallery id anywhere in the Mac model', () => {
      for (const profile of ['default', 'excel365Mac'] as const) {
        const ids = buildRibbonModel('en', { profile }).flatMap((tab) =>
          tab.groups.flatMap((group) => group.commands.map((command) => command.id)),
        );
        expect(ids).not.toContain('mac.automate.gallery');
      }
      expect(MAC_RIBBON_COMMAND_IDS).not.toContain('mac.automate.gallery');
      expect(COMMAND_IDS as readonly string[]).not.toContain('mac.automate.gallery');
    });

    it('opens the script gallery dialog from View Scripts > Samples', async () => {
      const sheet = await mountStubSheet({ locale: 'en' });
      try {
        const deps = {
          inst: sheet.instance,
          runtime: { projectFormatToolbar: () => undefined },
        } as unknown as ApplyRibbonCommandDeps;
        expect(isMacRibbonCommandSupported('mac.automate.showScripts.samples')).toBe(true);
        expect(document.querySelector('.fc-macautomationdlg')).toBeNull();
        expect(
          dispatchMacRibbonCommand('mac.automate.showScripts.samples', deps, () => false),
        ).toBe(true);
        expect(document.querySelector('.fc-macautomationdlg')).not.toBeNull();
        expect(document.querySelectorAll('[data-mac-automation-script]').length).toBe(7);
      } finally {
        document.body.innerHTML = '';
        sheet.dispose();
      }
    });
  });

  it('names the Draw tab controls as Excel for Mac does', () => {
    const draw = buildRibbonModel('ja', { profile: 'excel365Mac' }).find(
      (tab) => tab.id === 'draw',
    );
    const labels = draw?.groups.flatMap((group) => group.commands.map((command) => command.label));
    expect(labels).toContain('なげなわ選択');
    expect(labels).toContain('トラックパッドで 描画');
    expect(
      draw?.groups.flatMap((g) => g.commands).find((c) => c.id === 'mac.draw.lasso')?.disabled,
    ).toBe(true);
  });
});
