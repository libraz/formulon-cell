import { describe, expect, it, vi } from 'vitest';
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
import type { SpreadsheetInstance } from '../../../src/mount/types.js';
import { Spreadsheet } from '../../../src/mount.js';
import { createSpreadsheetStore, defaultPageSetup, mutators } from '../../../src/store/store.js';
import { ribbonActivationForCommand } from '../../../src/toolbar/ribbon/activation.js';
import { isMacRibbonCommandSupported } from '../../../src/toolbar/ribbon/mac/dispatch.js';
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
    workbook: { calcMode: () => 'auto' },
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
    const scenario = menu?.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="mac.data.scenarioManager"]',
    );
    expect(scenario?.getAttribute('aria-disabled')).toBe('true');
    expect(scenario?.title).toContain('この機能は未実装です。');
    expect(scenario?.title).not.toContain('Microsoft 365');
    expect(scenario?.textContent).toBe('シナリオ マネージャー');
    const shapes = createMacRibbonMenuFactory('ja')('mac.insert.shapes');
    expect(
      Array.from(shapes?.querySelectorAll('[role=menuitem]') ?? []).map((item) => item.textContent),
    ).toEqual(['直線', '矢印', '四角形', '角丸四角形', '楕円', '三角形', 'ひし形', '吹き出し']);
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
        if (command.disabled) {
          expect(ribbonActivationForCommand(command.id)).toEqual({ kind: 'disabled' });
        } else {
          expect(ribbonActivationForCommand(command.id)).toEqual({
            kind: 'dropdown',
            menuId: macRibbonMenuIdForCommand(command.id),
          });
        }
        if (command.id !== 'mac.formulas.more') {
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
    expect(fontIds.indexOf('editPhonetic')).toBeGreaterThan(fontIds.indexOf('fontColor'));
    expect(home?.groups.at(-1)?.commands.map((command) => command.id)).toEqual(['mac.home.addins']);
    expect(findCommand('home', 'mac.home.addins')).toMatchObject({
      disabled: true,
      disabledReason: 'Requires a cloud service connection.',
    });

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
      'mac.page.showGridlines',
      'mac.page.printGridlines',
      'mac.page.showHeadings',
      'mac.page.printHeadings',
    ]);
    expect(sheetOptions?.commands.every((command) => command.icon === undefined)).toBe(true);

    const view = model.find((tab) => tab.id === 'view');
    expect(
      view?.groups.find((group) => group.title === 'Sheet Views')?.commands.map((c) => c.id),
    ).toEqual(['sheetViewSelect', 'mac.view.sheetViewSave', 'mac.view.sheetViewDelete']);
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
      automate?.groups.find((group) => group.title === 'Automation')?.commands.map((c) => c.id),
    ).toContain('mac.automate.gallery');
    expect(automate?.groups.find((group) => group.title === 'Gallery')?.variant).toBe('gallery');
    expect(
      automate?.groups.find((group) => group.title === 'Gallery')?.commands.map((c) => c.id),
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
      'mac.formulas.errorCheck.ignore',
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
        filterRange: range,
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
      for (const item of items) {
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
    expect(command('mac.view.sheetViewSave')).toMatchObject({ disabled: undefined });
    expect(command('mac.view.sheetViewDelete')).toMatchObject({ disabled: undefined });
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
      'mac.insert.chartHierarchy',
      'mac.insert.chartStatistical',
      'mac.insert.chartWaterfall',
      'mac.insert.chartCombo',
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

  it('does not expose an enabled leaf without a Mac dispatcher', () => {
    const enabledLeaves = buildRibbonModel('en', { profile: 'excel365Mac' })
      .flatMap((tab) => tab.groups)
      .flatMap((group) => group.commands)
      .filter(
        (command) =>
          command.id.startsWith('mac.') &&
          !command.disabled &&
          !MAC_RIBBON_MENU_COMMAND_SET.has(command.id),
      );
    const unsupported = enabledLeaves
      .filter((command) => !isMacRibbonCommandSupported(command.id))
      .map((command) => command.id);
    expect(unsupported).toEqual([]);
  });
});
