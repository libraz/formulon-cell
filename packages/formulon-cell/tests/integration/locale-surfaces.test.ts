import { existsSync, readdirSync, readFileSync, statSync } from 'node:fs';
import { relative, resolve } from 'node:path';
import { afterEach, describe, expect, it } from 'vitest';

import { formatAsTableByStyleId } from '../../src/commands/format-as-table.js';
import { isWorkbookStructureProtected } from '../../src/commands/protection.js';
import { WorkbookHandle } from '../../src/engine/workbook-handle.js';
import type { Strings } from '../../src/i18n/strings/types.js';
import { dictionaries } from '../../src/i18n/strings.js';
import type { SpreadsheetInstance } from '../../src/mount/types.js';
import { mutators } from '../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../test-utils/index.js';
import {
  LOCALE_SURFACE_FACTORIES,
  LOCALE_SURFACE_LOCALES,
  LOCALE_SURFACE_SECTIONS,
  LOCALE_SURFACE_SHELL_FACTORIES,
  LOCALE_SURFACE_TOOLBAR_ROUTES,
  LOCALE_SURFACE_TOOLBAR_UNAVAILABLE,
  type LocaleSurfaceLocale,
  type LocaleSurfaceSection,
  type LocaleSurfaceToolbarRoute,
} from '../test-utils/locale-surface-manifest.js';

// Vitest serves transformed modules with a non-file `import.meta.url`; the
// package test command runs from this workspace package, so cwd is the stable
// source anchor for the source-level manifest.
const sourceRoot =
  [resolve(process.cwd(), 'src'), resolve(process.cwd(), 'packages/formulon-cell/src')].find(
    (candidate) => existsSync(candidate),
  ) ?? resolve(process.cwd(), 'src');

const asPosix = (path: string): string => path.replaceAll('\\', '/');

function sourceFiles(root: string): string[] {
  const result: string[] = [];
  for (const entry of readdirSync(root, { withFileTypes: true })) {
    const path = resolve(root, entry.name);
    if (entry.isDirectory()) result.push(...sourceFiles(path));
    else if (entry.isFile() && path.endsWith('.ts')) result.push(path);
  }
  return result;
}

function sourceText(relativePath: string): string {
  return readFileSync(resolve(sourceRoot, relativePath), 'utf8');
}

function relativeSource(path: string): string {
  return asPosix(relative(sourceRoot, path));
}

function stringLeaves(value: unknown, prefix = ''): Map<string, string> {
  const result = new Map<string, string>();
  if (typeof value === 'string') {
    result.set(prefix, value);
    return result;
  }
  if (Array.isArray(value)) {
    value.forEach((item, index) => {
      for (const [path, leaf] of stringLeaves(item, `${prefix}[${index}]`)) {
        result.set(path, leaf);
      }
    });
    return result;
  }
  if (!value || typeof value !== 'object') return result;
  for (const [key, child] of Object.entries(value)) {
    const path = prefix ? `${prefix}.${key}` : key;
    for (const [leafPath, leaf] of stringLeaves(child, path)) result.set(leafPath, leaf);
  }
  return result;
}

function placeholderSet(value: string): string[] {
  return [...value.matchAll(/\{[^{}]+\}/g)].map((match) => match[0]).sort();
}

function sectionValue(strings: Strings, section: LocaleSurfaceSection): unknown {
  return (strings as unknown as Record<string, unknown>)[section];
}

function portalFor(host: HTMLElement): HTMLElement | null {
  const instanceId = host.dataset.fcInstId;
  if (!instanceId) return null;
  return (
    Array.from(document.querySelectorAll<HTMLElement>('.fc-overlay-portal')).find(
      (portal) => portal.dataset.fcInstId === instanceId,
    ) ?? null
  );
}

function ownedRoots(host: HTMLElement): HTMLElement[] {
  const portal = portalFor(host);
  return [host, ...(portal ? [portal] : [])];
}

function isVisible(element: HTMLElement): boolean {
  if (!element.isConnected || element.hidden) return false;
  for (let node: HTMLElement | null = element; node; node = node.parentElement) {
    if (node.hidden || node.getAttribute('aria-hidden') === 'true') return false;
  }
  return true;
}

function visibleOverlays(host: HTMLElement): HTMLElement[] {
  const selector =
    '[role="dialog"], [role="alertdialog"], .fc-find, .fc-cmtnote, .fc-filter-dropdown, .fc-quick, .fc-objects';
  return ownedRoots(host).flatMap((root) =>
    Array.from(root.querySelectorAll<HTMLElement>(selector)).filter(isVisible),
  );
}

function closeVisibleOverlay(host: HTMLElement, instance: SpreadsheetInstance): void {
  instance.closeFindReplace();
  instance.closeWatchWindow();
  for (const overlay of visibleOverlays(host)) {
    overlay.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
  }
  document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
  // A shared shell handles Escape itself. Manual panels may only listen while
  // focused, so the second pass targets the currently focused ancestor.
  const active = document.activeElement;
  if (active instanceof HTMLElement) {
    active.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
  }
}

function firstVisibleOverlay(host: HTMLElement): HTMLElement {
  const overlay = visibleOverlays(host)[0];
  if (!overlay) throw new Error('expected a visible locale surface');
  return overlay;
}

function assertOverlayLabel(overlay: HTMLElement, expected: string): void {
  const aria = overlay.getAttribute('aria-label');
  if (aria === expected) return;
  expect(overlay.textContent ?? '').toContain(expected);
}

function toolbarButton(host: HTMLElement, command: string): HTMLButtonElement {
  const button = host.querySelector<HTMLButtonElement>(`[data-ribbon-command="${command}"]`);
  if (!button) throw new Error(`missing rendered ribbon command ${command}`);
  return button;
}

function clickToolbarTab(host: HTMLElement, tab: string): void {
  const button = host.querySelector<HTMLButtonElement>(`[data-ribbon-tab="${tab}"]`);
  if (!button) throw new Error(`missing rendered ribbon tab ${tab}`);
  button.click();
}

function closeToolbarTransientSurfaces(host: HTMLElement, instance: SpreadsheetInstance): void {
  instance.toolbar?.dropdownsApi?.closeAllDynamicRibbonDropdowns();
  closeVisibleOverlay(host, instance);
}

function openToolbarMenu(host: HTMLElement, route: LocaleSurfaceToolbarRoute): HTMLElement {
  const button = toolbarButton(host, route.command);
  const chevron = button.querySelector<SVGElement>('.fc-tb__rb-split-chevron');
  if (chevron) {
    // Route through the mounted toolbar's delegated click handler. Calling the
    // dropdown API directly would bypass the user's split-button wiring.
    chevron.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true }));
  } else {
    button.click();
  }
  const menu = route.menuId ? host.querySelector<HTMLElement>(`#${route.menuId}`) : null;
  if (!menu || menu.hidden) throw new Error(`ribbon menu did not open: ${route.id}`);
  return menu;
}

function clickToolbarMenuAction(
  menu: HTMLElement,
  route: LocaleSurfaceToolbarRoute,
): HTMLButtonElement {
  const attr = route.actionAttr;
  const action = route.action;
  if (!attr || action === undefined) throw new Error(`route has no menu action: ${route.id}`);
  const button = Array.from(menu.querySelectorAll<HTMLButtonElement>('button')).find(
    (candidate) => candidate.dataset[attr] === action,
  );
  if (!button) throw new Error(`missing ${attr}=${action} in ${route.id}`);
  if (button.disabled || button.getAttribute('aria-disabled') === 'true') {
    throw new Error(
      `disabled toolbar action ${route.id}: ${button.dataset.menuDisabledReason ?? ''}`,
    );
  }
  button.click();
  return button;
}

function openToolbarRoute(host: HTMLElement, route: LocaleSurfaceToolbarRoute): void {
  clickToolbarTab(host, route.tab);
  if (route.kind === 'command') {
    toolbarButton(host, route.command).click();
    return;
  }
  if (route.kind === 'dynamic-menu') {
    const menu = openToolbarMenu(host, route);
    clickToolbarMenuAction(menu, route);
    return;
  }
  const select = host.querySelector<HTMLElement>(`[data-ribbon-select="${route.command}"]`);
  if (!select) throw new Error(`missing ribbon select ${route.command}`);
  select.querySelector<HTMLButtonElement>('.fc-tb__rb-dd__btn')?.click();
  const option = select.querySelector<HTMLButtonElement>(
    `[role="option"][data-value="${route.controlValue ?? ''}"]`,
  );
  if (!option) throw new Error(`missing ribbon select option ${route.controlValue}: ${route.id}`);
  option.click();
}

function cancelToolbarDialog(host: HTMLElement): HTMLElement {
  const dialog = firstVisibleOverlay(host);
  const cancel = dialog.querySelector<HTMLButtonElement>(
    '.fc-fmtdlg__footer .fc-fmtdlg__btn:not(.fc-fmtdlg__btn--primary)',
  );
  if (!cancel) throw new Error('locale toolbar surface has no cancel button');
  cancel.click();
  return dialog;
}

const toolbarFixtureRange = { sheet: 0, r0: 0, c0: 0, r1: 4, c1: 2 } as const;

function seedToolbarFixture(mountedSheet: MountedStubSheet): void {
  const { workbook, instance } = mountedSheet;
  const headers = ['Item', 'Qty', 'Delimited'];
  for (const [col, value] of headers.entries()) {
    workbook.setText({ sheet: 0, row: 0, col }, value);
  }
  const rows = [
    ['banana', 7, 'a,b'],
    ['apple', 12, 'c,d'],
    ['cherry', 3, 'e,f'],
    ['date', 25, 'g,h'],
  ] as const;
  rows.forEach(([item, qty, delimited], row) => {
    workbook.setText({ sheet: 0, row: row + 1, col: 0 }, item);
    workbook.setNumber({ sheet: 0, row: row + 1, col: 1 }, qty);
    workbook.setText({ sheet: 0, row: row + 1, col: 2 }, delimited);
  });
  workbook.recalc();
  mutators.replaceCells(instance.store, workbook.cells(0));
  mutators.setRange(instance.store, toolbarFixtureRange);
  formatAsTableByStyleId(instance.store, toolbarFixtureRange, 'medium', undefined, 'banded', {
    showHeader: true,
    workbook,
  });
  mutators.setRange(instance.store, toolbarFixtureRange);
}

function toolbarFixtureSnapshot(mountedSheet: MountedStubSheet, includeTables = false): string {
  const { workbook } = mountedSheet;
  return JSON.stringify({
    cells: Array.from({ length: toolbarFixtureRange.r1 - toolbarFixtureRange.r0 + 1 }, (_, row) =>
      Array.from({ length: toolbarFixtureRange.c1 - toolbarFixtureRange.c0 + 1 }, (_, col) =>
        workbook.getValue({ sheet: 0, row, col }),
      ),
    ),
    tables: includeTables ? workbook.getTables() : [],
  });
}

function dictionaryLabel(strings: Strings, path: string): string {
  const value = path.split('.').reduce<unknown>((current, key) => {
    if (!current || typeof current !== 'object') return undefined;
    return (current as Record<string, unknown>)[key];
  }, strings);
  if (typeof value !== 'string' || value.length === 0) {
    throw new Error(`dictionary label ${path} is not a non-empty string`);
  }
  return value;
}

const PUBLIC_MODAL_OPENERS: readonly {
  readonly id: string;
  readonly section: LocaleSurfaceSection;
  readonly open: (instance: SpreadsheetInstance) => void;
  readonly titlePath: string;
}[] = [
  {
    id: 'formatDialog',
    section: 'formatDialog',
    open: (instance) => instance.openFormatDialog('number'),
    titlePath: 'formatDialog.title',
  },
  {
    id: 'conditionalDialog',
    section: 'conditionalDialog',
    open: (instance) => instance.openConditionalDialog(),
    titlePath: 'conditionalDialog.title',
  },
  {
    id: 'cfRulesDialog',
    section: 'cfRulesDialog',
    open: (instance) => instance.openCfRulesDialog(),
    titlePath: 'cfRulesDialog.title',
  },
  {
    id: 'namedRangeDialog',
    section: 'namedRangeDialog',
    open: (instance) => instance.openNamedRangeDialog(),
    titlePath: 'namedRangeDialog.title',
  },
  {
    id: 'iterativeDialog',
    section: 'iterativeDialog',
    open: (instance) => instance.openIterativeDialog(),
    titlePath: 'iterativeDialog.title',
  },
  {
    id: 'externalLinksDialog',
    section: 'externalLinksDialog',
    open: (instance) => instance.openExternalLinksDialog(),
    titlePath: 'externalLinksDialog.title',
  },
  {
    id: 'evaluateFormulaDialog',
    section: 'evaluateFormulaDialog',
    open: (instance) => instance.openEvaluateFormulaDialog(),
    titlePath: 'evaluateFormulaDialog.title',
  },
  {
    id: 'functionArguments',
    section: 'fxDialog',
    open: (instance) => instance.openFunctionArguments('SUM'),
    titlePath: 'fxDialog.title',
  },
  {
    id: 'hyperlinkDialog',
    section: 'hyperlinkDialog',
    open: (instance) => instance.openHyperlinkDialog(),
    titlePath: 'hyperlinkDialog.title',
  },
  {
    id: 'commentDialog',
    section: 'commentDialog',
    open: (instance) => instance.openCommentDialog({ mode: 'note' }),
    titlePath: 'commentDialog.title',
  },
  {
    id: 'findReplace',
    section: 'findReplace',
    open: (instance) => instance.openFindReplace('find'),
    titlePath: 'findReplace.title',
  },
  {
    id: 'pageSetup',
    section: 'pageSetup',
    open: (instance) => instance.openPageSetup('page'),
    titlePath: 'pageSetup.title',
  },
  {
    id: 'goToDialog',
    section: 'goToDialog',
    open: (instance) => instance.openGoTo(),
    titlePath: 'goToDialog.goToTitle',
  },
  {
    id: 'goToSpecial',
    section: 'goToDialog',
    open: (instance) => instance.openGoToSpecial(),
    titlePath: 'goToDialog.title',
  },
  {
    id: 'dataValidation',
    section: 'formatDialog',
    open: (instance) => instance.openDataValidationDialog(),
    titlePath: 'formatDialog.validationLegend',
  },
  {
    id: 'cellStylesGallery',
    section: 'cellStylesGallery',
    open: (instance) => instance.openCellStylesGallery(),
    titlePath: 'cellStylesGallery.title',
  },
];

const LIVE_PANEL_OPENERS: readonly {
  readonly id: string;
  readonly section: LocaleSurfaceSection;
  readonly open: (instance: SpreadsheetInstance) => void;
  readonly close: (instance: SpreadsheetInstance) => void;
  readonly selector: string;
  readonly titlePath: string;
}[] = [
  {
    id: 'watchPanel',
    section: 'watchPanel',
    open: (instance) => instance.openWatchWindow(),
    close: (instance) => instance.closeWatchWindow(),
    selector: '.fc-watch',
    titlePath: 'watchPanel.title',
  },
  {
    id: 'workbookObjects',
    section: 'workbookObjects',
    open: (instance) => instance.openWorkbookObjects(),
    close: (_instance) => {
      document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    },
    selector: '.fc-objects',
    titlePath: 'workbookObjects.title',
  },
];

let mounted: MountedStubSheet | undefined;

afterEach(() => {
  mounted?.dispose();
  mounted = undefined;
});

describe('seven-language locale surface manifest', () => {
  it('keeps all 52 dictionary sections complete and placeholder-compatible', () => {
    expect(LOCALE_SURFACE_LOCALES).toHaveLength(7);
    expect(Object.keys(dictionaries).sort()).toEqual([...LOCALE_SURFACE_LOCALES].sort());
    expect(Object.keys(dictionaries.en)).toHaveLength(52);
    expect(Object.keys(dictionaries.en).sort()).toEqual([...LOCALE_SURFACE_SECTIONS].sort());

    const baseline = stringLeaves(dictionaries.en);
    expect(baseline.size).toBeGreaterThan(2_000);
    for (const locale of LOCALE_SURFACE_LOCALES) {
      const leaves = stringLeaves(dictionaries[locale]);
      for (const [path, english] of baseline) {
        const translated = leaves.get(path);
        expect(translated, `${locale}:${path}`).toBeDefined();
        expect(translated, `${locale}:${path}`).not.toBe('');
        expect(placeholderSet(translated ?? ''), `${locale}:${path}`).toEqual(
          placeholderSet(english),
        );
      }
      for (const section of LOCALE_SURFACE_SECTIONS) {
        const value = sectionValue(dictionaries[locale], section);
        expect(value, `${locale}:${section}`).toBeTruthy();
      }
    }
  });

  it('classifies every shared-shell caller and manual locale surface in source', () => {
    const allFactories = [...LOCALE_SURFACE_FACTORIES];
    expect(new Set(allFactories.map((factory) => factory.id)).size).toBe(allFactories.length);
    expect(new Set(allFactories.map((factory) => factory.source)).size).toBeGreaterThan(30);
    for (const factory of allFactories) {
      const path = resolve(sourceRoot, factory.source);
      expect(statSync(path).isFile(), factory.id).toBe(true);
      expect(sourceText(factory.source), factory.id).toContain(factory.marker);
      expect(LOCALE_SURFACE_SECTIONS).toContain(factory.section);
      if (factory.runtime === 'unavailable') expect(factory.reason).toBeTruthy();
    }

    const discoveredShellCallers = sourceFiles(sourceRoot)
      .filter((path) => readFileSync(path, 'utf8').includes('createDialogShell'))
      .map(relativeSource)
      .filter((path) => path !== 'interact/dialog-shell.ts' && path !== 'toolbar/dialogs/shell.ts')
      .sort();
    const manifestShellCallers = LOCALE_SURFACE_SHELL_FACTORIES.map(
      (factory) => factory.source,
    ).sort();
    expect(discoveredShellCallers).toEqual(manifestShellCallers);
  });

  it('keeps every mounted toolbar route and unavailable browser boundary explicit', () => {
    const factoryIds = new Set(LOCALE_SURFACE_FACTORIES.map((factory) => factory.id));
    const routeIds = new Set(LOCALE_SURFACE_TOOLBAR_ROUTES.map((route) => route.id));
    expect(routeIds.size).toBe(LOCALE_SURFACE_TOOLBAR_ROUTES.length);
    expect(LOCALE_SURFACE_TOOLBAR_ROUTES.length).toBeGreaterThanOrEqual(15);
    for (const route of LOCALE_SURFACE_TOOLBAR_ROUTES) {
      expect(factoryIds, route.id).toContain(route.factoryId);
      expect(LOCALE_SURFACE_SECTIONS, route.id).toContain(route.section);
      expect(route.titlePath, route.id).toContain('.');
      if (route.kind === 'dynamic-menu') {
        expect(route.menuId, route.id).toBeTruthy();
        expect(route.actionAttr, route.id).toBeTruthy();
        expect(route.action, route.id).toBeTruthy();
      }
      if (route.kind === 'ribbon-select') expect(route.controlValue, route.id).toBeTruthy();
    }
    for (const boundary of LOCALE_SURFACE_TOOLBAR_UNAVAILABLE) {
      expect(LOCALE_SURFACE_SECTIONS, boundary.id).toContain(boundary.section);
      expect(boundary.reason, boundary.id).toMatch(/\S/);
    }
  });

  it.each(LOCALE_SURFACE_LOCALES)(
    'mounts the public locale surfaces and keeps chrome labels in %s',
    async (locale: LocaleSurfaceLocale) => {
      mounted = await mountStubSheet({
        locale,
        toolbar: true,
        ui: { features: { localeSwitcher: true } },
      });
      const { host, instance } = mounted;
      const strings = dictionaries[locale];

      expect(host.getAttribute('aria-label')).toBe(dictionaryLabel(strings, 'a11y.spreadsheet'));
      expect(host.querySelector('.fc-host__formulabar-tag')?.getAttribute('aria-label')).toBe(
        dictionaryLabel(strings, 'a11y.nameBox'),
      );
      expect(host.querySelector('.fc-host__formulabar-input')?.getAttribute('aria-label')).toBe(
        dictionaryLabel(strings, 'a11y.formulaBar'),
      );
      expect(
        host.querySelector('.fc-host__locale-switcher-select')?.getAttribute('aria-label'),
      ).toBe(dictionaryLabel(strings, 'localeSwitcher.label'));
      expect(host.querySelector('[data-ribbon-tab="home"]')?.textContent).toBe(
        dictionaryLabel(strings, 'ribbon.tabs.home'),
      );

      for (const opener of PUBLIC_MODAL_OPENERS) {
        closeVisibleOverlay(host, instance);
        expect(() => opener.open(instance), opener.id).not.toThrow();
        const overlays = visibleOverlays(host);
        expect(overlays.length, opener.id).toBeGreaterThan(0);
        const overlay = overlays[0] as HTMLElement;
        assertOverlayLabel(overlay, dictionaryLabel(strings, opener.titlePath));
      }

      // A true modal owns an editable draft boundary. The node, text and
      // active locale must survive a rejected transition. The same dialog is
      // then reopened after a successful transition to prove live strings are
      // resolved at runtime rather than only at mount time.
      closeVisibleOverlay(host, instance);
      instance.openFormatDialog('number');
      const modal = firstVisibleOverlay(host);
      const modalText = modal.textContent;
      const target = locale === 'ja' ? 'de' : locale === 'th' ? 'en' : 'ja';
      expect(instance.setLocale(target)).toBe(false);
      expect(instance.locale).toBe(locale);
      expect(modal.isConnected).toBe(true);
      expect(modal.textContent).toBe(modalText);
      closeVisibleOverlay(host, instance);
      expect(instance.setLocale(target)).toBe(true);
      expect(instance.locale).toBe(target);
      instance.openFormatDialog('number');
      const reopened = firstVisibleOverlay(host);
      assertOverlayLabel(reopened, dictionaryLabel(dictionaries[target], 'formatDialog.title'));

      for (const panel of LIVE_PANEL_OPENERS) {
        closeVisibleOverlay(host, instance);
        panel.open(instance);
        const live = host.querySelector<HTMLElement>(panel.selector);
        expect(live, panel.id).not.toBeNull();
        if (!live) continue;
        const before = live;
        const beforeText = before.textContent;
        const liveTarget = instance.locale === 'en' ? 'de' : 'en';
        expect(instance.setLocale(liveTarget)).toBe(true);
        expect(instance.locale).toBe(liveTarget);
        expect(before.isConnected).toBe(true);
        expect(before.hidden).toBe(false);
        expect(before.getAttribute('aria-label')).toBe(
          dictionaryLabel(dictionaries[liveTarget], panel.titlePath),
        );
        expect(before.textContent).not.toBe(beforeText);
        panel.close(instance);
      }
    },
    60_000,
  );

  it.each(LOCALE_SURFACE_LOCALES)(
    'drives mounted toolbar locale routes through the user click paths in %s',
    async (locale: LocaleSurfaceLocale) => {
      const toolbarWorkbook = await WorkbookHandle.createDefault({ locale });
      mounted = await mountStubSheet({
        locale,
        toolbar: true,
        ui: { features: { localeSwitcher: true } },
        workbook: toolbarWorkbook,
      });
      seedToolbarFixture(mounted);
      const { host, instance } = mounted;
      const strings = dictionaries[locale];

      for (const route of LOCALE_SURFACE_TOOLBAR_ROUTES) {
        closeToolbarTransientSurfaces(host, instance);
        if (route.id === 'toolbar.protection.unprotectSheet') {
          instance.setSheetProtected(true, 'toolbar-locale-password');
        }
        const trackTables =
          route.id === 'toolbar.formatAsTable' || route.id === 'toolbar.tableStyle';
        const beforeCells = toolbarFixtureSnapshot(mounted, trackTables);
        const beforeAllowedRanges = instance.store.getState().protection.allowedEditRanges;
        openToolbarRoute(host, route);
        await Promise.resolve();

        if (route.cancelable) {
          const overlay = firstVisibleOverlay(host);
          assertOverlayLabel(overlay, dictionaryLabel(strings, route.titlePath));
          const beforeText = overlay.textContent;
          const switched = locale === 'en' ? 'ja' : 'en';
          if (route.id === 'toolbar.sort.custom') {
            // A toolbar-created draft must own the same transition guard as a
            // public opener. Keep the node and draft text intact on rejection.
            expect(instance.setLocale(switched)).toBe(false);
            expect(instance.locale).toBe(locale);
            expect(overlay.isConnected).toBe(true);
            expect(overlay.textContent).toBe(beforeText);
          }
          cancelToolbarDialog(host);
          await Promise.resolve();
          expect(toolbarFixtureSnapshot(mounted, trackTables), route.id).toBe(beforeCells);
          expect(instance.store.getState().protection.allowedEditRanges).toEqual(
            beforeAllowedRanges,
          );
          if (route.id === 'toolbar.protection.unprotectSheet') {
            expect(instance.isSheetProtected()).toBe(true);
            instance.setSheetProtected(false);
          }
        } else if (route.id === 'toolbar.protection.protectWorkbook') {
          expect(isWorkbookStructureProtected(instance.store.getState())).toBe(true);
        } else if (route.id === 'toolbar.protection.unprotectWorkbook') {
          expect(isWorkbookStructureProtected(instance.store.getState())).toBe(false);
        }
      }

      // Switch after a toolbar-created dialog has been rejected, then reopen
      // through the same rendered route. This catches stale menu factories as
      // well as stale modal strings.
      const switched = locale === 'en' ? 'ja' : 'en';
      closeToolbarTransientSurfaces(host, instance);
      openToolbarRoute(host, LOCALE_SURFACE_TOOLBAR_ROUTES[0] as LocaleSurfaceToolbarRoute);
      expect(firstVisibleOverlay(host)).toBeTruthy();
      cancelToolbarDialog(host);
      expect(instance.setLocale(switched)).toBe(true);
      closeToolbarTransientSurfaces(host, instance);
      openToolbarRoute(host, LOCALE_SURFACE_TOOLBAR_ROUTES[0] as LocaleSurfaceToolbarRoute);
      assertOverlayLabel(
        firstVisibleOverlay(host),
        dictionaryLabel(dictionaries[switched], 'ribbonMenu.sortDialogTitle'),
      );
      cancelToolbarDialog(host);
      expect(instance.setLocale(locale)).toBe(true);
    },
    60_000,
  );
});
