import { execFileSync } from 'node:child_process';
import { existsSync, readFileSync } from 'node:fs';
import { dirname, resolve as resolvePath } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { MAC_AUTOMATION_GALLERY } from '../../../src/toolbar/ribbon/mac/automation-gallery.js';
import {
  GROUP_LABELS,
  labels,
  macRibbonLabelForCommand,
} from '../../../src/toolbar/ribbon/mac/labels.js';
import { MAC_BUNDLE_LABELS } from '../../../src/toolbar/ribbon/mac/labels-bundle.js';
import { MAC_LOCALES, type MacLocale } from '../../../src/toolbar/ribbon/mac/labels-data.js';
import { MAC_RIBBON_MENU_ITEMS } from '../../../src/toolbar/ribbon/mac/menu-manifest.js';
import { createMacRibbonMenuFactory } from '../../../src/toolbar/ribbon/mac/menus.js';
import { applyMenuVariant } from '../../../src/toolbar/ribbon/menus/mac-variants.js';
import { pivotStyleLabel } from '../../../src/toolbar/ribbon/pivot-style-gallery.js';
import { buildRibbonModel, CONTEXTUAL_RIBBON_TABS } from '../../../src/toolbar/ribbon-model.js';

const MAC_APP = '/Applications/Microsoft Excel.app';
const HERE = dirname(fileURLToPath(import.meta.url));
const GENERATOR = resolvePath(HERE, '../../../../../scripts/gen-mac-ribbon-labels.py');
const BUNDLE_FILE = resolvePath(HERE, '../../../src/toolbar/ribbon/mac/labels-bundle.ts');

const AUTOMATION_SOURCE = resolvePath(
  HERE,
  '../../../src/toolbar/ribbon/mac/automation-gallery.ts',
);

/** A Mac menu row carrying `rowKey` (`dataset:value`), worded for the Mac in `locale`. */
const menuRowText = (rowKey: string, locale: MacLocale): string => {
  const [attr = '', value = ''] = rowKey.split(':');
  const menu = document.createElement('div');
  const row = document.createElement('button');
  row.className = 'fc-tb__menu-item';
  row.dataset[attr] = value;
  menu.append(row);
  applyMenuVariant(menu, locale, { mac: true });
  return row.textContent ?? '';
};

const menuHeadingText = (menuId: string, locale: MacLocale): string => {
  const menu = document.createElement('div');
  menu.id = menuId;
  applyMenuVariant(menu, locale, { mac: true });
  return menu.firstElementChild?.textContent ?? '';
};

const homeCommand = (id: string, locale: MacLocale) =>
  buildRibbonModel(locale, { profile: 'excel365Mac' })
    .find((tab) => tab.id === 'home')
    ?.groups.flatMap((group) => group.commands)
    .find((command) => command.id === id);

/** The text a gallery's own row or caption is drawn with. */
const galleryText = (id: string, locale: MacLocale): string | undefined => {
  const menuId = Object.keys(MAC_RIBBON_MENU_ITEMS).find((key) =>
    MAC_RIBBON_MENU_ITEMS[key]?.some((item) => item.id === id),
  );
  const items = menuId ? (MAC_RIBBON_MENU_ITEMS[menuId] ?? []) : [];
  const menu = menuId ? createMacRibbonMenuFactory(locale)(menuId) : null;
  const item = items.find((entry) => entry.id === id);
  if (!menu || !item) return undefined;
  if (!item.heading) {
    return menu.querySelector(`[data-ribbon-command="${id}"]`)?.textContent ?? undefined;
  }
  const headings = items.filter((entry) => entry.heading);
  return menu.querySelectorAll('.fc-tb__menu-heading')[headings.indexOf(item)]?.textContent ?? '';
};

const PIVOT_STYLES = ['Light', 'Medium', 'Dark'].flatMap((family) =>
  Array.from({ length: 28 }, (_, i) => `PivotStyle${family}${i + 1}`),
);

/** The style name part of a built-in style's accessible name, after its colour. */
const pivotStyleName = (name: string, locale: MacLocale): string => {
  if (name === 'none') return pivotStyleLabel('', locale);
  const [, ...rest] = pivotStyleLabel(name, locale).split(', ');
  return rest.join(', ');
};

/** The colours the built-in styles' accessible names open with. */
const pivotColourWords = (locale: MacLocale): string[] =>
  PIVOT_STYLES.map((name) => pivotStyleLabel(name, locale).split(', ')[0] ?? '');

const contextualModels = new Map<MacLocale, ReturnType<typeof buildRibbonModel>>();

/** The label of a Mac contextual tab, or of one of its commands. */
const contextualText = (locale: MacLocale, tabId: string, commandId?: string) => {
  let model = contextualModels.get(locale);
  if (!model) {
    model = buildRibbonModel(locale, {
      profile: 'excel365Mac',
      contextualTabs: CONTEXTUAL_RIBBON_TABS,
    });
    contextualModels.set(locale, model);
  }
  const tab = model.find((candidate) => candidate.id === tabId);
  if (!commandId) return tab?.label;
  return tab?.groups.flatMap((group) => group.commands).find((command) => command.id === commandId)
    ?.label;
};

/** Where each table and chart tab wording is drawn: its tab, and the command when it is one. */
const OBJECT_TAB_SITES: Readonly<Record<string, readonly [string, string?]>> = {
  tableTab: ['tableDesign'],
  tableName: ['tableDesign', 'tableName'],
  summarizePivot: ['tableDesign', 'tableSummarizePivot'],
  removeDuplicates: ['tableDesign', 'mac.data.removeDuplicates'],
  convertRange: ['tableDesign', 'tableConvertRange'],
  insertSlicer: ['tableDesign', 'mac.insert.slicer'],
  refresh: ['tableDesign', 'tableRefresh'],
  headerRow: ['tableDesign', 'tableHeaderRow'],
  totalRow: ['tableDesign', 'tableTotalRow'],
  bandedRows: ['tableDesign', 'tableBandedRows'],
  firstColumn: ['tableDesign', 'tableFirstColumn'],
  lastColumn: ['tableDesign', 'tableLastColumn'],
  bandedColumns: ['tableDesign', 'tableBandedColumns'],
  filterButton: ['tableDesign', 'tableFilterButton'],
  quickStyles: ['tableDesign', 'tableQuickStyles'],
  addElement: ['chartDesign', 'chartAddElement'],
  quickLayout: ['chartDesign', 'chartQuickLayout'],
  changeColors: ['chartDesign', 'chartColors'],
  switchRowColumn: ['chartDesign', 'chartSwitchRowColumn'],
  selectData: ['chartDesign', 'chartSelectData'],
  changeType: ['chartDesign', 'chartChangeType'],
  moveChart: ['chartDesign', 'chartMove'],
  chartElement: ['chartFormat', 'chartElementSelect'],
  resetStyle: ['chartFormat', 'chartResetStyle'],
  shapes: ['chartFormat', 'chartShapes'],
  changeShape: ['chartFormat', 'chartChangeShape'],
  shapeFill: ['chartFormat', 'chartShapeFill'],
  shapeOutline: ['chartFormat', 'chartShapeOutline'],
  shapeEffects: ['chartFormat', 'chartShapeEffects'],
  textFill: ['chartFormat', 'chartTextFill'],
  textOutline: ['chartFormat', 'chartTextOutline'],
  textEffects: ['chartFormat', 'chartTextEffects'],
  bringForward: ['chartFormat', 'objectBringForward'],
  sendBackward: ['chartFormat', 'objectSendBackward'],
  selectionPane: ['chartFormat', 'objectSelectionPane'],
  arrangeObjects: ['chartFormat', 'chartArrangeObjects'],
  align: ['chartFormat', 'chartAlign'],
  group: ['chartFormat', 'chartGroup'],
  rotate: ['chartFormat', 'chartRotate'],
  lockAspect: ['chartFormat', 'chartLockAspect'],
  formatPane: ['chartFormat', 'chartFormatSelection'],
};

/** Where each PivotTable tab wording is drawn: its tab and command. */
const PIVOT_TAB_SITES: Readonly<Record<string, readonly [string, string]>> = {
  subtotals: ['pivotDesign', 'pivotSubtotals'],
  grandTotals: ['pivotDesign', 'pivotGrandTotals'],
  reportLayout: ['pivotDesign', 'pivotReportLayout'],
  blankRows: ['pivotDesign', 'pivotBlankRows'],
  name: ['pivotAnalyze', 'pivotName'],
  options: ['pivotAnalyze', 'pivotOptions'],
  activeField: ['pivotAnalyze', 'pivotActiveField'],
  fieldSettings: ['pivotAnalyze', 'pivotFieldSettings'],
  expandField: ['pivotAnalyze', 'pivotExpandField'],
  collapseField: ['pivotAnalyze', 'pivotCollapseField'],
  groupSelection: ['pivotAnalyze', 'pivotGroupSelection'],
  insertSlicer: ['pivotAnalyze', 'pivotInsertSlicer'],
  insertTimeline: ['pivotAnalyze', 'pivotInsertTimeline'],
  filterConnections: ['pivotAnalyze', 'pivotFilterConnections'],
  refresh: ['pivotAnalyze', 'pivotRefresh'],
  changeSource: ['pivotAnalyze', 'pivotChangeSource'],
  clear: ['pivotAnalyze', 'pivotClear'],
  select: ['pivotAnalyze', 'pivotSelect'],
  move: ['pivotAnalyze', 'pivotMove'],
  fieldsItemsSets: ['pivotAnalyze', 'pivotFieldsItemsSets'],
  pivotChart: ['pivotAnalyze', 'pivotChart'],
  fieldList: ['pivotAnalyze', 'pivotFieldList'],
  buttons: ['pivotAnalyze', 'pivotButtons'],
  headers: ['pivotAnalyze', 'pivotHeaders'],
  rowHeaders: ['pivotDesign', 'pivotRowHeaders'],
  colHeaders: ['pivotDesign', 'pivotColHeaders'],
  rowStripes: ['pivotDesign', 'pivotRowStripes'],
  colStripes: ['pivotDesign', 'pivotColStripes'],
  quickStyles: ['pivotDesign', 'pivotQuickStyles'],
};

const resolve = (namespace: string, name: string, locale: MacLocale) => {
  if (namespace === 'labels') return labels[name]?.[locale];
  if (namespace === 'groups') return GROUP_LABELS[name]?.[locale];
  if (namespace === 'menu') return macRibbonLabelForCommand(name, locale);
  if (namespace === 'row') return menuRowText(name, locale);
  if (namespace === 'heading') return menuHeadingText(name, locale);
  if (namespace === 'home') return homeCommand(name, locale)?.label;
  if (namespace === 'home-help') return homeCommand(name, locale)?.title;
  if (namespace === 'pivot-style') return pivotStyleName(name, locale);
  if (namespace === 'gallery') return galleryText(name, locale);
  if (namespace === 'pivot-tab') {
    const [tabId = '', commandId] = PIVOT_TAB_SITES[name] ?? [];
    return contextualText(locale, tabId, commandId);
  }
  if (namespace === 'object-tab') {
    const [tabId = '', commandId] = OBJECT_TAB_SITES[name] ?? [];
    return contextualText(locale, tabId, commandId);
  }
  const script = MAC_AUTOMATION_GALLERY.find((entry) => entry.id === name);
  if (namespace === 'automation') return script?.labels[locale];
  if (namespace === 'automation-detail') return script?.details[locale];
  throw new Error(`Unknown bundle label namespace: ${namespace}`);
};

describe('generated Mac ribbon labels', () => {
  const entries = Object.entries(MAC_BUNDLE_LABELS);

  it('is not empty', () => {
    expect(entries.length).toBeGreaterThan(0);
  });

  it('never names the desktop product', () => {
    const named = entries.flatMap(([key, bundle]) =>
      Object.entries(bundle)
        .filter(([, text]) => /Excel/.test(text))
        .map(([locale]) => `${key} (${locale})`),
    );
    expect(named).toEqual([]);
  });

  it.each(entries)('%s resolves to the bundle label in every locale', (key, bundle) => {
    const split = key.indexOf(':');
    const namespace = key.slice(0, split);
    const name = key.slice(split + 1);
    if (namespace === 'colour') {
      for (const locale of MAC_LOCALES) {
        expect(pivotColourWords(locale), `${key} (${locale})`).toContain(bundle[locale]);
      }
      return;
    }
    if (namespace === 'automation-ui') {
      // Dialog text the gallery reads by key; its rendering is covered by the gallery tests.
      expect(readFileSync(AUTOMATION_SOURCE, 'utf8')).toContain(`'${key}'`);
      return;
    }
    for (const locale of MAC_LOCALES) {
      expect(resolve(namespace, name, locale), `${key} (${locale})`).toBe(bundle[locale]);
    }
  });

  it('cleans accelerator markers the way the generator documents', () => {
    expect(() =>
      execFileSync('python3', ['-I', '-m', 'doctest', GENERATOR], { stdio: 'pipe' }),
    ).not.toThrow();
  });

  it.skipIf(!existsSync(MAC_APP))('matches what the generator produces from the app bundle', () => {
    expect(() =>
      execFileSync('python3', ['-I', GENERATOR, '--check', BUNDLE_FILE], { stdio: 'pipe' }),
    ).not.toThrow();
  });
});
