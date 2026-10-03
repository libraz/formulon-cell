import { canExecuteBuiltIn } from '../../../commands/built-in-command-policy.js';
import {
  type CatalogFunctionCategory,
  type FunctionCatalogReader,
  isFunctionUnavailableForInsertion,
  supportedFunctionNames,
} from '../../../commands/function-categories.js';
import { getRecentFunctions } from '../../../commands/function-history.js';
import type { SpreadsheetStore } from '../../../store/store.js';
import { focusMenuItem, prepareMenu, projectDisabledState } from '../../menu-a11y.js';
import { createRibbonButton } from '../button.js';
import {
  createSubmenu,
  menuSubmenuTrigger,
  type VisualTileIconSlug,
  visualMenuGrid,
  visualMenuTile,
} from '../menus/general.js';
import {
  MAC_FORMULAS_MORE_MENU_ID,
  MAC_RIBBON_MENU_ITEMS,
  type MacRibbonMenuItem,
  macRibbonDisabledReason,
  macRibbonLabelForCommand,
  macRibbonMenuIdForCommand,
} from './model.js';

export type MacRibbonMenuFactory = (commandId: string) => HTMLDivElement | null;

type LiveFunctionNames = ReadonlySet<string> | readonly string[] | null | undefined;

export interface MacFunctionAvailabilityContext {
  /** The live workbook reader, supplied only when functionNames() was live. */
  reader?: FunctionCatalogReader | null;
  /** Localized reason shown on intrinsic unavailable leaves. */
  unavailableReason?: string;
}

type ReaderAvailabilityCache = {
  id: number;
  byLocale: Map<0 | 1, Map<string, number | undefined>>;
};

const readerAvailabilityCache = new WeakMap<object, ReaderAvailabilityCache>();
let nextReaderAvailabilityId = 1;

/** Fallback tooltip for a function leaf the live engine cannot insert. */
export const defaultFunctionUnavailableReason = (lang: 'en' | 'ja'): string =>
  lang === 'ja'
    ? 'この関数は現在の計算エンジンでは使用できません。'
    : 'This function is unavailable in the current calculation engine.';

const readerCacheFor = (reader: FunctionCatalogReader): ReaderAvailabilityCache => {
  const key = reader as object;
  const existing = readerAvailabilityCache.get(key);
  if (existing) return existing;
  const created: ReaderAvailabilityCache = {
    id: nextReaderAvailabilityId++,
    byLocale: new Map(),
  };
  readerAvailabilityCache.set(key, created);
  return created;
};

const readerIdFor = (reader: FunctionCatalogReader | null | undefined): string =>
  reader ? String(readerCacheFor(reader).id) : 'none';

const functionAvailabilityFor = (
  reader: FunctionCatalogReader | null | undefined,
  lang: 'en' | 'ja',
  name: string,
): number | undefined => {
  if (!reader?.functionMetadata) return undefined;
  const cache = readerCacheFor(reader);
  const locale: 0 | 1 = lang === 'ja' ? 1 : 0;
  let byName = cache.byLocale.get(locale);
  if (!byName) {
    byName = new Map();
    cache.byLocale.set(locale, byName);
  }
  if (byName.has(name)) return byName.get(name);
  const value = reader.functionMetadata(name, locale)?.availability;
  byName.set(name, value);
  return value;
};

const functionAvailabilitySignature = (
  context: MacFunctionAvailabilityContext | undefined,
  lang: 'en' | 'ja',
): string => `${readerIdFor(context?.reader)}:${lang}`;

const projectFunctionButtonState = (
  button: HTMLButtonElement,
  name: string,
  store: SpreadsheetStore,
  lang: 'en' | 'ja',
  context: MacFunctionAvailabilityContext | undefined,
): void => {
  const availability = functionAvailabilityFor(context?.reader, lang, name);
  const unavailable = isFunctionUnavailableForInsertion(availability);
  const reason = context?.unavailableReason ?? defaultFunctionUnavailableReason(lang);
  const decision = canExecuteBuiltIn(store, `mac.function.${name}`, 'ribbon');
  const disabledReason = unavailable
    ? reason
    : decision.allowed
      ? null
      : (decision.reason ?? decision.code);
  projectDisabledState(button, unavailable || !decision.allowed, disabledReason, {
    datasetKey: 'ribbonDisabledReason',
    titlePrefix: name,
  });
  if (unavailable) {
    button.dataset.functionUnavailable = 'true';
    button.setAttribute('aria-description', reason);
    button.dataset.ribbonDisabledReason = reason;
    button.title = `${name}\n${reason}`;
  } else {
    delete button.dataset.functionUnavailable;
  }
};

const functionCategoryMenuIds: Readonly<Record<string, CatalogFunctionCategory>> = {
  'mac.formulas.financial': 'financial',
  'mac.formulas.logical': 'logical',
  'mac.formulas.text': 'text',
  'mac.formulas.dateTime': 'datetime',
  'mac.formulas.lookup': 'lookup',
  'mac.formulas.math': 'math',
};

type MacMoreFunctionCategory = {
  category: CatalogFunctionCategory;
  label: { en: string; ja: string };
};

const MORE_FUNCTION_CATEGORIES: readonly MacMoreFunctionCategory[] = [
  { category: 'statistical', label: { en: 'Statistical', ja: '統計' } },
  { category: 'engineering', label: { en: 'Engineering', ja: 'エンジニアリング' } },
  { category: 'cube', label: { en: 'Cube', ja: 'キューブ' } },
  { category: 'information', label: { en: 'Information', ja: '情報' } },
  { category: 'compatibility', label: { en: 'Compatibility', ja: '互換性' } },
  { category: 'web', label: { en: 'Web', ja: 'Web' } },
];

const moreFunctionPanelId = (category: CatalogFunctionCategory): string =>
  `menu-mac-formulas-more-${category}`;

const liveNameSet = (names: LiveFunctionNames): ReadonlySet<string> | undefined => {
  if (names === null || names === undefined) return undefined;
  return names instanceof Set ? names : new Set(names);
};

const liveNameSignature = (names: ReadonlySet<string> | undefined): string =>
  names === undefined ? 'static' : [...names].sort().join(',');

const menuLabel = (commandId: string, lang: 'ja' | 'en'): string =>
  macRibbonLabelForCommand(commandId, lang);

const CHART_ICONS: Readonly<Record<string, VisualTileIconSlug>> = {
  'mac.insert.chartColumn': 'chart-column',
  'mac.insert.chartBar': 'chart-bar',
  'mac.insert.chartLine': 'chart-line',
  'mac.insert.chartArea': 'chart-area',
  'mac.insert.chartPie': 'chart-pie',
  'mac.insert.chartScatter': 'chart-scatter',
};
const SHAPE_ICONS: Readonly<Record<string, VisualTileIconSlug>> = {
  shapeLine: 'shape-line',
  shapeArrow: 'shape-arrow',
  shapeRectangle: 'shape-rectangle',
  shapeRoundedRectangle: 'shape-rounded-rectangle',
  shapeOval: 'shape-oval',
  shapeTriangle: 'shape-triangle',
  shapeDiamond: 'shape-diamond',
};

const menuItemButton = (
  item: MacRibbonMenuItem,
  lang: 'ja' | 'en',
  visual = false,
): HTMLButtonElement => {
  const label = item.labelJa && lang === 'ja' ? item.labelJa : item.label;
  const icon = CHART_ICONS[item.id] ?? SHAPE_ICONS[item.icon ?? ''];
  const button =
    visual && icon
      ? visualMenuTile({
          label,
          attr: 'ribbonCommand',
          value: item.id,
          icon,
        })
      : createRibbonButton({
          className: 'fc-tb__menu-item fc-tb__menu-item--mac',
          role: 'menuitem',
          title: label,
          ariaLabel: label,
          dataset: { ribbonCommand: item.id },
          text: label,
        });
  projectDisabledState(
    button,
    !!item.disabled,
    macRibbonDisabledReason(item.disabledReason, lang) ?? null,
    {
      datasetKey: 'ribbonDisabledReason',
      titlePrefix: label,
    },
  );
  return button;
};

const createMoreFunctionsMenu = (lang: 'ja' | 'en'): HTMLDivElement => {
  const menu = document.createElement('div');
  menu.className = 'fc-tb__menu fc-tb__menu--mac fc-tb__menu--mac-more';
  menu.id = MAC_FORMULAS_MORE_MENU_ID;
  menu.hidden = true;
  prepareMenu(menu, macRibbonLabelForCommand('mac.formulas.more', lang));

  for (const { category, label } of MORE_FUNCTION_CATEGORIES) {
    const localizedLabel = label[lang];
    const panelId = moreFunctionPanelId(category);
    const trigger = createRibbonButton({
      className: 'fc-tb__menu-item fc-tb__menu-item--mac',
      role: 'menuitem',
      title: localizedLabel,
      ariaLabel: localizedLabel,
      text: localizedLabel,
    });
    menuSubmenuTrigger(trigger, { functionCategorySubmenu: category }, { controlsId: panelId });
    const panel = createSubmenu({
      id: panelId,
      className: 'fc-tb__submenu fc-tb__submenu--mac-function-category',
      label: localizedLabel,
      dataset: { functionCategoryPanel: category },
    });
    panel.append(
      ...supportedFunctionNames(category).map((name) =>
        menuItemButton({ id: `mac.function.${name}`, label: name }, lang),
      ),
    );
    prepareMenu(panel, localizedLabel);
    menu.append(trigger, panel);
  }
  prepareMenu(menu, macRibbonLabelForCommand('mac.formulas.more', lang));
  return menu;
};

/** Creates the static, data-command based menus used by the Excel 365 Mac
 *  profile. Menu leaves deliberately use the same command attribute as top
 *  level buttons so toolbar delegation, keyboard focus restoration, and
 *  restricted-embedding checks remain shared. */
export const createMacRibbonMenuFactory = (lang: 'ja' | 'en'): MacRibbonMenuFactory => {
  return (commandId: string): HTMLDivElement | null => {
    if (commandId === 'mac.formulas.more') return createMoreFunctionsMenu(lang);
    const items = MAC_RIBBON_MENU_ITEMS[commandId];
    const id = macRibbonMenuIdForCommand(commandId);
    if (!items || !id) return null;
    const menu = document.createElement('div');
    menu.className = 'fc-tb__menu fc-tb__menu--mac';
    menu.id = id;
    menu.hidden = true;
    prepareMenu(menu, menuLabel(commandId, lang));
    const shapes = commandId === 'mac.insert.shapes';
    const charts =
      commandId === 'mac.insert.recommendedChart' || commandId.startsWith('mac.insert.chart');
    if (shapes || charts) {
      menu.classList.add(
        'fc-tb__menu--visual',
        shapes ? 'fc-tb__menu--shapes' : 'fc-tb__menu--charts',
      );
      const grid = visualMenuGrid(
        shapes ? 'fc-tb__visual-grid--shapes' : 'fc-tb__visual-grid--charts',
      );
      for (const item of items) {
        const button = menuItemButton(item, lang, true);
        if (button.classList.contains('fc-tb__visual-tile')) grid.appendChild(button);
        else menu.appendChild(button);
      }
      menu.prepend(grid);
    } else {
      for (const item of items) menu.appendChild(menuItemButton(item, lang));
    }
    return menu;
  };
};

export const createMacRibbonMenus = (
  lang: 'ja' | 'en',
): Readonly<Record<string, MacRibbonMenuFactory>> => {
  const factory = createMacRibbonMenuFactory(lang);
  return Object.fromEntries(
    Object.keys(MAC_RIBBON_MENU_ITEMS).map((commandId) => [commandId, factory]),
  );
};

/** Update the session MRU without replacing an unchanged menu or losing its keyboard focus. */
export const projectMacRecentMenu = (
  host: HTMLElement,
  store: SpreadsheetStore,
  lang: 'en' | 'ja',
  liveNames?: LiveFunctionNames,
  availabilityContext?: MacFunctionAvailabilityContext,
): void => {
  const menu = host.querySelector<HTMLElement>('#menu-mac-formulas-recent');
  if (!menu) return;
  const names = getRecentFunctions(store, liveNameSet(liveNames));
  const signature = `${lang}:${liveNameSignature(liveNameSet(liveNames))}:${functionAvailabilitySignature(availabilityContext, lang)}:${names.join(',')}`;
  const rebuild = menu.dataset.recentFunctions !== signature;
  const focused =
    rebuild && menu.contains(menu.ownerDocument.activeElement)
      ? (menu.ownerDocument.activeElement as HTMLElement).dataset.ribbonCommand
      : undefined;
  const items: MacRibbonMenuItem[] = [
    ...names.map((name) => ({ id: `mac.function.${name}`, label: name })),
    ...(MAC_RIBBON_MENU_ITEMS['mac.formulas.recent'] ?? []),
  ];
  if (rebuild) menu.replaceChildren(...items.map((item) => menuItemButton(item, lang)));
  const buttons = Array.from(menu.querySelectorAll<HTMLButtonElement>('[data-ribbon-command]'));
  for (const button of buttons) {
    const name = button.dataset.ribbonCommand?.replace(/^mac\.function\./, '');
    if (name && button.dataset.ribbonCommand?.startsWith('mac.function.')) {
      projectFunctionButtonState(button, name, store, lang, availabilityContext);
      continue;
    }
    const decision = canExecuteBuiltIn(store, button.dataset.ribbonCommand ?? '', 'ribbon');
    projectDisabledState(
      button,
      !decision.allowed,
      decision.allowed ? null : (decision.reason ?? decision.code ?? null),
      {
        datasetKey: 'ribbonDisabledReason',
        titlePrefix: button.textContent ?? '',
      },
    );
  }
  menu.dataset.recentFunctions = signature;
  if (focused) {
    const next =
      Array.from(menu.querySelectorAll<HTMLButtonElement>('[data-ribbon-command]')).find(
        (button) => button.dataset.ribbonCommand === focused,
      ) ??
      menu.querySelector<HTMLButtonElement>('[data-ribbon-command="mac.formulas.category.recent"]');
    next?.focus({ preventScroll: true });
  }
};

const projectFunctionCategoryMenu = (
  menu: HTMLElement,
  category: CatalogFunctionCategory,
  store: SpreadsheetStore,
  lang: 'en' | 'ja',
  liveSet: ReadonlySet<string> | undefined,
  liveSignature: string,
  availabilityContext: MacFunctionAvailabilityContext | undefined,
): void => {
  const names = supportedFunctionNames(category, liveSet);
  const signature = `${lang}:${category}:${liveSignature}:${functionAvailabilitySignature(availabilityContext, lang)}:${names.join(',')}`;
  const rebuild = menu.dataset.functionNames !== signature;
  const focused =
    rebuild && menu.contains(menu.ownerDocument.activeElement)
      ? (menu.ownerDocument.activeElement as HTMLElement).dataset.ribbonCommand
      : undefined;
  const items = names.map(
    (name): MacRibbonMenuItem => ({
      id: `mac.function.${name}`,
      label: name,
    }),
  );
  if (rebuild) {
    menu.replaceChildren(...items.map((item) => menuItemButton(item, lang)));
    prepareMenu(menu, menu.getAttribute('aria-label') ?? undefined);
  }
  for (const button of menu.querySelectorAll<HTMLButtonElement>('[data-ribbon-command]')) {
    const command = button.dataset.ribbonCommand ?? '';
    const name = command.startsWith('mac.function.') ? command.slice('mac.function.'.length) : '';
    if (name) projectFunctionButtonState(button, name, store, lang, availabilityContext);
  }
  menu.dataset.functionNames = signature;
  if (!focused) return;
  const next = menu.querySelector<HTMLButtonElement>(`[data-ribbon-command="${focused}"]`);
  if (next) next.focus({ preventScroll: true });
  else focusMenuItem(menu);
};

/** Re-projects the six curated Function Library menus and the six native More
 * submenus against the workbook's recognized names. Static membership remains
 * the fallback when the engine cannot enumerate its catalog; an empty live set
 * is authoritative and leaves the category panel empty. */
export const projectMacFunctionCategoryMenus = (
  host: HTMLElement,
  store: SpreadsheetStore,
  lang: 'en' | 'ja',
  liveNames?: LiveFunctionNames,
  availabilityContext?: MacFunctionAvailabilityContext,
): void => {
  const liveSet = liveNameSet(liveNames);
  const liveSignature = liveNameSignature(liveSet);
  for (const [commandId, category] of Object.entries(functionCategoryMenuIds)) {
    const menuId = macRibbonMenuIdForCommand(commandId);
    const menu = host.querySelector<HTMLElement>(`#${menuId}`);
    if (menu)
      projectFunctionCategoryMenu(
        menu,
        category,
        store,
        lang,
        liveSet,
        liveSignature,
        availabilityContext,
      );
  }
  const moreMenu = host.querySelector<HTMLElement>(`#${MAC_FORMULAS_MORE_MENU_ID}`);
  if (!moreMenu) return;
  for (const { category } of MORE_FUNCTION_CATEGORIES) {
    const panel = moreMenu.querySelector<HTMLElement>(
      `[data-function-category-panel="${category}"]`,
    );
    if (panel)
      projectFunctionCategoryMenu(
        panel,
        category,
        store,
        lang,
        liveSet,
        liveSignature,
        availabilityContext,
      );
  }
};
