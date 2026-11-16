import {
  canExecuteBuiltIn,
  canExecuteBuiltInRequest,
  macFunctionCommandId,
  macFunctionNameOf,
} from '../../../commands/built-in-command-policy.js';
import {
  type CatalogFunctionCategory,
  type CatalogLocaleOrdinal,
  type FunctionCatalogReader,
  isFunctionUnavailableForInsertion,
  supportedFunctionNames,
} from '../../../commands/function-categories.js';
import { getRecentFunctions } from '../../../commands/function-history.js';
import { interactionControllerFor } from '../../../commands/interaction-controller.js';
import { localizedDecisionReason } from '../../../commands/permission-reasons.js';
import { dictionaries, type Locale } from '../../../i18n/strings.js';
import { catalogLocaleOrdinal, catalogLocaleTag } from '../../../interact/function-catalog-text.js';
import type { SpreadsheetInstance } from '../../../mount/types.js';
import type { SpreadsheetStore } from '../../../store/store.js';
import {
  focusMenuItem,
  prepareMenu,
  projectDisabledState,
  RIBBON_DISABLED_REASON_KEY,
} from '../../menu-a11y.js';
import { createRibbonButton } from '../button.js';
import {
  createSubmenu,
  menuSectionHeader,
  menuSeparator,
  menuSubmenuTrigger,
  type VisualTileIconSlug,
  visualMenuGrid,
  visualMenuTile,
} from '../menus/general.js';
import { macFunctionCategoryLabel } from './labels.js';
import { localizedFunctionName, MAC_CONTEXT_LABELS } from './labels-data.js';
import {
  MAC_FORMULAS_MORE_MENU_ID,
  MAC_RIBBON_MENU_ITEMS,
  type MacRibbonMenuItem,
  macRibbonDisabledReason,
  macRibbonLabelForCommand,
  macRibbonMenuIdForCommand,
} from './model.js';

export type MacRibbonMenuFactory = (commandId: string) => HTMLDivElement | null;

export type LiveFunctionNames = ReadonlySet<string> | readonly string[] | null | undefined;

export interface MacFunctionAvailabilityContext {
  /** The live workbook reader, supplied only when functionNames() was live. */
  reader?: LocalizedFunctionCatalogReader | null;
  /** Localized reason shown on intrinsic unavailable leaves. */
  unavailableReason?: string;
}

type LocalizedFunctionCatalogReader = FunctionCatalogReader & {
  canonicalizeFunctionName?: (name: string, locale: CatalogLocaleOrdinal) => string | null;
  localizeFunctionName?: (name: string, locale: CatalogLocaleOrdinal) => string | null;
};

type ReaderAvailabilityCache = {
  id: number;
  /** Cache by BCP-47 tag so zh/ko/th do not collapse into the en ordinal. */
  byLocale: Map<string, Map<string, number | undefined>>;
  displayByLocale: Map<string, Map<string, string | undefined>>;
};

const readerAvailabilityCache = new WeakMap<object, ReaderAvailabilityCache>();
let nextReaderAvailabilityId = 1;

/** Fallback tooltip for a function leaf the live engine cannot insert. */
export const defaultFunctionUnavailableReason = (lang: Locale): string =>
  MAC_CONTEXT_LABELS.functionUnavailable[lang];

const readerCacheFor = (reader: FunctionCatalogReader): ReaderAvailabilityCache => {
  const key = reader as object;
  const existing = readerAvailabilityCache.get(key);
  if (existing) return existing;
  const created: ReaderAvailabilityCache = {
    id: nextReaderAvailabilityId++,
    byLocale: new Map(),
    displayByLocale: new Map(),
  };
  readerAvailabilityCache.set(key, created);
  return created;
};

const readerIdFor = (reader: FunctionCatalogReader | null | undefined): string =>
  reader ? String(readerCacheFor(reader).id) : 'none';

const functionAvailabilityFor = (
  reader: FunctionCatalogReader | null | undefined,
  lang: Locale,
  name: string,
): number | undefined => {
  if (!reader?.functionMetadata && !reader?.functionMetadataForLocale) return undefined;
  const cache = readerCacheFor(reader);
  const localeTag = catalogLocaleTag(lang);
  const localeOrdinal = catalogLocaleOrdinal(lang);
  let byName = cache.byLocale.get(localeTag);
  if (!byName) {
    byName = new Map();
    cache.byLocale.set(localeTag, byName);
  }
  if (byName.has(name)) return byName.get(name);
  const metadata =
    reader.functionMetadataForLocale?.(name, localeTag) ??
    reader.functionMetadata?.(name, localeOrdinal) ??
    null;
  const value = metadata?.availability;
  byName.set(name, value);
  return value;
};

const functionAvailabilitySignature = (
  context: MacFunctionAvailabilityContext | undefined,
  lang: Locale,
): string =>
  `${readerIdFor(context?.reader)}:${lang}:${context?.unavailableReason ?? defaultFunctionUnavailableReason(lang)}`;

const projectFunctionButtonState = (
  button: HTMLButtonElement,
  name: string,
  store: SpreadsheetStore,
  lang: Locale,
  context: MacFunctionAvailabilityContext | undefined,
): void => {
  const availability = functionAvailabilityFor(context?.reader, lang, name);
  const displayName = functionDisplayName(context?.reader, lang, name);
  const unavailable = isFunctionUnavailableForInsertion(availability);
  const reason = context?.unavailableReason ?? defaultFunctionUnavailableReason(lang);
  const decision = canExecuteBuiltIn(store, macFunctionCommandId(name), 'ribbon');
  const disabledReason = unavailable
    ? reason
    : decision.allowed
      ? null
      : localizedDecisionReason(decision, dictionaries[lang]);
  projectDisabledState(button, unavailable || !decision.allowed, disabledReason, {
    datasetKey: RIBBON_DISABLED_REASON_KEY,
    titlePrefix: displayName,
  });
  if (unavailable) {
    button.dataset.functionUnavailable = 'true';
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
};

const MORE_FUNCTION_CATEGORIES: readonly MacMoreFunctionCategory[] = [
  { category: 'statistical' },
  { category: 'engineering' },
  { category: 'cube' },
  { category: 'information' },
  { category: 'compatibility' },
  { category: 'web' },
];

const moreFunctionPanelId = (category: CatalogFunctionCategory): string =>
  `menu-mac-formulas-more-${category}`;

export const liveNameSet = (names: LiveFunctionNames): ReadonlySet<string> | undefined => {
  if (names === null || names === undefined) return undefined;
  return names instanceof Set ? names : new Set(names);
};

const liveNameSignature = (names: ReadonlySet<string> | undefined): string =>
  names === undefined ? 'static' : [...names].sort().join(',');

type PreparedLiveFunctionNames = {
  set: ReadonlySet<string> | undefined;
  signature: string;
};

const prepareLiveFunctionNames = (
  names: LiveFunctionNames,
  lang: Locale,
  reader: LocalizedFunctionCatalogReader | null | undefined,
): PreparedLiveFunctionNames => {
  const set = canonicalLiveNameSet(liveNameSet(names), lang, reader);
  return { set, signature: liveNameSignature(set) };
};

const canonicalFunctionName = (
  reader: LocalizedFunctionCatalogReader | null | undefined,
  lang: Locale,
  name: string,
): string => {
  const normalized = name.trim().toUpperCase();
  const canonical = reader?.canonicalizeFunctionName?.(normalized, catalogLocaleOrdinal(lang));
  return canonical?.trim().toUpperCase() || normalized;
};

const canonicalLiveNameSet = (
  names: ReadonlySet<string> | undefined,
  lang: Locale,
  reader: LocalizedFunctionCatalogReader | null | undefined,
): ReadonlySet<string> | undefined => {
  if (names === undefined) return undefined;
  return new Set([...names].map((name) => canonicalFunctionName(reader, lang, name)));
};

const functionDisplayName = (
  reader: LocalizedFunctionCatalogReader | null | undefined,
  lang: Locale,
  canonicalName: string,
): string => {
  const localeTag = catalogLocaleTag(lang);
  const localeOrdinal = catalogLocaleOrdinal(lang);
  if (reader?.functionMetadata || reader?.functionMetadataForLocale) {
    const cache = readerCacheFor(reader);
    let byName = cache.displayByLocale.get(localeTag);
    if (!byName) {
      byName = new Map();
      cache.displayByLocale.set(localeTag, byName);
    }
    if (byName.has(canonicalName)) return byName.get(canonicalName) ?? canonicalName;
    const metadata =
      reader.functionMetadataForLocale?.(canonicalName, localeTag) ??
      reader.functionMetadata?.(canonicalName, localeOrdinal) ??
      null;
    const metadataName = metadata?.localizedName?.trim();
    if (metadataName) {
      byName.set(canonicalName, metadataName);
      return metadataName;
    }
    const localized = reader.localizeFunctionName?.(canonicalName, localeOrdinal)?.trim();
    if (localized) {
      byName.set(canonicalName, localized);
      return localized;
    }
    byName.set(canonicalName, undefined);
  }
  return localizedFunctionName(canonicalName, lang);
};

const menuLabel = (commandId: string, lang: Locale): string =>
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
  lang: Locale,
  visual = false,
  displayName?: string,
): HTMLButtonElement => {
  const functionName = macFunctionNameOf(item.id);
  const localized =
    displayName ??
    (functionName
      ? functionDisplayName(undefined, lang, functionName)
      : macRibbonLabelForCommand(item.id, lang));
  const label =
    item.own || localized === item.id
      ? item.labelJa && lang === 'ja'
        ? item.labelJa
        : item.label
      : localized;
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
  if (item.check) {
    button.setAttribute('role', item.check === 'radio' ? 'menuitemradio' : 'menuitemcheckbox');
    button.setAttribute('aria-checked', 'false');
  }
  projectDisabledState(
    button,
    !!item.disabled,
    macRibbonDisabledReason(item.disabledReason, lang) ?? null,
    {
      datasetKey: RIBBON_DISABLED_REASON_KEY,
      titlePrefix: label,
    },
  );
  return button;
};

/** A submenu row and its panel. The row carries no command, so a click opens the panel. */
const menuSubmenuRow = (
  item: MacRibbonMenuItem,
  menuId: string,
  lang: Locale,
): [HTMLButtonElement, HTMLDivElement] => {
  const label = macRibbonLabelForCommand(item.id, lang);
  const key = (item.id.split('.').at(-1) ?? item.id).replace(
    /[A-Z]/g,
    (c) => `-${c.toLowerCase()}`,
  );
  const panelId = `${menuId}-${key}`;
  const trigger = menuSubmenuTrigger(
    createRibbonButton({
      className: 'fc-tb__menu-item fc-tb__menu-item--mac',
      role: 'menuitem',
      title: label,
      ariaLabel: label,
      text: label,
    }),
    { formatSubmenu: item.id },
    { controlsId: panelId },
  );
  const panel = createSubmenu({
    id: panelId,
    className: 'fc-tb__submenu fc-tb__submenu--format',
    label,
    dataset: { formatPanel: item.id },
  });
  panel.append(
    ...(item.submenu ?? []).map((entry) =>
      entry.separator ? menuSeparator() : menuItemButton(entry, lang),
    ),
  );
  prepareMenu(panel, label);
  return [trigger, panel];
};

const createMoreFunctionsMenu = (lang: Locale): HTMLDivElement => {
  const menu = document.createElement('div');
  menu.className = 'fc-tb__menu fc-tb__menu--mac fc-tb__menu--mac-more';
  menu.id = MAC_FORMULAS_MORE_MENU_ID;
  menu.hidden = true;

  for (const { category } of MORE_FUNCTION_CATEGORIES) {
    const localizedLabel = macFunctionCategoryLabel(category, lang);
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
        menuItemButton({ id: macFunctionCommandId(name), label: name }, lang),
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
export const createMacRibbonMenuFactory = (lang: Locale): MacRibbonMenuFactory => {
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
      const gridClass = shapes ? 'fc-tb__visual-grid--shapes' : 'fc-tb__visual-grid--charts';
      let grid: HTMLElement | null = null;
      for (const item of items) {
        if (item.heading) {
          menu.appendChild(
            menuSectionHeader(lang === 'ja' && item.labelJa ? item.labelJa : item.label),
          );
          grid = null;
          continue;
        }
        const button = menuItemButton(item, lang, true);
        if (button.classList.contains('fc-tb__visual-tile')) {
          if (!grid) {
            grid = visualMenuGrid(gridClass);
            menu.appendChild(grid);
          }
          grid.appendChild(button);
        } else {
          menu.appendChild(button);
          grid = null;
        }
      }
    } else {
      for (const item of items) {
        if (item.submenu) menu.append(...menuSubmenuRow(item, id, lang));
        else if (item.separator) menu.appendChild(menuSeparator());
        else menu.appendChild(menuItemButton(item, lang));
      }
    }
    return menu;
  };
};

/** Update the session MRU without replacing an unchanged menu or losing its keyboard focus. */
const projectMacRecentMenuPrepared = (
  host: HTMLElement,
  store: SpreadsheetStore,
  lang: Locale,
  liveSet: ReadonlySet<string> | undefined,
  liveSignature: string,
  availabilityContext?: MacFunctionAvailabilityContext,
): void => {
  const menu = host.querySelector<HTMLElement>('#menu-mac-formulas-recent');
  if (!menu) return;
  const names = getRecentFunctions(store, liveSet);
  const signature = `${lang}:${liveSignature}:${functionAvailabilitySignature(availabilityContext, lang)}:${names.join(',')}`;
  const rebuild = menu.dataset.recentFunctions !== signature;
  const wasRestricted = menu.dataset.functionPolicyRestricted === 'true';
  const restricted = interactionControllerFor(store)?.restricted ?? false;
  const focused =
    rebuild && menu.contains(menu.ownerDocument.activeElement)
      ? (menu.ownerDocument.activeElement as HTMLElement).dataset.ribbonCommand
      : undefined;
  const items: MacRibbonMenuItem[] = [
    ...names.map((name) => ({ id: macFunctionCommandId(name), label: name })),
    ...(MAC_RIBBON_MENU_ITEMS['mac.formulas.recent'] ?? []),
  ];
  if (rebuild) {
    menu.replaceChildren(
      ...items.map((item) => {
        const name = macFunctionNameOf(item.id);
        return menuItemButton(
          item,
          lang,
          false,
          name ? functionDisplayName(availabilityContext?.reader, lang, name) : undefined,
        );
      }),
    );
  }
  if (rebuild || restricted || wasRestricted) {
    const buttons = Array.from(menu.querySelectorAll<HTMLButtonElement>('[data-ribbon-command]'));
    for (const button of buttons) {
      const name = macFunctionNameOf(button.dataset.ribbonCommand ?? '');
      if (name) {
        projectFunctionButtonState(button, name, store, lang, availabilityContext);
        continue;
      }
      const decision = canExecuteBuiltInRequest(
        store,
        button.dataset.ribbonCommand ?? '',
        'ribbon',
      );
      projectDisabledState(
        button,
        !decision.allowed,
        decision.allowed ? null : localizedDecisionReason(decision, dictionaries[lang]),
        {
          datasetKey: RIBBON_DISABLED_REASON_KEY,
          titlePrefix: button.textContent ?? '',
        },
      );
    }
  }
  menu.dataset.recentFunctions = signature;
  menu.dataset.functionPolicyRestricted = String(restricted);
  if (focused) {
    const next =
      Array.from(menu.querySelectorAll<HTMLButtonElement>('[data-ribbon-command]')).find(
        (button) => button.dataset.ribbonCommand === focused,
      ) ??
      menu.querySelector<HTMLButtonElement>('[data-ribbon-command="mac.formulas.category.recent"]');
    next?.focus({ preventScroll: true });
  }
};

export const projectMacRecentMenu = (
  host: HTMLElement,
  store: SpreadsheetStore,
  lang: Locale,
  liveNames?: LiveFunctionNames,
  availabilityContext?: MacFunctionAvailabilityContext,
): void => {
  const { set, signature } = prepareLiveFunctionNames(liveNames, lang, availabilityContext?.reader);
  projectMacRecentMenuPrepared(host, store, lang, set, signature, availabilityContext);
};

const projectFunctionCategoryMenu = (
  menu: HTMLElement,
  category: CatalogFunctionCategory,
  store: SpreadsheetStore,
  lang: Locale,
  liveSet: ReadonlySet<string> | undefined,
  liveSignature: string,
  availabilityContext: MacFunctionAvailabilityContext | undefined,
): void => {
  const signature = `${lang}:${category}:${liveSignature}:${functionAvailabilitySignature(availabilityContext, lang)}`;
  const rebuild = menu.dataset.functionNames !== signature;
  const wasRestricted = menu.dataset.functionPolicyRestricted === 'true';
  const restricted = interactionControllerFor(store)?.restricted ?? false;
  const focused =
    rebuild && menu.contains(menu.ownerDocument.activeElement)
      ? (menu.ownerDocument.activeElement as HTMLElement).dataset.ribbonCommand
      : undefined;
  if (rebuild) {
    menu.replaceChildren(
      ...supportedFunctionNames(category, liveSet).map((name) =>
        menuItemButton(
          { id: macFunctionCommandId(name), label: name },
          lang,
          false,
          functionDisplayName(availabilityContext?.reader, lang, name),
        ),
      ),
    );
    prepareMenu(menu, menu.getAttribute('aria-label') ?? undefined);
  }
  if (rebuild || restricted || wasRestricted) {
    for (const button of menu.querySelectorAll<HTMLButtonElement>('[data-ribbon-command]')) {
      const command = button.dataset.ribbonCommand ?? '';
      const name = macFunctionNameOf(command);
      if (name) projectFunctionButtonState(button, name, store, lang, availabilityContext);
    }
  }
  menu.dataset.functionNames = signature;
  menu.dataset.functionPolicyRestricted = String(restricted);
  if (!focused) return;
  const next = menu.querySelector<HTMLButtonElement>(`[data-ribbon-command="${focused}"]`);
  if (next) next.focus({ preventScroll: true });
  else focusMenuItem(menu);
};

/** Re-projects the six curated Function Library menus and the six native More
 * submenus against the workbook's recognized names. Static membership remains
 * the fallback when the engine cannot enumerate its catalog; an empty live set
 * is authoritative and leaves the category panel empty. */
const projectMacFunctionCategoryMenusPrepared = (
  host: HTMLElement,
  store: SpreadsheetStore,
  lang: Locale,
  liveSet: ReadonlySet<string> | undefined,
  liveSignature: string,
  availabilityContext?: MacFunctionAvailabilityContext,
): void => {
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

export const projectMacFunctionCategoryMenus = (
  host: HTMLElement,
  store: SpreadsheetStore,
  lang: Locale,
  liveNames?: LiveFunctionNames,
  availabilityContext?: MacFunctionAvailabilityContext,
): void => {
  const { set, signature } = prepareLiveFunctionNames(liveNames, lang, availabilityContext?.reader);
  projectMacFunctionCategoryMenusPrepared(host, store, lang, set, signature, availabilityContext);
};

/** Projects the Mac recent and function-category menus from the instance's live function names. */
export const projectMacFunctionMenus = (
  host: HTMLElement,
  instance: SpreadsheetInstance,
  lang: Locale,
): void => {
  const reportedNames =
    typeof instance.workbook.functionNames === 'function'
      ? instance.workbook.functionNames()
      : null;
  // functionNames() is already canonical, so index it once without the WASM canonicalizer.
  const liveSet = reportedNames === null ? undefined : new Set(reportedNames);
  const liveSignature = reportedNames === null ? 'static' : reportedNames.join(',');
  const availabilityContext = {
    reader:
      reportedNames === null
        ? undefined
        : (instance.workbook as unknown as LocalizedFunctionCatalogReader),
    unavailableReason:
      instance.i18n.strings.fxDialog.functionUnavailable ?? defaultFunctionUnavailableReason(lang),
  };
  projectMacRecentMenuPrepared(
    host,
    instance.store,
    lang,
    liveSet,
    liveSignature,
    availabilityContext,
  );
  projectMacFunctionCategoryMenusPrepared(
    host,
    instance.store,
    lang,
    liveSet,
    liveSignature,
    availabilityContext,
  );
};
