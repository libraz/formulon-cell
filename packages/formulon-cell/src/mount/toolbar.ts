import { selectedSessionChart } from '../commands/chart-design.js';
import { subscribeRecentFunctions } from '../commands/function-history.js';
import { selectedSessionIllustration } from '../commands/session-illustration.js';
import { activeTableOverlay } from '../commands/table-design.js';
import { registerOverlayOwner } from '../interact/overlay-portal.js';
import { attachRibbonScreenTips } from '../interact/ribbon-screen-tip.js';
import { projectMacFunctionMenus } from '../toolbar/ribbon/mac/menus.js';

// `Spreadsheet.mountToolbar` — public entry that wires the ribbon into a host
// element on top of an existing `SpreadsheetInstance`.
//
// What the toolbar owns vs. what the caller owns:
//  - Toolbar owns: per-session UI state (active tab, collapsed flag, backstage
//    flag, display-menu flag, theme, border style/color, formula-bar visibility),
//    the renderer, click delegation, and the imperative `ToolbarInstance` API.
//  - Caller owns: the renderer helpers (select/color/icon/svg), the submenu
//    factories (`menus`), and the optional feature hooks (`hooks`). These
//    still live outside core because they reach into framework-specific or
//    app-specific glue (illustrations, custom dialog flows, …). Core ships
//    defaults for these (see `toolbar-defaults.ts`); the caller's versions
//    take precedence.
//
// The toolbar listens to `instance.store.subscribe()` so it can re-project
// active-state (bold/italic/etc.) on every selection or format change. It
// does NOT re-render the whole ribbon on every change — only the active-state
// projection runs in the hot path. Tab switches and similar topology changes
// go through `renderRibbon()` once.

import { withSelectionFormatOrigin } from '../commands/format.js';
import { interactionControllerFor } from '../commands/interaction-controller.js';
import { recordRepeatableFormatChange } from '../commands/slice-history.js';
import { resolveSpreadsheetPlatform, type SpreadsheetPlatform } from '../extensions/ui-options.js';
import { ensureMacInk, getMacInk, type MacInkController } from '../interact/mac-ink.js';
import type { CellBorderStyle, State } from '../store/types.js';
import { cancelOpenAppDialogs } from '../toolbar/dialogs/shell.js';
import { ribbonDisplayText, type ToolbarMenuText, toolbarMenuText } from '../toolbar/menu-text.js';
import {
  isRibbonMenuFirstCommand,
  RIBBON_BORDERS_MENU_ID,
  RIBBON_HOST_MENU_FIRST_COMMANDS,
} from '../toolbar/ribbon/activation.js';
import {
  applyRibbonCommand,
  type RibbonHooks,
  type RibbonRuntime,
} from '../toolbar/ribbon/apply-ribbon-command.js';
import {
  type BorderMenuApi,
  type BorderMenuCtx,
  createBorderMenu,
} from '../toolbar/ribbon/border-menu.js';
import type { RibbonFormatMutator } from '../toolbar/ribbon/command-tables.js';
import {
  type DynamicDropdownsApi,
  type DynamicDropdownsCtx,
  ribbonDropdownMenuIdForCommand,
} from '../toolbar/ribbon/dynamic-dropdowns.js';
import { attachRibbonFit } from '../toolbar/ribbon/fit-ribbon.js';
import { toolbarLangForLocale } from '../toolbar/ribbon/mac/locale.js';
import { projectMacRibbonState } from '../toolbar/ribbon/mac/model.js';
import {
  createRenderRibbon,
  projectRibbonShell,
  type RibbonDisplayMode,
  type RibbonMenus,
  type RibbonRenderHelpers,
} from '../toolbar/ribbon/render-ribbon.js';
import {
  closeStaticRibbonMenus,
  hasOpenStaticRibbonMenu,
  toggleStaticRibbonSubmenu,
} from '../toolbar/ribbon/static-menus.js';
import { projectRibbonActiveState } from '../toolbar/ribbon-active-state.js';
import {
  CHART_CONTEXTUAL_RIBBON_TABS,
  CONTEXTUAL_RIBBON_TABS,
  EXCEL365_MAC_RIBBON_TABS,
  EXCEL365_STANDARD_RIBBON_TABS,
  PICTURE_CONTEXTUAL_RIBBON_TABS,
  type RibbonTab,
  SHAPE_CONTEXTUAL_RIBBON_TABS,
  TABLE_CONTEXTUAL_RIBBON_TABS,
  type ToolbarLang,
  type ToolbarText,
  toolbarText,
} from '../toolbar/ribbon-model.js';
import {
  createDefaultRibbonHelpers,
  createDefaultRibbonHooks,
  createDefaultRibbonMenus,
} from './toolbar-defaults.js';
import {
  attachDynamicDropdownDelegation,
  type DynamicDropdownDelegation,
} from './toolbar-dropdown-delegation.js';
import { attachRibbonInteractionPolicy } from './toolbar-interaction-policy.js';
import type { SpreadsheetInstance } from './types.js';

type UiTheme = 'paper' | 'ink' | 'contrast';

export type { RibbonDisplayMode } from '../toolbar/ribbon/render-ribbon.js';

const DEFAULT_BORDER_STYLE: CellBorderStyle = 'thin';
const DEFAULT_BORDER_COLOR = '#000000';

const isImmediateMacFormulaOpener = (commandId: string): boolean =>
  commandId === 'mac.formulas.insertFunction' || commandId.startsWith('mac.function.');

const isImmediateMacFormulaPrimary = (target: Element, commandId: string): boolean =>
  isImmediateMacFormulaOpener(commandId) &&
  !target.closest('.fc-tb__rb-split-chevron') &&
  !isRibbonMenuFirstCommand(commandId);

export interface MountToolbarOptions {
  /** Platform used for platform-specific ribbon defaults. Inherited from the
   *  nearest `.fc-host` when omitted. */
  platform?: SpreadsheetPlatform;
  /** Language for built-in ribbon labels. Defaults to the instance locale. */
  lang?: ToolbarLang;
  /** Override the auto-derived ToolbarText (button titles, group names). */
  text?: ToolbarText;
  /** Override the auto-derived ToolbarMenuText (dropdown labels). */
  menuText?: ToolbarMenuText;

  /** Renderer helpers from `createControlDispatch` / `createSelectColorRibbon`.
   *  Optional — when omitted the toolbar uses `createDefaultRibbonHelpers`
   *  which composes both factories from `instance` plus the wired sheet /
   *  focus / refresh closures below. Pass a partial helpers bundle to swap a
   *  single factory (e.g. a custom color picker) without taking over the
   *  whole set. */
  helpers?: Partial<RibbonRenderHelpers>;

  /** Submenu factories keyed by category. Missing entries leave the matching
   *  split-button without a dropdown — useful for trimming the toolbar to a
   *  feature subset. */
  menus?: RibbonMenus;

  /** Feature hooks the toolbar dispatches into for commands beyond core's
   *  built-ins (clipboard, sort/filter, insert, page, review, …). Each group
   *  is independently optional. */
  hooks?: RibbonHooks;

  /** Factory for the element rendered while `setBackstageOpen(true)` is in
   *  effect. Clicking the "File" tab does not open it. Core wires only the
   *  Back button and Escape; other `data-backstage-action` buttons are the
   *  host's to handle. Without it, the open backstage is an empty panel. */
  createBackstageView?: () => HTMLElement;

  /** Initial UI state. */
  activeTab?: RibbonTab;
  /** Explicit ribbon tab surface. Pass `EXCEL365_STANDARD_RIBBON_TABS` for
   *  the Microsoft 365 baseline and append optional add-in/automation tabs
   *  only when the host has those surfaces wired. Defaults to
   *  `EXCEL365_STANDARD_RIBBON_TABS` (`EXCEL365_MAC_RIBBON_TABS` on macOS). */
  ribbonTabs?: readonly RibbonTab[];
  ribbonDisplayMode?: RibbonDisplayMode;
  collapsed?: boolean;
  formulaBarVisible?: boolean;
  theme?: UiTheme;
  borderStyle?: CellBorderStyle;
  borderColor?: string;

  /** Click delegation toggles. Tabs / display-menu / display-option clicks
   *  are always handled. `commandDelegation` controls whether the toolbar
   *  also auto-dispatches `[data-ribbon-command]` clicks through
   *  `applyRibbonCommand`. Set false when the host wires its own command
   *  handlers (e.g. legacy `btn-*` listeners) to avoid double-firing. */
  commandDelegation?: boolean;

  /** Lets the host short-circuit a ribbon command before it reaches
   *  `applyRibbonCommand`. Returning `true` means "handled — skip
   *  dispatch", `false` (or undefined) means "fall through to dispatch".
   *  Used by hosts that own dropdown menus tied to ribbon commands (the
   *  playground wires `dynamic-dropdowns` open/close here so the menu
   *  survives ribbon re-renders without per-button listeners).
   *  The third arg is the click event so split-button hosts can branch
   *  on whether the chevron (vs the primary face) was clicked. */
  interceptCommand?: (id: string, button: HTMLButtonElement, event: MouseEvent) => boolean;

  /** Opt-in for the built-in dropdown-menu click delegation. Pass `true` for
   *  the full default ctx (fill / clear / autosum / etc. derived from the
   *  instance), a partial bag whose keys override individual handlers, or a
   *  getter that returns the partial bag — the getter form is for hosts
   *  whose ctx isn't ready at mount time (the playground builds its ctx
   *  after `mountToolbar` returns). When omitted, clicks inside an open
   *  menu do nothing. */
  dynamicDropdowns?: true | Partial<DynamicDropdownsCtx> | (() => Partial<DynamicDropdownsCtx>);

  /** Pluggable runtime hooks the dispatcher needs but core can't yet derive
   *  on its own. Each falls back to a no-op or a sensible default. */
  focusSheet?: () => void;
  refreshCells?: () => void;
  refreshZoom?: () => void;
  projectFormatToolbar?: () => void;
  showMessage?: RibbonRuntime['showMessage'];
  applyRibbonFormat?: RibbonRuntime['applyRibbonFormat'];

  /** Lifecycle callbacks fired after the matching internal state change.
   *  `onTabChange('file')` is the hook for hosts that render their own
   *  backstage; the File tab does not open one by itself. */
  onTabChange?: (tab: RibbonTab) => void;
  onCollapsedChange?: (collapsed: boolean) => void;
  onDisplayModeChange?: (mode: RibbonDisplayMode) => void;
  onBackstageOpenChange?: (open: boolean) => void;
  onThemeChange?: (theme: UiTheme) => void;
  onFormulaBarChange?: (visible: boolean) => void;
  /** Fires after a ribbon command dispatched. `applied=false` means
   *  `applyRibbonCommand` returned false (no handler matched) — useful for
   *  custom command ids the host wants to handle as a fallback. */
  onCommand?: (id: string, applied: boolean) => void;
}

/** Either a directly held instance or a late-bound getter — useful when the
 *  spreadsheet mount is async and the toolbar shell needs to render its empty
 *  state synchronously. The getter is re-invoked on every dispatch so the
 *  toolbar always reaches the current instance. */
export type ToolbarInstanceRef = SpreadsheetInstance | (() => SpreadsheetInstance | null);

export interface ToolbarInstance {
  readonly host: HTMLElement;
  /** The current spreadsheet instance, or `null` when mounted with a deferred
   *  getter that hasn't been satisfied yet. */
  readonly instance: SpreadsheetInstance | null;
  /** Full re-render of the ribbon shell. Costly — only call on topology
   *  changes (tab switch, collapse toggle, backstage open). State changes
   *  inside the active panel should ride on `instance.store.subscribe`. */
  rerender(): void;
  /** Dispatch a ribbon command id through the same path as a button click.
   *  Returns true when a handler matched. */
  applyCommand(id: string): boolean;
  /** Focuses the active ribbon tab. Used by Excel-style F6 landmark cycling. */
  focusActiveTab(): boolean;
  setActiveTab(tab: RibbonTab): void;
  getActiveTab(): RibbonTab;
  setCollapsed(collapsed: boolean): void;
  getCollapsed(): boolean;
  setDisplayMode(mode: RibbonDisplayMode): void;
  getDisplayMode(): RibbonDisplayMode;
  /** Show or hide core's presentational backstage (`createBackstageView`).
   *  Core closes it on Back and Escape; every other action is host-handled. */
  setBackstageOpen(open: boolean): void;
  /** Whether core's presentational backstage is currently shown. */
  getBackstageOpen(): boolean;
  setDisplayMenuOpen(open: boolean): void;
  setFormulaBarVisible(visible: boolean): void;
  getFormulaBarVisible(): boolean;
  setTheme(theme: UiTheme): void;
  getTheme(): UiTheme;
  setBorderStyle(style: CellBorderStyle): void;
  getBorderStyle(): CellBorderStyle;
  setBorderColor(color: string): void;
  getBorderColor(): string;
  /** When the toolbar was mounted with `dynamicDropdowns` enabled, exposes
   *  the core's dropdown api so hosts can drive open/close (interceptCommand,
   *  click-outside, arrow-key nav) without re-creating their own ctx. Null
   *  when the option was omitted. */
  readonly dropdownsApi: DynamicDropdownsApi | null;
  dispose(): void;
}

export { RIBBON_HOST_MENU_FIRST_COMMANDS } from '../toolbar/ribbon/activation.js';

/** `interceptCommand` implementation for the menu-first contract above. Pass
 *  it straight through from a host's `mountToolbar` options; it returns false
 *  for every other command so the default dispatch still runs. */
export const openHostMenuFirstDropdown = (
  toolbar: ToolbarInstance | null,
  command: string,
  button: HTMLButtonElement,
): boolean => {
  if (!RIBBON_HOST_MENU_FIRST_COMMANDS.has(command)) return false;
  const menuId = ribbonDropdownMenuIdForCommand(command);
  const api = toolbar?.dropdownsApi;
  if (!menuId || !api) return false;
  api.openDynamicRibbonDropdown({ command, menuId }, button);
  return true;
};

const defaultApplyRibbonFormat =
  (getInstance: () => SpreadsheetInstance | null) =>
  (fn: RibbonFormatMutator): void => {
    const inst = getInstance();
    if (!inst) return;
    recordRepeatableFormatChange(inst.history, inst.store, () =>
      withSelectionFormatOrigin(inst.store, 'ribbon', () => fn(inst.store.getState(), inst.store)),
    );
  };

export function mountToolbar(
  host: HTMLElement,
  instance: ToolbarInstanceRef,
  opts: MountToolbarOptions,
): ToolbarInstance {
  if (!host) throw new Error('Spreadsheet.mountToolbar: host element required');
  if (instance === null || instance === undefined) {
    throw new Error('Spreadsheet.mountToolbar: instance ref required');
  }

  const getInstance: () => SpreadsheetInstance | null =
    typeof instance === 'function'
      ? (instance as () => SpreadsheetInstance | null)
      : () => instance;
  const initialInstance = getInstance();
  const previousPlatform = host.dataset.fcPlatform;
  const previousLanguage = host.getAttribute('lang');
  const inheritedPlatform =
    host.dataset.fcPlatform ??
    host.closest<HTMLElement>('.fc-host')?.dataset.fcPlatform ??
    initialInstance?.host.dataset.fcPlatform;
  const platformOverride = opts.platform !== undefined;
  const ribbonTabsOverride = opts.ribbonTabs !== undefined;
  let platform = resolveSpreadsheetPlatform(
    opts.platform ??
      (inheritedPlatform === 'mac' || inheritedPlatform === 'default' ? inheritedPlatform : 'auto'),
  );
  host.dataset.fcPlatform = platform;
  const platformRibbonTabs = (p: typeof platform): readonly RibbonTab[] =>
    p === 'mac' ? EXCEL365_MAC_RIBBON_TABS : EXCEL365_STANDARD_RIBBON_TABS;
  let defaultRibbonTabs: readonly RibbonTab[] = opts.ribbonTabs ?? platformRibbonTabs(platform);

  // Probe once at mount-time for language inference and the initial subscribe.
  // The toolbar continues to work if the probe returns null (deferred mount);
  // in that case lang falls back to opts.lang or 'ja' and the store subscription
  // is attached lazily on the first call to `attachStoreSubscription`.
  const unregisterOverlayOwner = registerOverlayOwner(host, () => getInstance()?.host ?? null);

  const langOverride = opts.lang;
  let lang: ToolbarLang =
    opts.lang ?? (initialInstance ? toolbarLangForLocale(initialInstance.i18n.locale) : 'ja');
  host.lang = lang;
  const screenTips = attachRibbonScreenTips({
    host,
    getSheetHost: () => getInstance()?.host,
    getPlatform: () => platform,
  });
  const ribbonFit = attachRibbonFit(host);
  let text = opts.text ?? toolbarText(lang);
  let menuText = opts.menuText ?? toolbarMenuText(lang);
  let displayOptionsText = ribbonDisplayText(lang);

  let activeTab: RibbonTab = opts.activeTab ?? 'home';
  let displayMode: RibbonDisplayMode =
    opts.ribbonDisplayMode ?? (opts.collapsed ? 'tabsOnly' : 'full');
  let ribbonPeek = false;
  let backstageOpen = false;
  let backstageReturnTab: RibbonTab = 'home';
  let displayMenuOpen = false;
  let formulaBarVisible = opts.formulaBarVisible ?? true;
  let theme: UiTheme = opts.theme ?? 'paper';
  let borderStyle: CellBorderStyle = opts.borderStyle ?? DEFAULT_BORDER_STYLE;
  let borderColor = opts.borderColor ?? DEFAULT_BORDER_COLOR;

  // Tabs a selected chart, shape or picture adds, and the tab to go back to once it is deselected.
  const NO_CONTEXTUAL_TABS: readonly RibbonTab[] = [];
  let contextualTabs = NO_CONTEXTUAL_TABS;
  let contextualReturnTab: RibbonTab = 'home';
  let knownObjectIds: ReadonlySet<string> = new Set();
  // Objects added in the current task: an insert or paste selects the object it adds.
  const freshObjectIds = new Set<string>();
  const isContextualTab = (tab: RibbonTab): boolean => CONTEXTUAL_RIBBON_TABS.includes(tab);

  const rehomeActiveTabForPlatform = (): void => {
    const available = opts.ribbonTabs ?? defaultRibbonTabs;
    if (available.includes(activeTab) || contextualTabs.includes(activeTab)) return;
    activeTab = available[0] ?? 'home';
  };
  rehomeActiveTabForPlatform();

  /** The tabs of the selected object, as one of the shared lists so that an
   *  unchanged selection compares equal. */
  const selectedObjectTabs = (state: State): readonly RibbonTab[] => {
    if (selectedSessionChart(state)) return CHART_CONTEXTUAL_RIBBON_TABS;
    const item = selectedSessionIllustration(state);
    if (!item) return activeTableOverlay(state) ? TABLE_CONTEXTUAL_RIBBON_TABS : NO_CONTEXTUAL_TABS;
    return item.kind === 'image' ? PICTURE_CONTEXTUAL_RIBBON_TABS : SHAPE_CONTEXTUAL_RIBBON_TABS;
  };

  /** Show the tabs of the selected chart, shape or picture; like the
   *  spreadsheet, an object selected as it is added opens its first tab.
   *  True when the tabs changed. */
  const syncContextualTabs = (current: SpreadsheetInstance | null, adopt: boolean): boolean => {
    const state = current?.store.getState();
    const previousIds = knownObjectIds;
    knownObjectIds = new Set(
      state
        ? [...state.charts.charts, ...state.illustrations.illustrations].map((entry) => entry.id)
        : [],
    );
    for (const id of knownObjectIds) {
      if (previousIds.has(id) || !adopt) continue;
      if (freshObjectIds.size === 0) queueMicrotask(() => freshObjectIds.clear());
      freshObjectIds.add(id);
    }
    const next = state ? selectedObjectTabs(state) : NO_CONTEXTUAL_TABS;
    const selectedId = state?.ui.selectedObject?.id;
    const inserted =
      next !== NO_CONTEXTUAL_TABS && selectedId !== undefined && freshObjectIds.delete(selectedId);
    if (next === contextualTabs && !inserted) return false;
    contextualTabs = next;
    const previousTab = activeTab;
    if (inserted) {
      if (!isContextualTab(activeTab)) contextualReturnTab = activeTab;
      activeTab = next[0] ?? activeTab;
    } else if (isContextualTab(activeTab) && !next.includes(activeTab)) {
      activeTab = contextualReturnTab;
    }
    if (activeTab !== previousTab) opts.onTabChange?.(activeTab);
    return true;
  };

  const focusSheet =
    opts.focusSheet ??
    ((): void => {
      getInstance()?.host.focus();
    });
  const refreshCells = opts.refreshCells ?? ((): void => undefined);
  const refreshZoom = opts.refreshZoom ?? ((): void => undefined);
  const projectFormatToolbar = (): void => {
    projectRibbonActiveState(host, getInstance(), {
      background: text.background,
      sheetBackgroundClear: menuText.sheetBackgroundClear,
    });
    const instance = getInstance();
    if (instance && platform === 'mac') {
      projectMacFunctionMenus(host, instance, lang);
      projectMacRibbonState(host, instance);
    }
    opts.projectFormatToolbar?.();
  };
  const showMessage = opts.showMessage ?? ((): void => undefined);
  const rawApplyRibbonFormat = opts.applyRibbonFormat ?? defaultApplyRibbonFormat(getInstance);
  const applyRibbonFormat = (fn: RibbonFormatMutator): void =>
    rawApplyRibbonFormat((state, store) =>
      withSelectionFormatOrigin(store, 'ribbon', () => fn(state, store)),
    );
  const isCollapsedMode = (): boolean => displayMode === 'tabsOnly' || displayMode === 'autoHide';
  let borderMenuApi: BorderMenuApi | null = null;
  const setDisplayMode = (next: RibbonDisplayMode): void => {
    if (next === displayMode) return;
    const wasCollapsed = isCollapsedMode();
    displayMode = next;
    ribbonPeek = false;
    opts.onDisplayModeChange?.(next);
    const collapsed = isCollapsedMode();
    if (collapsed !== wasCollapsed) opts.onCollapsedChange?.(collapsed);
    renderToolbar();
  };

  // Defaults are rebuilt when the instance locale changes. They close over
  // the live `borderStyle/borderColor` via getters so the borders submenu
  // always picks the most recent value. The host may still override
  // individual helpers/menus/hooks by spreading on top.
  const defaultsInstance = initialInstance;
  let mergedHelpers: RibbonRenderHelpers;
  let mergedMenus: RibbonMenus;
  let mergedHooks: RibbonHooks;
  const rebuildLocalizedDefaults = (): void => {
    const defaultHelpers: RibbonRenderHelpers | null = defaultsInstance
      ? createDefaultRibbonHelpers(defaultsInstance, {
          lang,
          focusSheet,
          refreshCells,
          projectFormatToolbar,
        })
      : null;
    const defaultMenus: RibbonMenus = defaultsInstance
      ? createDefaultRibbonMenus(defaultsInstance, {
          lang,
          getBorderColor: () => borderColor,
          setBorderColor: (color) => {
            borderColor = color;
            defaultsInstance.borderDraw?.setColor(color);
          },
        })
      : {};
    const defaultHooks: RibbonHooks = defaultsInstance
      ? createDefaultRibbonHooks(defaultsInstance, { lang, refreshZoom })
      : {};

    mergedHelpers = {
      ...(defaultHelpers ?? {}),
      ...(opts.helpers ?? {}),
    } as RibbonRenderHelpers;
    mergedMenus = { ...defaultMenus, ...(opts.menus ?? {}) };
    // Hooks merge: per-category shallow merge so the host can extend (not
    // replace) any single group. Categories the host doesn't mention keep the
    // defaults; categories it does mention spread on top of the default ones.
    mergedHooks = { ...defaultHooks };
    if (opts.hooks) {
      for (const key of Object.keys(opts.hooks) as (keyof RibbonHooks)[]) {
        const hostGroup = opts.hooks[key];
        if (!hostGroup) continue;
        const defaultGroup = mergedHooks[key];
        // biome-ignore lint/suspicious/noExplicitAny: index access onto union
        (mergedHooks as any)[key] = { ...(defaultGroup ?? {}), ...hostGroup };
      }
    }
  };
  rebuildLocalizedDefaults();

  // Auto-wire the default dynamic-dropdowns delegation when the host opts in;
  // dispose detaches its document listeners so they do not leak across re-mounts.
  let dropdownsDelegation: DynamicDropdownDelegation | null = null;
  let dropdownsApi: DynamicDropdownsApi | null = null;
  if (opts.dynamicDropdowns) {
    dropdownsDelegation = attachDynamicDropdownDelegation(opts.dynamicDropdowns, {
      defaultsInstance,
      getInstance,
      focusSheet,
      projectFormatToolbar,
      refreshCells,
      closeBorderMenu: (restoreFocus?: boolean) => {
        borderMenuApi?.closeBorderMenu(restoreFocus);
      },
      onMenuClosedByCommand: () => dismissRibbonPeek(),
    });
    dropdownsApi = dropdownsDelegation.api;
  }

  const createRibbonRenderer = () =>
    createRenderRibbon({
      getInst: getInstance,
      ribbonLang: lang,
      ribbonText: text,
      ribbonMenuText: menuText,
      ribbonDisplayOptionsText: displayOptionsText,
      getProfile: () => (platform === 'mac' ? 'excel365Mac' : 'default'),
      get ribbonTabs() {
        return defaultRibbonTabs;
      },
      ribbonRoot: host,
      state: {
        getActiveTab: () => activeTab,
        getCollapsed: () => isCollapsedMode(),
        getDisplayMode: () => displayMode,
        getAutoHidePeek: () => ribbonPeek,
        getBackstageOpen: () => backstageOpen,
        getDisplayMenuOpen: () => displayMenuOpen,
        getFormulaBarVisible: () => formulaBarVisible,
        getContextualTabs: () => contextualTabs,
      },
      helpers: mergedHelpers,
      menus: mergedMenus,
      createBackstageView: opts.createBackstageView ?? (() => document.createElement('div')),
      projectFormatToolbar,
    });
  let renderApi = createRibbonRenderer();

  const wireBorderMenu = (): void => {
    borderMenuApi?.detach();
    borderMenuApi = null;
    if (!host.querySelector(`#${RIBBON_BORDERS_MENU_ID}`)) return;
    const current = getInstance();
    if (!current) return;
    if (interactionControllerFor(current.store)?.policy !== undefined) return;
    borderMenuApi = createBorderMenu({
      getInst: getInstance as unknown as BorderMenuCtx['getInst'],
      sheetEl: current.host,
      getSelectedBorderStyle: () => borderStyle,
      setSelectedBorderStyle: (style) => {
        borderStyle = style;
        current.borderDraw?.setStyle(style);
      },
      getSelectedBorderColor: () => borderColor,
      applyRibbonFormat: applyRibbonFormat as unknown as BorderMenuCtx['applyRibbonFormat'],
    });
  };

  const renderToolbar = (): void => {
    screenTips.refresh();
    host.lang = lang;
    borderMenuApi?.detach();
    borderMenuApi = null;
    renderApi.renderRibbon();
    ribbonFit.refit();
    wireBorderMenu();
    interactionPolicy.project();
    host
      .querySelector<HTMLElement>('[data-ribbon-tab][aria-selected="true"]')
      ?.scrollIntoView?.({ block: 'nearest', inline: 'nearest' });
  };

  const syncInheritedPlatform = (current: SpreadsheetInstance | null): boolean => {
    if (platformOverride) return false;
    const inherited = current?.host.dataset.fcPlatform;
    const next = resolveSpreadsheetPlatform(
      inherited === 'mac' || inherited === 'default' ? inherited : 'auto',
    );
    if (next === platform) return false;
    platform = next;
    host.dataset.fcPlatform = next;
    if (!ribbonTabsOverride) {
      defaultRibbonTabs = platformRibbonTabs(next);
      rehomeActiveTabForPlatform();
    }
    return true;
  };

  // Mac Draw owns transient controller state outside the spreadsheet store.
  // Bind its notifications only while this toolbar is on the Mac surface so
  // active pen / trackpad buttons update without a synthetic store mutation.
  let unsubMacInk: (() => void) | null = null;
  let subscribedMacInk: MacInkController | null = null;
  const syncMacInkSubscription = (current: SpreadsheetInstance | null): void => {
    const next =
      platform === 'mac' && current ? (getMacInk(current) ?? ensureMacInk(current) ?? null) : null;
    if (next === subscribedMacInk) return;
    unsubMacInk?.();
    unsubMacInk = null;
    subscribedMacInk = next;
    if (next) unsubMacInk = next.subscribe(projectFormatToolbar);
  };

  const interactionPolicy = attachRibbonInteractionPolicy(host, {
    getInstance,
    closeDynamicDropdowns: () => dropdownsApi?.closeAllDynamicRibbonDropdowns(),
  });

  const applyCommand = (id: string): boolean => {
    const applied = applyRibbonCommand(id, {
      inst: getInstance(),
      text,
      menuText,
      ui: { theme, borderStyle, borderColor, formulaBarVisible },
      runtime: {
        focusSheet,
        refreshCells,
        refreshZoom,
        projectFormatToolbar,
        applyRibbonFormat,
        applyUiTheme: (next) => {
          theme = next;
          opts.onThemeChange?.(next);
          renderToolbar();
        },
        setFormulaBarVisible: (next) => {
          formulaBarVisible = next;
          opts.onFormulaBarChange?.(next);
        },
        featureFlags: renderApi.playgroundFeatureFlags,
        showMessage,
      },
      hooks: mergedHooks,
    });
    opts.onCommand?.(id, applied);
    dismissRibbonPeek();
    return applied;
  };

  const focusActiveTab = (): boolean => {
    const tab =
      host.querySelector<HTMLButtonElement>(`[data-ribbon-tab="${activeTab}"]`) ??
      host.querySelector<HTMLButtonElement>('[data-ribbon-tab]');
    if (!tab || tab.disabled) return false;
    tab.focus({ preventScroll: true });
    return document.activeElement === tab;
  };

  const projectRibbonTabs = (): void => {
    projectRibbonShell(host, { activeTab, displayMode, peekRequested: ribbonPeek });
    ribbonFit.refit();
    host
      .querySelector<HTMLElement>('[data-ribbon-tab][aria-selected="true"]')
      ?.scrollIntoView?.({ block: 'nearest', inline: 'nearest' });
  };

  const onRibbonFocus = (event: FocusEvent): void => {
    const target = event.target as HTMLElement | null;
    if (!target?.closest || !host.contains(target)) return;
    if (target.closest('.fc-tb__menu, .fc-tb__submenu, .fc-tb__rb-dd')) return;
    if (target.closest('.fc-tb__ribbon-tabs, .fc-tb__ribbon')) {
      target.scrollIntoView?.({ block: 'nearest', inline: 'nearest' });
    }
  };
  host.addEventListener('focusin', onRibbonFocus, true);

  const dismissRibbonPeek = (): void => {
    if (!ribbonPeek) return;
    ribbonPeek = false;
    dropdownsApi?.closeAllDynamicRibbonDropdowns();
    closeStaticRibbonMenus(host);
    projectRibbonTabs();
  };

  const activateRibbonTab = (tab: RibbonTab, reveal: boolean): void => {
    const button = host.querySelector<HTMLButtonElement>(`[data-ribbon-tab="${tab}"]`);
    if (!button || button.disabled || button.getAttribute('aria-disabled') === 'true') return;
    dropdownsApi?.closeAllDynamicRibbonDropdowns();
    closeStaticRibbonMenus(host);
    const changed = tab !== activeTab;
    if (tab === 'file' && activeTab !== 'file') backstageReturnTab = activeTab;
    if (isContextualTab(tab) && !isContextualTab(activeTab)) contextualReturnTab = activeTab;
    activeTab = tab;
    if (reveal && tab !== 'file' && isCollapsedMode()) ribbonPeek = true;
    projectRibbonTabs();
    if (changed) opts.onTabChange?.(tab);
  };

  const closeBackstage = (): void => {
    if (!backstageOpen) return;
    backstageOpen = false;
    const restored = activeTab === 'file' && backstageReturnTab !== 'file';
    if (restored) activeTab = backstageReturnTab;
    opts.onBackstageOpenChange?.(false);
    renderToolbar();
    if (restored) opts.onTabChange?.(activeTab);
  };

  const hasOpenStaticMenu = (): boolean => !dropdownsApi && hasOpenStaticRibbonMenu(host);

  const onMouseDown = (e: MouseEvent): void => {
    if (e.button !== 0) return;
    const target = e.target;
    if (!(target instanceof Element)) return;
    const cmdBtn = target.closest<HTMLButtonElement>('[data-ribbon-command]');
    const id = cmdBtn?.dataset.ribbonCommand;
    if (
      !cmdBtn ||
      !id ||
      cmdBtn.disabled ||
      cmdBtn.getAttribute('aria-disabled') === 'true' ||
      !isImmediateMacFormulaPrimary(target, id)
    )
      return;
    // Preserve the active formula edit until the palette has captured it.
    e.preventDefault();
  };

  const onClick = (e: MouseEvent): void => {
    const target = e.target;
    if (!(target instanceof Element)) return;

    if (backstageOpen && target.closest('[data-backstage-action="back"]')) {
      closeBackstage();
      return;
    }

    const tabBtn = target.closest<HTMLButtonElement>('[data-ribbon-tab]');
    if (tabBtn) {
      const tab = tabBtn.dataset.ribbonTab as RibbonTab | undefined;
      if (tab) activateRibbonTab(tab, true);
      return;
    }

    const toggleBtn = target.closest<HTMLButtonElement>('[data-ribbon-toggle]');
    if (toggleBtn) {
      displayMenuOpen = !displayMenuOpen;
      renderToolbar();
      return;
    }

    const optionBtn = target.closest<HTMLButtonElement>('[data-ribbon-display-option]');
    if (optionBtn) {
      const option = optionBtn.dataset.ribbonDisplayOption;
      displayMenuOpen = false;
      if (
        option === 'full' ||
        option === 'singleLine' ||
        option === 'tabsOnly' ||
        option === 'autoHide'
      ) {
        setDisplayMode(option);
      } else if (option === 'expanded') setDisplayMode('full');
      else if (option === 'collapsed') setDisplayMode('tabsOnly');
      else renderToolbar();
      return;
    }

    if (opts.commandDelegation === false) {
      const command = target.closest<HTMLButtonElement>('[data-ribbon-command]');
      if (
        command &&
        !command.disabled &&
        !isRibbonMenuFirstCommand(command.dataset.ribbonCommand ?? '') &&
        !target.closest('.fc-tb__rb-split-chevron')
      )
        dismissRibbonPeek();
      return;
    }
    const cmdBtn = target.closest<HTMLButtonElement>('[data-ribbon-command]');
    if (cmdBtn?.dataset.ribbonCommand) {
      const id = cmdBtn.dataset.ribbonCommand;
      if (cmdBtn.disabled || cmdBtn.getAttribute('aria-disabled') === 'true') return;
      // WebKit follows the macOS convention of not focusing a <button> on
      // click. Ribbon keyboard navigation and host dialogs that restore focus
      // to the command that opened them both rely on the invoked command being
      // the active element, so normalize it here. Anything the command itself
      // focuses afterwards (menu item, dialog field, the sheet) still wins.
      if (document.activeElement !== cmdBtn && !isImmediateMacFormulaPrimary(target, id)) {
        cmdBtn.focus({ preventScroll: true });
      }
      if (opts.interceptCommand?.(id, cmdBtn, e)) {
        const openMenu = host.querySelector(
          '.fc-tb__menu:not([hidden]), .fc-tb__submenu:not([hidden]), .fc-tb__rb-dd--open',
        );
        if (!openMenu) dismissRibbonPeek();
        return;
      }
      // The chevron is part of the command button's DOM, so delegated clicks
      // otherwise look identical to primary-face clicks. Route it to the
      // attached menu before the primary action; SVG paths are covered by
      // `closest()` just like the SVG element itself.
      if (target.closest('.fc-tb__rb-split-chevron') && cmdBtn.dataset.ribbonMenuId) {
        const menuId = cmdBtn.dataset.ribbonMenuId;
        if (dropdownsApi) {
          dropdownsApi.openDynamicRibbonDropdown({ command: id, menuId }, cmdBtn);
          return;
        }
        toggleStaticRibbonSubmenu(host, cmdBtn);
        return;
      }
      // Fallback dropdown behaviour: if the button has a sibling submenu
      // attached via render-ribbon's `tools.appendChild(submenu())`, toggle
      // it. Split buttons with a primary face action skip this so
      // applyRibbonCommand can fire their primary handler; their chevron
      // routing is handled by the branch above.
      if (isRibbonMenuFirstCommand(id) && !cmdBtn.closest('.fc-tb__menu--mac')) {
        const menuId = cmdBtn.dataset.ribbonMenuId;
        if (dropdownsApi && menuId) {
          dropdownsApi.openDynamicRibbonDropdown({ command: id, menuId }, cmdBtn);
          return;
        }
        if (toggleStaticRibbonSubmenu(host, cmdBtn)) return;
      }
      const immediateFunctionLeaf =
        id.startsWith('mac.function.') && isImmediateMacFormulaPrimary(target, id);
      let returnFocusTarget: HTMLElement | null = null;
      if (cmdBtn.closest('.fc-tb__menu--mac')) {
        const menu = cmdBtn.closest<HTMLElement>('.fc-tb__menu--mac');
        const spec = menu ? dropdownsApi?.dynamicDropdownSpecForMenu(menu) : null;
        returnFocusTarget =
          immediateFunctionLeaf && spec
            ? (dropdownsApi?.dynamicDropdownButtonForSpec(spec) ?? null)
            : null;
        if (spec) dropdownsApi?.closeDynamicRibbonDropdown(spec, !immediateFunctionLeaf);
        closeStaticRibbonMenus(host);
      }
      const applied = applyCommand(id);
      if (applied && immediateFunctionLeaf && returnFocusTarget) {
        const dialog = getInstance()?.features.fxDialog as
          | { setReturnFocusTarget?: (target: HTMLElement | null) => void }
          | undefined;
        dialog?.setReturnFocusTarget?.(returnFocusTarget);
      }
      dismissRibbonPeek();
    }
  };
  host.addEventListener('mousedown', onMouseDown);
  host.addEventListener('click', onClick);

  // Double-clicking an active ribbon tab toggles the collapsed-tabs-only
  // ribbon mode — Excel-style shortcut.
  const onDoubleClick = (e: MouseEvent): void => {
    const target = e.target;
    if (!(target instanceof Element)) return;
    const tabBtn = target.closest<HTMLButtonElement>('[data-ribbon-tab]');
    if (!tabBtn) return;
    if (tabBtn.dataset.ribbonTab === 'file') return;
    e.preventDefault();
    setDisplayMode(isCollapsedMode() ? 'full' : 'tabsOnly');
  };
  host.addEventListener('dblclick', onDoubleClick);

  // Excel-style keyboard navigation across ribbon tabs: ArrowLeft / Right
  // cycle, Home / End jump to the first / last tab. Only fires when focus is
  // already on a tab so plain typing still works inside menus.
  const onKey = (e: KeyboardEvent): void => {
    const target = e.target;
    if (!(target instanceof HTMLElement)) return;
    const tabBtn = target.closest<HTMLButtonElement>('[data-ribbon-tab]');
    if (!tabBtn) return;
    const key = e.key;
    if (key !== 'ArrowLeft' && key !== 'ArrowRight' && key !== 'Home' && key !== 'End') return;
    // Hidden tabs are filtered so custom tab profiles can omit optional
    // add-in surfaces without leaving dead stops in the roving tabindex.
    const tabs = Array.from(host.querySelectorAll<HTMLButtonElement>('[data-ribbon-tab]')).filter(
      (btn) =>
        btn.offsetParent !== null && !btn.disabled && btn.getAttribute('aria-disabled') !== 'true',
    );
    if (tabs.length === 0) return;
    const currentIndex = tabs.indexOf(tabBtn);
    if (currentIndex < 0) return;
    let nextIndex = currentIndex;
    if (key === 'ArrowLeft') nextIndex = (currentIndex - 1 + tabs.length) % tabs.length;
    else if (key === 'ArrowRight') nextIndex = (currentIndex + 1) % tabs.length;
    else if (key === 'Home') nextIndex = 0;
    else if (key === 'End') nextIndex = tabs.length - 1;
    if (nextIndex === currentIndex) return;
    e.preventDefault();
    const nextTab = tabs[nextIndex];
    if (!nextTab) return;
    const nextId = nextTab.dataset.ribbonTab as RibbonTab | undefined;
    if (nextId) {
      activateRibbonTab(nextId, true);
      nextTab.focus();
    }
  };
  host.addEventListener('keydown', onKey);

  // Display-options menu keyboard navigation. ArrowDown from the toggle
  // opens the menu and focuses the first option; ArrowUp opens and focuses
  // the last; arrows inside the menu cycle; Home / End jump; Escape closes.
  const focusDisplayOption = (which: 'first' | 'last' | number): void => {
    const items = Array.from(
      host.querySelectorAll<HTMLButtonElement>('[data-ribbon-display-option]'),
    );
    if (items.length === 0) return;
    const idx =
      which === 'first'
        ? 0
        : which === 'last'
          ? items.length - 1
          : Math.max(0, Math.min(which, items.length - 1));
    items[idx]?.focus();
  };
  const onDisplayKey = (e: KeyboardEvent): void => {
    const target = e.target;
    if (!(target instanceof HTMLElement)) return;
    const toggleBtn = target.closest<HTMLElement>('[data-ribbon-toggle]');
    const optionBtn = target.closest<HTMLElement>('[data-ribbon-display-option]');
    if (toggleBtn && (e.key === 'ArrowDown' || e.key === 'ArrowUp')) {
      e.preventDefault();
      if (!displayMenuOpen) {
        displayMenuOpen = true;
        renderToolbar();
      }
      focusDisplayOption(e.key === 'ArrowDown' ? 'first' : 'last');
      return;
    }
    // Escape is handled by the document-level handler, which catches it from
    // anywhere the menu can be dismissed from.
    if (!optionBtn) return;
    const items = Array.from(
      host.querySelectorAll<HTMLButtonElement>('[data-ribbon-display-option]'),
    );
    const idx = items.indexOf(optionBtn as HTMLButtonElement);
    if (idx < 0) return;
    let next = idx;
    if (e.key === 'ArrowDown') next = (idx + 1) % items.length;
    else if (e.key === 'ArrowUp') next = (idx - 1 + items.length) % items.length;
    else if (e.key === 'Home') next = 0;
    else if (e.key === 'End') next = items.length - 1;
    else return;
    if (next === idx) return;
    e.preventDefault();
    items[next]?.focus();
  };
  host.addEventListener('keydown', onDisplayKey);

  // Ctrl+F1 toggles the collapsed-tabs-only ribbon mode regardless of focus
  // location — Excel-style global shortcut. Attached at document so the
  // sheet (or any other focus target) doesn't need to route the key.
  const onGlobalKey = (e: KeyboardEvent): void => {
    if (e.key === 'Escape' && backstageOpen) {
      e.preventDefault();
      closeBackstage();
      return;
    }
    if (e.key === 'Escape' && hasOpenStaticMenu()) {
      e.preventDefault();
      closeStaticRibbonMenus(host, undefined, true);
      return;
    }
    // The display menu is reachable by click, and re-rendering the ribbon on
    // open drops focus back to the document, so Escape has to be caught here
    // rather than on the ribbon itself.
    if (e.key === 'Escape' && displayMenuOpen) {
      e.preventDefault();
      const active = document.activeElement;
      const focusWasLoose = active === null || active === document.body || host.contains(active);
      displayMenuOpen = false;
      renderToolbar();
      if (focusWasLoose) host.querySelector<HTMLButtonElement>('[data-ribbon-toggle]')?.focus();
      return;
    }
    if (e.ctrlKey && e.key === 'F1') {
      e.preventDefault();
      setDisplayMode(isCollapsedMode() ? 'full' : 'tabsOnly');
      return;
    }
    if (displayMode === 'autoHide' && e.key === 'Alt' && !ribbonPeek) {
      e.preventDefault();
      ribbonPeek = true;
      renderToolbar();
      host.querySelector<HTMLButtonElement>(`[data-ribbon-tab="${activeTab}"]`)?.focus();
      return;
    }
    if (isCollapsedMode() && e.key === 'Escape' && ribbonPeek) {
      e.preventDefault();
      dismissRibbonPeek();
      focusActiveTab();
    }
  };
  document.addEventListener('keydown', onGlobalKey);

  // Clicking outside the ribbon while the display menu is open dismisses it
  // — Excel-style behaviour. Uses mousedown so the close happens before the
  // outside element's own click handler fires.
  const onDocumentMouseDown = (e: MouseEvent): void => {
    const shouldCloseStaticMenus = hasOpenStaticMenu();
    const shouldRenderDisplayState = displayMenuOpen || (isCollapsedMode() && ribbonPeek);
    if (!shouldRenderDisplayState && !shouldCloseStaticMenus) return;
    const target = e.target;
    if (!(target instanceof Element)) return;
    if (host.contains(target)) return;
    if (shouldCloseStaticMenus) closeStaticRibbonMenus(host);
    if (displayMenuOpen) displayMenuOpen = false;
    if (isCollapsedMode()) ribbonPeek = false;
    if (shouldRenderDisplayState) renderToolbar();
  };
  document.addEventListener('mousedown', onDocumentMouseDown);

  // Re-project active-state on any store mutation. When the toolbar is
  // mounted with a deferred getter we may not have an instance yet — track
  // the last subscribed instance and re-bind whenever the getter starts
  // returning a different one. The caller is expected to trigger at least one
  // `tb.rerender()` after `getInstance()` becomes non-null so the binding
  // attaches; in practice playground does this in its boot path.
  let unsubStore: (() => void) | null = null;
  let unsubPolicy: (() => void) | null = null;
  let unsubFunctionHistory: (() => void) | null = null;
  let unsubLocale: (() => void) | null = null;
  let subscribedInstance: SpreadsheetInstance | null = null;
  const refreshLocale = (current: SpreadsheetInstance): void => {
    if (langOverride !== undefined) return;
    const next = toolbarLangForLocale(current.i18n.locale);
    if (next === lang) return;
    lang = next;
    if (opts.text === undefined) text = toolbarText(lang);
    if (opts.menuText === undefined) menuText = toolbarMenuText(lang);
    displayOptionsText = ribbonDisplayText(lang);
    rebuildLocalizedDefaults();
    renderApi = createRibbonRenderer();
    renderToolbar();
  };
  const ensureStoreSubscription = (): void => {
    const current = getInstance();
    if (current === subscribedInstance) {
      syncInheritedPlatform(current);
      syncMacInkSubscription(current);
      return;
    }
    unsubStore?.();
    unsubPolicy?.();
    unsubFunctionHistory?.();
    unsubLocale?.();
    subscribedInstance = current;
    syncInheritedPlatform(current);
    syncMacInkSubscription(current);
    unsubLocale =
      current && langOverride === undefined
        ? current.i18n.subscribe(() => refreshLocale(current))
        : null;
    if (current) refreshLocale(current);
    syncContextualTabs(current, false);
    unsubStore =
      current?.store.subscribe(() => {
        const nextInstance = getInstance();
        const platformChanged = syncInheritedPlatform(nextInstance);
        const tabsChanged = syncContextualTabs(nextInstance, true);
        syncMacInkSubscription(nextInstance);
        projectFormatToolbar();
        if (platformChanged || tabsChanged) renderToolbar();
        else interactionPolicy.project();
      }) ?? null;
    unsubFunctionHistory = current
      ? subscribeRecentFunctions(current.store, projectFormatToolbar)
      : null;
    unsubPolicy = current
      ? (interactionControllerFor(current.store)?.subscribe(() => renderToolbar()) ?? null)
      : null;
  };
  ensureStoreSubscription();

  const rerender = (): void => {
    ensureStoreSubscription();
    renderToolbar();
  };

  rerender();

  return {
    host,
    get instance() {
      return getInstance();
    },
    rerender,
    applyCommand,
    focusActiveTab,
    setActiveTab: (tab) => activateRibbonTab(tab, false),
    getActiveTab: () => activeTab,
    setCollapsed: (next) => {
      setDisplayMode(next ? 'tabsOnly' : 'full');
    },
    getCollapsed: () => isCollapsedMode(),
    setDisplayMode,
    getDisplayMode: () => displayMode,
    setBackstageOpen: (next) => {
      if (next === backstageOpen) return;
      backstageOpen = next;
      opts.onBackstageOpenChange?.(next);
      renderToolbar();
    },
    getBackstageOpen: () => backstageOpen,
    setDisplayMenuOpen: (next) => {
      if (next === displayMenuOpen) return;
      displayMenuOpen = next;
      renderToolbar();
    },
    setFormulaBarVisible: (next) => {
      formulaBarVisible = next;
    },
    getFormulaBarVisible: () => formulaBarVisible,
    setTheme: (next) => {
      if (next === theme) return;
      theme = next;
      opts.onThemeChange?.(next);
      renderToolbar();
    },
    getTheme: () => theme,
    setBorderStyle: (next) => {
      borderStyle = next;
    },
    getBorderStyle: () => borderStyle,
    setBorderColor: (next) => {
      borderColor = next;
    },
    getBorderColor: () => borderColor,
    get dropdownsApi() {
      return dropdownsApi;
    },
    dispose: () => {
      screenTips.detach();
      ribbonFit.detach();
      unregisterOverlayOwner();
      if (previousLanguage === null) host.removeAttribute('lang');
      else host.setAttribute('lang', previousLanguage);
      if (previousPlatform === undefined) delete host.dataset.fcPlatform;
      else host.dataset.fcPlatform = previousPlatform;
      interactionPolicy.detach();
      host.removeEventListener('focusin', onRibbonFocus, true);
      host.removeEventListener('mousedown', onMouseDown);
      host.removeEventListener('click', onClick);
      host.removeEventListener('dblclick', onDoubleClick);
      host.removeEventListener('keydown', onKey);
      host.removeEventListener('keydown', onDisplayKey);
      document.removeEventListener('keydown', onGlobalKey);
      document.removeEventListener('mousedown', onDocumentMouseDown);
      dropdownsDelegation?.detach();
      dropdownsDelegation = null;
      borderMenuApi?.detach();
      borderMenuApi = null;
      dropdownsApi = null;
      unsubMacInk?.();
      unsubMacInk = null;
      subscribedMacInk = null;
      unsubStore?.();
      unsubPolicy?.();
      unsubPolicy = null;
      unsubFunctionHistory?.();
      unsubFunctionHistory = null;
      unsubLocale?.();
      unsubLocale = null;
      unsubStore = null;
      subscribedInstance = null;
      cancelOpenAppDialogs();
      host.replaceChildren();
    },
  };
}
