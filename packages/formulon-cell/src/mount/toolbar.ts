import { subscribeRecentFunctions } from '../commands/function-history.js';
import { registerOverlayOwner } from '../interact/overlay-portal.js';
import { projectDisabledState } from '../toolbar/menu-a11y.js';
import {
  defaultFunctionUnavailableReason,
  projectMacFunctionCategoryMenus,
  projectMacRecentMenu,
} from '../toolbar/ribbon/mac/menus.js';

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

import { canExecuteBuiltIn } from '../commands/built-in-command-policy.js';
import { withSelectionFormatOrigin } from '../commands/format.js';
import type { FunctionCatalogReader } from '../commands/function-categories.js';
import { recordRepeatableFormatChange } from '../commands/history.js';
import { interactionControllerFor } from '../commands/interaction-controller.js';
import { resolveSpreadsheetPlatform, type SpreadsheetPlatform } from '../extensions/ui-options.js';
import { ensureMacInk, getMacInk, type MacInkController } from '../interact/mac-ink.js';
import type { CellBorderStyle } from '../store/types.js';
import { cancelOpenAppDialogs } from '../toolbar/dialogs/shell.js';
import { ribbonDisplayText, type ToolbarMenuText, toolbarMenuText } from '../toolbar/menu-text.js';
import { isRibbonMenuFirstCommand, RIBBON_BORDERS_MENU_ID } from '../toolbar/ribbon/activation.js';
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
  createDynamicDropdowns,
  type DynamicDropdownsApi,
  type DynamicDropdownsCtx,
  ribbonDropdownMenuIdForCommand,
} from '../toolbar/ribbon/dynamic-dropdowns.js';
import { toolbarLangForLocale } from '../toolbar/ribbon/mac/locale.js';
import { MAC_FORMULAS_MORE_MENU_ID, projectMacRibbonState } from '../toolbar/ribbon/mac/model.js';
import {
  createRenderRibbon,
  type RibbonDisplayMode,
  type RibbonMenus,
  type RibbonRenderHelpers,
} from '../toolbar/ribbon/render-ribbon.js';
import { projectActiveState, RIBBON_ACTIVE_COMMANDS } from '../toolbar/ribbon-active-state.js';
import {
  EXCEL365_MAC_RIBBON_TABS,
  RIBBON_TABS,
  type RibbonTab,
  type ToolbarLang,
  type ToolbarText,
  toolbarText,
} from '../toolbar/ribbon-model.js';
import { createDefaultDynamicDropdownsCtx } from './dynamic-dropdowns-defaults.js';
import {
  createDefaultRibbonHelpers,
  createDefaultRibbonHooks,
  createDefaultRibbonMenus,
} from './toolbar-defaults.js';
import type { SpreadsheetInstance } from './types.js';

type UiTheme = 'paper' | 'ink' | 'contrast';

export type { RibbonDisplayMode } from '../toolbar/ribbon/render-ribbon.js';

const DEFAULT_BORDER_STYLE: CellBorderStyle = 'thin';
const DEFAULT_BORDER_COLOR = '#000000';

const projectDefaultRibbonActiveState = (
  host: HTMLElement,
  instance: SpreadsheetInstance | null,
): void => {
  if (!instance) return;
  const active = projectActiveState(instance);
  for (const [command, key] of RIBBON_ACTIVE_COMMANDS) {
    const button = host.querySelector<HTMLButtonElement>(`[data-ribbon-command="${command}"]`);
    if (!button) continue;
    let pressed = Boolean(active[key]);
    if (command === 'viewNormal') pressed = active.workbookView === 'normal';
    else if (command === 'viewPageLayout') pressed = active.workbookView === 'pageLayout';
    else if (command === 'viewPageBreakPreview')
      pressed = active.workbookView === 'pageBreakPreview';
    button.classList.toggle('fc-tb__rb--active', pressed);
    button.setAttribute('aria-pressed', pressed ? 'true' : 'false');
  }

  const sheetBackground = host.querySelector<HTMLButtonElement>(
    '[data-ribbon-command="sheetBackground"]',
  );
  if (sheetBackground) {
    const state = instance.store.getState();
    const hasBackground = state.ui.sheetBackgroundImages.has(state.data.sheetIndex);
    const label = hasBackground
      ? instance.i18n.strings.ribbonMenu.sheetBackgroundClear
      : instance.i18n.strings.ribbon.background;
    sheetBackground.title = label;
    sheetBackground.setAttribute('aria-label', label);
    const labelEl = sheetBackground.querySelector('span');
    if (labelEl) labelEl.textContent = label;
  }
};

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

  /** Backstage view factory. When the user opens the "File" tab the toolbar
   *  replaces its body with the element this returns. Without it, the file
   *  tab simply switches to an empty panel. */
  createBackstageView?: () => HTMLElement;

  /** Initial UI state. */
  activeTab?: RibbonTab;
  /** Explicit ribbon tab surface. Pass `EXCEL365_STANDARD_RIBBON_TABS` for
   *  the Microsoft 365 baseline and append optional add-in/automation tabs
   *  only when the host has those surfaces wired. Defaults to the historical
   *  full tab set for backwards compatibility. */
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

  /** Lifecycle callbacks fired after the matching internal state change. */
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
  setBackstageOpen(open: boolean): void;
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

/** Split buttons whose menu is the entry point for hosts that own the actions
 *  behind it (`applyScriptAction` / `applyAddInAction`). Core classifies them
 *  as primary-action splits so a standalone ribbon fires its built-in dialog;
 *  a host that supplies the menu actions wants the face click to open the
 *  menu instead. */
export const RIBBON_HOST_MENU_FIRST_COMMANDS: ReadonlySet<string> = new Set(['script', 'addIn']);

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
  let defaultRibbonTabs =
    opts.ribbonTabs ?? (platform === 'mac' ? EXCEL365_MAC_RIBBON_TABS : undefined);

  // Probe once at mount-time for language inference and the initial subscribe.
  // The toolbar continues to work if the probe returns null (deferred mount);
  // in that case lang falls back to opts.lang or 'ja' and the store subscription
  // is attached lazily on the first call to `attachStoreSubscription`.
  const unregisterOverlayOwner = registerOverlayOwner(host, () => getInstance()?.host ?? null);

  const lang: ToolbarLang =
    opts.lang ?? (initialInstance ? toolbarLangForLocale(initialInstance.i18n.locale) : 'ja');
  const text = opts.text ?? toolbarText(lang);
  const menuText = opts.menuText ?? toolbarMenuText(lang);
  const displayOptionsText = ribbonDisplayText(lang);

  let activeTab: RibbonTab = opts.activeTab ?? 'home';
  let displayMode: RibbonDisplayMode =
    opts.ribbonDisplayMode ?? (opts.collapsed ? 'tabsOnly' : 'full');
  let ribbonPeek = false;
  let backstageOpen = false;
  let displayMenuOpen = false;
  let formulaBarVisible = opts.formulaBarVisible ?? true;
  let theme: UiTheme = opts.theme ?? 'paper';
  let borderStyle: CellBorderStyle = opts.borderStyle ?? DEFAULT_BORDER_STYLE;
  let borderColor = opts.borderColor ?? DEFAULT_BORDER_COLOR;

  const rehomeActiveTabForPlatform = (): void => {
    const available = opts.ribbonTabs ?? defaultRibbonTabs ?? RIBBON_TABS;
    if (available.includes(activeTab)) return;
    activeTab = available[0] ?? 'home';
  };
  rehomeActiveTabForPlatform();

  const focusSheet =
    opts.focusSheet ??
    ((): void => {
      getInstance()?.host.focus();
    });
  const refreshCells = opts.refreshCells ?? ((): void => undefined);
  const refreshZoom = opts.refreshZoom ?? ((): void => undefined);
  const projectFormatToolbar = (): void => {
    projectDefaultRibbonActiveState(host, getInstance());
    const instance = getInstance();
    if (instance && platform === 'mac') {
      const reportedNames =
        typeof instance.workbook.functionNames === 'function'
          ? instance.workbook.functionNames()
          : null;
      const liveNames = reportedNames === null ? null : new Set(reportedNames);
      const availabilityContext = {
        reader:
          reportedNames === null
            ? undefined
            : (instance.workbook as unknown as FunctionCatalogReader),
        unavailableReason:
          instance.i18n.strings.fxDialog.functionUnavailable ??
          defaultFunctionUnavailableReason(lang),
      };
      projectMacRecentMenu(host, instance.store, lang, liveNames, availabilityContext);
      projectMacFunctionCategoryMenus(host, instance.store, lang, liveNames, availabilityContext);
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

  // Defaults are derived once at mount time. They close over the live
  // `borderStyle/borderColor` via getters so the borders submenu always
  // picks the most recent value. The host may still override individual
  // helpers/menus/hooks by spreading on top.
  const defaultsInstance = initialInstance;
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

  const mergedHelpers: RibbonRenderHelpers = {
    ...(defaultHelpers ?? {}),
    ...(opts.helpers ?? {}),
  } as RibbonRenderHelpers;
  const mergedMenus: RibbonMenus = { ...defaultMenus, ...(opts.menus ?? {}) };
  // Hooks merge: per-category shallow merge so the host can extend (not
  // replace) any single group. Categories the host doesn't mention keep the
  // defaults; categories it does mention spread on top of the default ones.
  const mergedHooks: RibbonHooks = { ...defaultHooks };
  if (opts.hooks) {
    for (const key of Object.keys(opts.hooks) as (keyof RibbonHooks)[]) {
      const hostGroup = opts.hooks[key];
      if (!hostGroup) continue;
      const defaultGroup = mergedHooks[key];
      // biome-ignore lint/suspicious/noExplicitAny: index access onto union
      (mergedHooks as any)[key] = { ...(defaultGroup ?? {}), ...hostGroup };
    }
  }

  // Auto-wire the default dynamic-dropdowns click delegator when the host
  // opts in. The handler is attached to `document` (matching the playground
  // wiring) so clicks anywhere inside an open `.fc-tb__menu` reach the
  // dispatcher. We capture the unsubscribe and undo it in dispose so the
  // listener does not leak after re-mounts.
  let dynamicDropdownClickHandler: ((event: MouseEvent) => void) | null = null;
  let dynamicDropdownPointerDownHandler: ((event: MouseEvent) => void) | null = null;
  let dynamicDropdownFocusHandler: ((event: FocusEvent) => void) | null = null;
  let dynamicDropdownHoverHandler: ((event: MouseEvent) => void) | null = null;
  let dynamicDropdownKeyHandler: ((event: KeyboardEvent) => void) | null = null;
  let dropdownsApi: DynamicDropdownsApi | null = null;
  if (opts.dynamicDropdowns) {
    const hostOverrides: Partial<DynamicDropdownsCtx> | (() => Partial<DynamicDropdownsCtx>) =
      opts.dynamicDropdowns === true ? {} : opts.dynamicDropdowns;
    const withToolbarDropdownOverrides = (
      overrides: Partial<DynamicDropdownsCtx>,
    ): Partial<DynamicDropdownsCtx> => ({
      closeBorderMenu: (restoreFocus?: boolean) => {
        borderMenuApi?.closeBorderMenu(restoreFocus);
      },
      ...overrides,
    });
    const overridesOpt: Partial<DynamicDropdownsCtx> | (() => Partial<DynamicDropdownsCtx>) =
      typeof hostOverrides === 'function'
        ? () => withToolbarDropdownOverrides(hostOverrides())
        : withToolbarDropdownOverrides(hostOverrides);
    // `createDefaultDynamicDropdownsCtx` uses the `@libraz/formulon-cell`
    // self-import for `SpreadsheetInstance` (matching `dynamic-dropdowns.ts`)
    // so its parameter type resolves to dist. This file imports the
    // src-side declaration via `./types.js`, so the two structurally
    // identical declarations need one bridge cast.
    //
    // When `defaultsInstance` is null (deferred-mount hosts like the
    // playground), the built-in base handlers stay unreachable as long as
    // the host overrides every handler it dispatches. The override getter
    // (recommended for deferred hosts) captures the live instance via its
    // own closure so it can hand back the real `inst` once mounted.
    const dropdownsCtx = createDefaultDynamicDropdownsCtx(
      (defaultsInstance ?? ({} as SpreadsheetInstance)) as unknown as Parameters<
        typeof createDefaultDynamicDropdownsCtx
      >[0],
      {
        focusSheet,
        projectFormatToolbar,
        refreshCells,
        overrides: overridesOpt,
      },
    );
    dropdownsApi = createDynamicDropdowns(dropdownsCtx);
    dynamicDropdownClickHandler = (event: MouseEvent): void => {
      const current = getInstance();
      const target = event.target instanceof Element ? event.target : null;
      const menu = target?.closest<HTMLElement>('.fc-tb__menu') ?? null;
      if (
        current &&
        interactionControllerFor(current.store)?.policy !== undefined &&
        !target?.closest<HTMLElement>(
          `#${MAC_FORMULAS_MORE_MENU_ID} [data-function-category-submenu]`,
        )
      )
        return;
      if (dropdownsApi?.dynamicRibbonDropdownClick(event) && menu?.hidden) dismissRibbonPeek();
    };
    dynamicDropdownPointerDownHandler = (event: MouseEvent): void => {
      dropdownsApi?.dynamicRibbonDropdownPointerDown(event);
    };
    dynamicDropdownFocusHandler = (event: FocusEvent): void => {
      dropdownsApi?.dynamicRibbonDropdownFocusIn(event);
    };
    dynamicDropdownHoverHandler = (event: MouseEvent): void => {
      dropdownsApi?.dynamicRibbonDropdownHover(event);
    };
    dynamicDropdownKeyHandler = (event: KeyboardEvent): void => {
      const current = getInstance();
      const target = event.target instanceof Element ? event.target : null;
      const menu = target?.closest<HTMLElement>('.fc-tb__menu') ?? null;
      if (
        current &&
        interactionControllerFor(current.store)?.policy !== undefined &&
        !target?.closest<HTMLElement>(`#${MAC_FORMULAS_MORE_MENU_ID}`)
      )
        return;
      if (
        dropdownsApi?.dynamicRibbonDropdownKeydown(event) &&
        menu?.hidden &&
        (event.key === 'Enter' || event.key === ' ')
      )
        dismissRibbonPeek();
    };
    document.addEventListener('click', dynamicDropdownClickHandler);
    document.addEventListener('mousedown', dynamicDropdownPointerDownHandler, true);
    document.addEventListener('focusin', dynamicDropdownFocusHandler);
    document.addEventListener('mouseover', dynamicDropdownHoverHandler);
    document.addEventListener('keydown', dynamicDropdownKeyHandler);
  }

  const renderApi = createRenderRibbon({
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
    },
    helpers: mergedHelpers,
    menus: mergedMenus,
    createBackstageView: opts.createBackstageView ?? (() => document.createElement('div')),
    projectFormatToolbar,
  });

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
    borderMenuApi?.detach();
    borderMenuApi = null;
    renderApi.renderRibbon();
    wireBorderMenu();
    projectInteractionPolicy();
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
      defaultRibbonTabs = next === 'mac' ? EXCEL365_MAC_RIBBON_TABS : undefined;
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

  const projectInteractionPolicy = (): void => {
    const current = getInstance();
    if (!current || interactionControllerFor(current.store)?.policy === undefined) return;
    dropdownsApi?.closeAllDynamicRibbonDropdowns();
    for (const button of host.querySelectorAll<HTMLButtonElement>('[data-ribbon-command]')) {
      // Intrinsic engine availability has priority over embedding policy. The
      // menu projector already supplied its localized reason; a policy tick
      // must not replace it with a generic denial message.
      if (button.dataset.functionUnavailable === 'true') continue;
      const decision = canExecuteBuiltIn(
        current.store,
        button.dataset.ribbonCommand ?? '',
        'ribbon',
      );
      if (decision.allowed) continue;
      projectDisabledState(button, true, decision.reason ?? decision.code);
    }
    for (const input of host.querySelectorAll<HTMLInputElement | HTMLSelectElement>(
      'input, select',
    )) {
      projectDisabledState(input, true, 'Unavailable in restricted embedding');
    }
  };

  // Capture before directly-bound control/dropdown handlers, including hosts
  // using a separately mounted toolbar. Unknown restricted routes fail closed.
  const guardInteraction = (event: Event): void => {
    const current = getInstance();
    if (!current || interactionControllerFor(current.store)?.policy === undefined) return;
    const target = event.target;
    if (!(target instanceof Element)) return;
    const command = target.closest<HTMLElement>('[data-ribbon-command]');
    const inMenu = target.closest('.fc-tb__menu');
    const isInput = target.closest('input, select, textarea');
    if (
      event.type === 'click' &&
      target.closest(`#${MAC_FORMULAS_MORE_MENU_ID} [data-function-category-submenu]`)
    )
      return;
    if (
      command &&
      canExecuteBuiltIn(current.store, command.dataset.ribbonCommand ?? '', 'ribbon').allowed &&
      !isInput
    )
      return;
    if (!command && !inMenu && !isInput) return;
    event.preventDefault();
    event.stopImmediatePropagation();
  };

  host.addEventListener('click', guardInteraction, true);
  host.addEventListener('change', guardInteraction, true);
  host.addEventListener('input', guardInteraction, true);

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

  const closeStaticRibbonMenus = (except?: HTMLElement, restoreFocus = false): void => {
    let restoreTarget: HTMLButtonElement | null = null;
    for (const menu of host.querySelectorAll<HTMLDivElement>('.fc-tb__menu')) {
      if (menu === except || menu.hidden) continue;
      menu.hidden = true;
      const button = host.querySelector<HTMLButtonElement>(`[data-ribbon-menu-id="${menu.id}"]`);
      button?.setAttribute('aria-expanded', 'false');
      restoreTarget ??= button;
    }
    for (const panel of host.querySelectorAll<HTMLElement>('[data-function-category-panel]')) {
      panel.hidden = true;
    }
    for (const trigger of host.querySelectorAll<HTMLElement>('[data-function-category-submenu]')) {
      trigger.classList.remove('fc-tb__menu-item--active');
      trigger.setAttribute('aria-expanded', 'false');
    }
    if (restoreFocus) restoreTarget?.focus();
  };

  const projectRibbonTabs = (): void => {
    const shell = host.querySelector<HTMLElement>('.fc-tb__ribbon-shell');
    if (!shell) return;
    const peek = isCollapsedMode() && ribbonPeek;
    const collapsed = isCollapsedMode() && !peek;
    shell.classList.toggle('fc-tb__ribbon-shell--peek', peek);
    shell.classList.toggle('fc-tb__ribbon-shell--autoHidePeek', displayMode === 'autoHide' && peek);
    shell.classList.toggle('fc-tb__ribbon-shell--collapsed', collapsed);
    if (peek) shell.dataset.ribbonPeek = 'true';
    else delete shell.dataset.ribbonPeek;
    if (displayMode === 'autoHide' && peek) shell.dataset.ribbonAutoHidePeek = 'true';
    else delete shell.dataset.ribbonAutoHidePeek;
    const tabs = shell.querySelector<HTMLElement>('.fc-tb__ribbon-tabs');
    if (tabs) tabs.dataset.ribbonCollapsed = collapsed ? 'true' : 'false';
    for (const button of shell.querySelectorAll<HTMLButtonElement>('[data-ribbon-tab]')) {
      const selected = button.dataset.ribbonTab === activeTab;
      button.classList.toggle('fc-tb__ribbon-tab--active', selected);
      button.setAttribute('aria-selected', String(selected));
      button.tabIndex = selected ? 0 : -1;
    }
    for (const panel of shell.querySelectorAll<HTMLElement>('[data-ribbon-panel]')) {
      panel.hidden = panel.dataset.ribbonPanel !== activeTab;
    }
  };

  const dismissRibbonPeek = (): void => {
    if (!ribbonPeek) return;
    ribbonPeek = false;
    dropdownsApi?.closeAllDynamicRibbonDropdowns();
    closeStaticRibbonMenus();
    projectRibbonTabs();
  };

  const activateRibbonTab = (tab: RibbonTab, reveal: boolean): void => {
    const button = host.querySelector<HTMLButtonElement>(`[data-ribbon-tab="${tab}"]`);
    if (!button || button.disabled || button.getAttribute('aria-disabled') === 'true') return;
    dropdownsApi?.closeAllDynamicRibbonDropdowns();
    closeStaticRibbonMenus();
    const changed = tab !== activeTab;
    activeTab = tab;
    if (reveal && tab !== 'file' && isCollapsedMode()) ribbonPeek = true;
    projectRibbonTabs();
    if (changed) opts.onTabChange?.(tab);
  };

  const hasOpenStaticRibbonMenu = (): boolean =>
    !dropdownsApi &&
    Array.from(host.querySelectorAll<HTMLDivElement>('.fc-tb__menu')).some((menu) => !menu.hidden);

  const onClick = (e: MouseEvent): void => {
    const target = e.target;
    if (!(target instanceof Element)) return;

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
      if (document.activeElement !== cmdBtn) cmdBtn.focus({ preventScroll: true });
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
        const submenu = cmdBtn.nextElementSibling;
        if (submenu instanceof HTMLDivElement && submenu.classList.contains('fc-tb__menu')) {
          const wasOpen = !submenu.hidden;
          closeStaticRibbonMenus(submenu);
          submenu.hidden = wasOpen;
          cmdBtn.setAttribute('aria-expanded', wasOpen ? 'false' : 'true');
        }
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
        const submenu = cmdBtn.nextElementSibling;
        if (submenu instanceof HTMLDivElement && submenu.classList.contains('fc-tb__menu')) {
          const wasOpen = !submenu.hidden;
          closeStaticRibbonMenus(submenu);
          submenu.hidden = wasOpen;
          cmdBtn.setAttribute('aria-expanded', wasOpen ? 'false' : 'true');
          return;
        }
      }
      if (cmdBtn.closest('.fc-tb__menu--mac')) {
        const menu = cmdBtn.closest<HTMLElement>('.fc-tb__menu--mac');
        const spec = menu ? dropdownsApi?.dynamicDropdownSpecForMenu(menu) : null;
        if (spec) dropdownsApi?.closeDynamicRibbonDropdown(spec, true);
        closeStaticRibbonMenus();
      }
      applyCommand(id);
      dismissRibbonPeek();
    }
  };
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
    if (e.key === 'Escape' && hasOpenStaticRibbonMenu()) {
      e.preventDefault();
      closeStaticRibbonMenus(undefined, true);
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
    const shouldCloseStaticMenus = hasOpenStaticRibbonMenu();
    const shouldRenderDisplayState = displayMenuOpen || (isCollapsedMode() && ribbonPeek);
    if (!shouldRenderDisplayState && !shouldCloseStaticMenus) return;
    const target = e.target;
    if (!(target instanceof Element)) return;
    if (host.contains(target)) return;
    if (shouldCloseStaticMenus) closeStaticRibbonMenus();
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
  let subscribedInstance: SpreadsheetInstance | null = null;
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
    subscribedInstance = current;
    syncInheritedPlatform(current);
    syncMacInkSubscription(current);
    unsubStore =
      current?.store.subscribe(() => {
        const nextInstance = getInstance();
        const platformChanged = syncInheritedPlatform(nextInstance);
        syncMacInkSubscription(nextInstance);
        projectFormatToolbar();
        if (platformChanged) renderToolbar();
        else projectInteractionPolicy();
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
      unregisterOverlayOwner();
      if (previousPlatform === undefined) delete host.dataset.fcPlatform;
      else host.dataset.fcPlatform = previousPlatform;
      host.removeEventListener('click', guardInteraction, true);
      host.removeEventListener('change', guardInteraction, true);
      host.removeEventListener('input', guardInteraction, true);
      host.removeEventListener('click', onClick);
      host.removeEventListener('dblclick', onDoubleClick);
      host.removeEventListener('keydown', onKey);
      host.removeEventListener('keydown', onDisplayKey);
      document.removeEventListener('keydown', onGlobalKey);
      document.removeEventListener('mousedown', onDocumentMouseDown);
      if (dynamicDropdownClickHandler) {
        document.removeEventListener('click', dynamicDropdownClickHandler);
        dynamicDropdownClickHandler = null;
      }
      if (dynamicDropdownPointerDownHandler) {
        document.removeEventListener('mousedown', dynamicDropdownPointerDownHandler, true);
        dynamicDropdownPointerDownHandler = null;
      }
      if (dynamicDropdownFocusHandler) {
        document.removeEventListener('focusin', dynamicDropdownFocusHandler);
        dynamicDropdownFocusHandler = null;
      }
      if (dynamicDropdownHoverHandler) {
        document.removeEventListener('mouseover', dynamicDropdownHoverHandler);
        dynamicDropdownHoverHandler = null;
      }
      if (dynamicDropdownKeyHandler) {
        document.removeEventListener('keydown', dynamicDropdownKeyHandler);
        dynamicDropdownKeyHandler = null;
      }
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
      unsubStore = null;
      subscribedInstance = null;
      cancelOpenAppDialogs();
      host.replaceChildren();
    },
  };
}
