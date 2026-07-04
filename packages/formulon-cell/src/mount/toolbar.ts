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
//    app-specific glue (illustrations, custom dialog flows, …). Phase 2 will
//    pull more of this inside, but the boundary at v0.1 keeps consumers in
//    control of their own surface.
//
// The toolbar listens to `instance.store.subscribe()` so it can re-project
// active-state (bold/italic/etc.) on every selection or format change. It
// does NOT re-render the whole ribbon on every change — only the active-state
// projection runs in the hot path. Tab switches and similar topology changes
// go through `renderRibbon()` once.

import type { CellBorderStyle } from '../store/types.js';
import { cancelOpenAppDialogs } from '../toolbar/dialogs/shell.js';
import { ribbonDisplayText, type ToolbarMenuText, toolbarMenuText } from '../toolbar/menu-text.js';
import {
  RIBBON_BORDERS_MENU_ID,
  RIBBON_MENU_FIRST_COMMANDS,
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
  createDynamicDropdowns,
  type DynamicDropdownsApi,
  type DynamicDropdownsCtx,
} from '../toolbar/ribbon/dynamic-dropdowns.js';
import {
  createRenderRibbon,
  type RibbonDisplayMode,
  type RibbonMenus,
  type RibbonRenderHelpers,
} from '../toolbar/ribbon/render-ribbon.js';
import { projectActiveState, RIBBON_ACTIVE_COMMANDS } from '../toolbar/ribbon-active-state.js';
import {
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
    button.classList.toggle('demo__rb--active', pressed);
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

const defaultApplyRibbonFormat =
  (getInstance: () => SpreadsheetInstance | null) =>
  (fn: RibbonFormatMutator): void => {
    const inst = getInstance();
    if (!inst) return;
    fn(inst.store.getState(), inst.store);
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

  // Probe once at mount-time for language inference and the initial subscribe.
  // The toolbar continues to work if the probe returns null (deferred mount);
  // in that case lang falls back to opts.lang or 'ja' and the store subscription
  // is attached lazily on the first call to `attachStoreSubscription`.
  const initialInstance = getInstance();

  const lang: ToolbarLang = opts.lang ?? (initialInstance?.i18n.locale === 'en' ? 'en' : 'ja');
  const text = opts.text ?? toolbarText(lang);
  const menuText = opts.menuText ?? toolbarMenuText(lang);
  const displayOptionsText = ribbonDisplayText(lang);

  let activeTab: RibbonTab = opts.activeTab ?? 'home';
  let displayMode: RibbonDisplayMode =
    opts.ribbonDisplayMode ?? (opts.collapsed ? 'tabsOnly' : 'full');
  let autoHidePeek = false;
  let backstageOpen = false;
  let displayMenuOpen = false;
  let formulaBarVisible = opts.formulaBarVisible ?? true;
  let theme: UiTheme = opts.theme ?? 'paper';
  let borderStyle: CellBorderStyle = opts.borderStyle ?? DEFAULT_BORDER_STYLE;
  let borderColor = opts.borderColor ?? DEFAULT_BORDER_COLOR;

  const focusSheet =
    opts.focusSheet ??
    ((): void => {
      getInstance()?.host.focus();
    });
  const refreshCells = opts.refreshCells ?? ((): void => undefined);
  const refreshZoom = opts.refreshZoom ?? ((): void => undefined);
  const projectFormatToolbar = (): void => {
    projectDefaultRibbonActiveState(host, getInstance());
    opts.projectFormatToolbar?.();
  };
  const showMessage = opts.showMessage ?? ((): void => undefined);
  const applyRibbonFormat = opts.applyRibbonFormat ?? defaultApplyRibbonFormat(getInstance);
  const isCollapsedMode = (): boolean => displayMode === 'tabsOnly' || displayMode === 'autoHide';
  let borderMenuApi: BorderMenuApi | null = null;
  const setDisplayMode = (next: RibbonDisplayMode): void => {
    if (next === displayMode) return;
    const wasCollapsed = isCollapsedMode();
    displayMode = next;
    autoHidePeek = false;
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
  // wiring) so clicks anywhere inside an open `.app__menu` reach the
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
      dropdownsApi?.dynamicRibbonDropdownClick(event);
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
      dropdownsApi?.dynamicRibbonDropdownKeydown(event);
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
    ribbonTabs: opts.ribbonTabs,
    ribbonRoot: host,
    state: {
      getActiveTab: () => activeTab,
      getCollapsed: () => isCollapsedMode(),
      getDisplayMode: () => displayMode,
      getAutoHidePeek: () => autoHidePeek,
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
  };

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
    for (const menu of host.querySelectorAll<HTMLDivElement>('.app__menu')) {
      if (menu === except || menu.hidden) continue;
      menu.hidden = true;
      const button = host.querySelector<HTMLButtonElement>(`[data-ribbon-menu-id="${menu.id}"]`);
      button?.setAttribute('aria-expanded', 'false');
      restoreTarget ??= button;
    }
    if (restoreFocus) restoreTarget?.focus();
  };

  const hasOpenStaticRibbonMenu = (): boolean =>
    !dropdownsApi &&
    Array.from(host.querySelectorAll<HTMLDivElement>('.app__menu')).some((menu) => !menu.hidden);

  const onClick = (e: MouseEvent): void => {
    const target = e.target;
    if (!(target instanceof Element)) return;

    const tabBtn = target.closest<HTMLButtonElement>('[data-ribbon-tab]');
    if (tabBtn) {
      const tab = tabBtn.dataset.ribbonTab as RibbonTab | undefined;
      if (tab && tab !== activeTab) {
        activeTab = tab;
        opts.onTabChange?.(tab);
        renderToolbar();
      }
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

    if (opts.commandDelegation === false) return;
    const cmdBtn = target.closest<HTMLButtonElement>('[data-ribbon-command]');
    if (cmdBtn?.dataset.ribbonCommand) {
      const id = cmdBtn.dataset.ribbonCommand;
      if (opts.interceptCommand?.(id, cmdBtn, e)) return;
      // Fallback dropdown behaviour: if the button has a sibling submenu
      // attached via render-ribbon's `tools.appendChild(submenu())`, toggle
      // it. Split buttons with a primary face action skip this so
      // applyRibbonCommand can fire their primary handler — the
      // chevron-vs-main split lives in the host.
      if (RIBBON_MENU_FIRST_COMMANDS.has(id)) {
        const menuId = cmdBtn.dataset.ribbonMenuId;
        if (dropdownsApi && menuId) {
          dropdownsApi.openDynamicRibbonDropdown({ command: id, menuId }, cmdBtn);
          return;
        }
        const submenu = cmdBtn.nextElementSibling;
        if (submenu instanceof HTMLDivElement && submenu.classList.contains('app__menu')) {
          const wasOpen = !submenu.hidden;
          closeStaticRibbonMenus(submenu);
          submenu.hidden = wasOpen;
          cmdBtn.setAttribute('aria-expanded', wasOpen ? 'false' : 'true');
          return;
        }
      }
      applyCommand(id);
      if (displayMode === 'autoHide' && autoHidePeek) {
        autoHidePeek = false;
        renderToolbar();
      }
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
      (btn) => btn.offsetParent !== null,
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
    if (nextId && nextId !== activeTab) {
      activeTab = nextId;
      opts.onTabChange?.(nextId);
      renderToolbar();
      // After re-render, refetch the tab from the freshly-rendered DOM and
      // restore focus + roving-tabindex.
      const focusTarget = host.querySelector<HTMLButtonElement>(`[data-ribbon-tab="${nextId}"]`);
      focusTarget?.focus();
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
    else if (e.key === 'Escape') {
      e.preventDefault();
      displayMenuOpen = false;
      renderToolbar();
      const reopenedToggle = host.querySelector<HTMLButtonElement>('[data-ribbon-toggle]');
      reopenedToggle?.focus();
      return;
    } else return;
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
    if (e.ctrlKey && e.key === 'F1') {
      e.preventDefault();
      setDisplayMode(isCollapsedMode() ? 'full' : 'tabsOnly');
      return;
    }
    if (displayMode === 'autoHide' && e.key === 'Alt' && !autoHidePeek) {
      e.preventDefault();
      autoHidePeek = true;
      renderToolbar();
      host.querySelector<HTMLButtonElement>(`[data-ribbon-tab="${activeTab}"]`)?.focus();
      return;
    }
    if (displayMode === 'autoHide' && e.key === 'Escape' && autoHidePeek) {
      e.preventDefault();
      autoHidePeek = false;
      renderToolbar();
    }
  };
  document.addEventListener('keydown', onGlobalKey);

  // Clicking outside the ribbon while the display menu is open dismisses it
  // — Excel-style behaviour. Uses mousedown so the close happens before the
  // outside element's own click handler fires.
  const onDocumentMouseDown = (e: MouseEvent): void => {
    const shouldCloseStaticMenus = hasOpenStaticRibbonMenu();
    const shouldRenderDisplayState =
      displayMenuOpen || (displayMode === 'autoHide' && autoHidePeek);
    if (!shouldRenderDisplayState && !shouldCloseStaticMenus) return;
    const target = e.target;
    if (!(target instanceof Element)) return;
    if (host.contains(target)) return;
    if (shouldCloseStaticMenus) closeStaticRibbonMenus();
    if (displayMenuOpen) displayMenuOpen = false;
    if (displayMode === 'autoHide') autoHidePeek = false;
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
  let subscribedInstance: SpreadsheetInstance | null = null;
  const ensureStoreSubscription = (): void => {
    const current = getInstance();
    if (current === subscribedInstance) return;
    unsubStore?.();
    subscribedInstance = current;
    unsubStore = current?.store.subscribe(() => projectFormatToolbar()) ?? null;
  };
  ensureStoreSubscription();

  const rerender = (): void => {
    ensureStoreSubscription();
    renderToolbar();
  };

  rerender();
  projectFormatToolbar();

  return {
    host,
    get instance() {
      return getInstance();
    },
    rerender,
    applyCommand,
    focusActiveTab,
    setActiveTab: (tab) => {
      if (tab === activeTab) return;
      activeTab = tab;
      opts.onTabChange?.(tab);
      renderToolbar();
    },
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
      unsubStore?.();
      unsubStore = null;
      subscribedInstance = null;
      cancelOpenAppDialogs();
      host.replaceChildren();
    },
  };
}
