import type { SpreadsheetPlatform } from '../extensions/ui-options.js';
import type { CellBorderStyle } from '../store/types.js';
import type { ToolbarMenuText } from '../toolbar/menu-text.js';
import type { RibbonHooks, RibbonRuntime } from '../toolbar/ribbon/apply-ribbon-command.js';
import type {
  DynamicDropdownsApi,
  DynamicDropdownsCtx,
  UiTheme,
} from '../toolbar/ribbon/dynamic-dropdowns.js';
import type {
  RibbonDisplayMode,
  RibbonMenus,
  RibbonRenderHelpers,
} from '../toolbar/ribbon/render-ribbon.js';
import type { RibbonTab, ToolbarLang, ToolbarText } from '../toolbar/ribbon-model.js';
import type { SpreadsheetInstance } from './types.js';

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

export interface ToolbarInstance {
  readonly host: HTMLElement;
  /** The spreadsheet instance the toolbar was mounted on. */
  readonly instance: SpreadsheetInstance;
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
