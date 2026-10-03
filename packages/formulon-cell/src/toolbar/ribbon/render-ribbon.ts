// Ribbon DOM renderer. Owns the tab/panel layout, the per-command button
// rendering, the split-button chevron, the display-mode toggle, and the
// backstage hand-off. State (active tab, collapsed flag, backstage flag,
// display-menu flag) stays in the host; this factory reads them through
// getters so successive renders see the latest values.
//
// The 40-some submenu factories used to be passed as flat fields; they are
// now bundled into a single `menus` map so consumers can spread an
// auto-generated object. Each entry receives the command id as its argument
// so factories that vary by panel (e.g. autosum-home vs. autosum-formulas)
// can branch without needing dedicated wrapper props.

import type { FeatureFlags } from '../../extensions/index.js';
import { clampPanelToViewport, panelSize, viewportSize } from '../../interact/overlay-position.js';
import type { SpreadsheetInstance } from '../../mount/types.js';
import { projectDisabledState } from '../menu-a11y.js';
import type { RibbonDisplayText, ToolbarMenuText } from '../menu-text.js';
import {
  buildRibbonModel,
  HOME_MIXED_LAYOUT_GROUP_VARIANTS,
  HOME_STACKED_LAYOUT_GROUP_VARIANTS,
  HOME_TILE_LAYOUT_GROUP_VARIANTS,
  RIBBON_KEYSHORTCUTS,
  type RibbonCommand,
  type RibbonProfile,
  type RibbonTab,
  type ToolbarText,
} from '../ribbon-model.js';
import {
  RIBBON_MENU_FACTORY_FOR_COMMAND,
  RIBBON_SPLIT_BUTTON_COMMANDS,
  ribbonActivationForCommand,
} from './activation.js';
import { createRibbonButton } from './button.js';
import { createMacRibbonMenuFactory, type MacRibbonMenuFactory } from './mac/menus.js';

export type RibbonDisplayMode = 'full' | 'singleLine' | 'tabsOnly' | 'autoHide';

/** Peek/collapse flags for a display mode; a peek only exists in a collapsing mode. */
export const ribbonShellDisplayState = (
  mode: RibbonDisplayMode,
  peekRequested: boolean,
): { peek: boolean; autoHidePeek: boolean; collapsed: boolean } => {
  const collapsing = mode === 'tabsOnly' || mode === 'autoHide';
  const peek = collapsing && peekRequested;
  return { peek, autoHidePeek: mode === 'autoHide' && peek, collapsed: collapsing && !peek };
};

/** Re-projects tab selection, panel visibility and display-state flags onto a rendered ribbon shell. */
export const projectRibbonShell = (
  root: HTMLElement,
  state: { activeTab: string; displayMode: RibbonDisplayMode; peekRequested: boolean },
): void => {
  const shell = root.querySelector<HTMLElement>('.fc-tb__ribbon-shell');
  if (!shell) return;
  const { activeTab, displayMode, peekRequested } = state;
  const { peek, autoHidePeek, collapsed } = ribbonShellDisplayState(displayMode, peekRequested);
  shell.classList.toggle('fc-tb__ribbon-shell--peek', peek);
  shell.classList.toggle('fc-tb__ribbon-shell--autoHidePeek', autoHidePeek);
  shell.classList.toggle('fc-tb__ribbon-shell--collapsed', collapsed);
  if (peek) shell.dataset.ribbonPeek = 'true';
  else delete shell.dataset.ribbonPeek;
  if (autoHidePeek) shell.dataset.ribbonAutoHidePeek = 'true';
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

/** Submenu factory invoked when the user clicks a split-button. Receives the
 *  ribbon command id so a single factory can serve multiple panels (e.g.
 *  `menu-autosum-home` vs. `menu-autosum-formulas`). */
export type RibbonMenuFactory = (commandId: string) => HTMLDivElement;

/** All known submenu slots. Missing entries are silently skipped — the
 *  split-button still renders but its menu is empty until the host wires it. */
export interface RibbonMenus {
  paste?: RibbonMenuFactory;
  copy?: RibbonMenuFactory;
  pivotTable?: RibbonMenuFactory;
  definedNames?: RibbonMenuFactory;
  links?: RibbonMenuFactory;
  borders?: RibbonMenuFactory;
  underline?: RibbonMenuFactory;
  wrap?: RibbonMenuFactory;
  merge?: RibbonMenuFactory;
  textOrientation?: RibbonMenuFactory;
  conditional?: RibbonMenuFactory;
  fill?: RibbonMenuFactory;
  insertCells?: RibbonMenuFactory;
  deleteCells?: RibbonMenuFactory;
  formatCells?: RibbonMenuFactory;
  autoSum?: RibbonMenuFactory;
  freeze?: RibbonMenuFactory;
  clearArrows?: RibbonMenuFactory;
  errorChecking?: RibbonMenuFactory;
  watch?: RibbonMenuFactory;
  reviewComments?: RibbonMenuFactory;
  protect?: RibbonMenuFactory;
  calcOptions?: RibbonMenuFactory;
  sort?: RibbonMenuFactory;
  textToColumns?: RibbonMenuFactory;
  dataValidation?: RibbonMenuFactory;
  findSelect?: RibbonMenuFactory;
  pictureInsert?: RibbonMenuFactory;
  shapesInsert?: RibbonMenuFactory;
  screenshotInsert?: RibbonMenuFactory;
  chartInsert?: RibbonMenuFactory;
  tableStyle?: RibbonMenuFactory;
  cellStyles?: RibbonMenuFactory;
  currency?: RibbonMenuFactory;
  pageTheme?: RibbonMenuFactory;
  arrange?: RibbonMenuFactory;
  printArea?: RibbonMenuFactory;
  pageBreaks?: RibbonMenuFactory;
  symbol?: RibbonMenuFactory;
  script?: RibbonMenuFactory;
  addIn?: RibbonMenuFactory;
  pdf?: RibbonMenuFactory;
  clear?: RibbonMenuFactory;
}

/** Renderer helpers from select-color.ts / control-dispatch.ts. These create
 *  the inline select / color / icon DOM that ribbon buttons embed. */
export interface RibbonRenderHelpers {
  createSelect: (command: RibbonCommand) => HTMLDivElement;
  createColor: (command: RibbonCommand) => HTMLDivElement;
  createIcon: (name: string) => SVGSVGElement | null;
  makeSvg: (viewBox: string, pathData: string, className: string) => SVGSVGElement;
  chevronPath: string;
}

/** Host-owned ribbon state read on every render. */
export interface RibbonRenderState {
  getActiveTab: () => RibbonTab;
  getCollapsed: () => boolean;
  getDisplayMode: () => RibbonDisplayMode;
  getAutoHidePeek: () => boolean;
  getBackstageOpen: () => boolean;
  getDisplayMenuOpen: () => boolean;
  getFormulaBarVisible: () => boolean;
}

export interface RenderRibbonCtx {
  getInst: () => SpreadsheetInstance | null;
  ribbonLang: 'ja' | 'en';
  ribbonText: ToolbarText;
  ribbonMenuText: ToolbarMenuText;
  ribbonDisplayOptionsText: RibbonDisplayText;
  ribbonTabs?: readonly RibbonTab[];
  /** Built-in profile selected by the mount. `getProfile` wins when the host
   *  changes platform after deferred instance attachment. */
  profile?: RibbonProfile;
  getProfile?: () => RibbonProfile;
  ribbonRoot: HTMLElement | null;
  state: RibbonRenderState;
  helpers: RibbonRenderHelpers;
  menus?: RibbonMenus;
  macMenus?: MacRibbonMenuFactory;
  createBackstageView: () => HTMLElement;
  projectFormatToolbar: () => void;
}

export interface RenderRibbonApi {
  renderRibbon: () => void;
  playgroundFeatureFlags: () => FeatureFlags;
  legacyCommandIds: Record<string, string>;
  RIBBON_SPLIT_BUTTON_COMMANDS: ReadonlySet<string>;
}

/** Legacy DOM ids stamped onto ribbon buttons that pre-date the
 *  `data-ribbon-command` attribute. Existing host wirings (e.g. `wireFormat`
 *  in the playground) still look up these ids — exported so consumers don't
 *  have to mount a renderer to discover them. */
export const LEGACY_COMMAND_IDS: Record<string, string> = {
  alignC: 'btn-align-center',
  alignL: 'btn-align-left',
  alignR: 'btn-align-right',
  bold: 'btn-bold',
  borders: 'btn-borders',
  currency: 'btn-currency',
  decDown: 'btn-decimals-down',
  decUp: 'btn-decimals-up',
  fontGrow: 'btn-font-grow',
  fontShrink: 'btn-font-shrink',
  formatPainter: 'btn-format-painter',
  freeze: 'btn-freeze',
  italic: 'btn-italic',
  merge: 'btn-merge',
  middle: 'btn-middle',
  percent: 'btn-percent',
  comma: 'btn-comma',
  commentInsert: 'btn-comment',
  hyperlinkInsert: 'btn-hyperlink',
  newCommentReview: 'btn-review-comment',
  pivotTableInsert: 'btn-pivot',
  redoHome: 'btn-redo',
  strike: 'btn-strike',
  top: 'btn-top',
  underline: 'btn-underline',
  undoHome: 'btn-undo',
  wrap: 'btn-wrap',
};

/** Split-button commands that need an extra chevron, aria-haspopup, and the
 *  open/close state on the primary button. Exported so consumers can match
 *  the renderer's choice without re-listing the ids. */
export const SPLIT_BUTTON_COMMANDS = RIBBON_SPLIT_BUTTON_COMMANDS;

const TILE_LAYOUT_GROUP_VARIANTS = new Set(['tiles', ...HOME_TILE_LAYOUT_GROUP_VARIANTS]);
const STACKED_LAYOUT_GROUP_VARIANTS: ReadonlySet<string> = new Set(
  HOME_STACKED_LAYOUT_GROUP_VARIANTS,
);
const MIXED_LAYOUT_GROUP_VARIANTS: ReadonlySet<string> = new Set(HOME_MIXED_LAYOUT_GROUP_VARIANTS);

const createRibbonTabButton = (
  tab: { id: RibbonTab; label: string },
  activeRibbonTab: RibbonTab,
): HTMLButtonElement => {
  return createRibbonButton({
    className: `fc-tb__ribbon-tab${tab.id === 'file' ? ' fc-tb__ribbon-tab--file' : ''}${
      tab.id === activeRibbonTab ? ' fc-tb__ribbon-tab--active' : ''
    }`,
    role: 'tab',
    ariaSelected: tab.id === activeRibbonTab,
    tabIndex: tab.id === activeRibbonTab ? 0 : -1,
    dataset: { ribbonTab: tab.id },
    text: tab.label,
  });
};

const createRibbonCommandButton = (
  command: RibbonCommand,
  ctx: {
    ribbonText: ToolbarText;
    createIcon: RibbonRenderHelpers['createIcon'];
    makeSvg: RibbonRenderHelpers['makeSvg'];
    chevronPath: string;
  },
): HTMLButtonElement => {
  const layoutClass = command.layout === 'stacked' ? ' fc-tb__rb--stacked' : '';
  const keyshortcuts = RIBBON_KEYSHORTCUTS[command.id];
  const activation = ribbonActivationForCommand(command.id);
  const legacyId = LEGACY_COMMAND_IDS[command.id];
  const button = createRibbonButton({
    className: `fc-tb__rb${command.kind === 'large' ? ' fc-tb__rb--large' : ''}${
      command.kind === 'wide' ? ' fc-tb__rb--wide' : ''
    }${command.kind === 'mono' ? ' fc-tb__rb--mono' : ''}${layoutClass}${
      command.className ? ` ${command.className}` : ''
    }`,
    id: legacyId,
    title: command.title,
    ariaLabel: command.title,
    ariaKeyshortcuts: keyshortcuts,
    dataset: {
      ribbonCommand: command.id,
      ribbonActivation: activation.kind,
      ...(activation.menuId ? { ribbonMenuId: activation.menuId } : {}),
    },
  });
  const disabled = !!command.disabled || activation.kind === 'disabled';
  if (disabled) {
    const disabledReason = command.disabledReason ?? ctx.ribbonText.disabled;
    projectDisabledState(button, disabled, disabledReason, {
      datasetKey: 'ribbonDisabledReason',
      titlePrefix: command.title,
    });
  }
  const textOnly = !command.icon || command.kind === 'mono';
  const showLabel = textOnly || command.kind === 'wide' || command.kind === 'large';
  const icon = command.icon && command.kind !== 'mono' ? ctx.createIcon(command.icon) : null;
  if (icon) button.appendChild(icon);
  if (showLabel || (!icon && command.kind !== 'mono')) {
    const label = document.createElement('span');
    label.textContent = command.label;
    button.appendChild(label);
  }
  if (activation.menuId) {
    button.setAttribute('aria-haspopup', 'menu');
    button.setAttribute('aria-expanded', 'false');
    button.appendChild(ctx.makeSvg('0 0 12 12', ctx.chevronPath, 'fc-tb__rb-split-chevron'));
  }
  return button;
};

const createRibbonDisplayToggleButton = (
  text: RibbonDisplayText,
  menuOpen: boolean,
): HTMLButtonElement => {
  return createRibbonButton({
    className: 'fc-tb__ribbon-toggle',
    dataset: { ribbonToggle: 'true' },
    ariaHaspopup: 'menu',
    ariaExpanded: menuOpen,
    ariaLabel: text.label,
    title: text.label,
  });
};

const createRibbonDisplayOptionButton = (
  label: string,
  checked: boolean,
  option: string,
): HTMLButtonElement => {
  return createRibbonButton({
    className: 'fc-tb__ribbon-display-option',
    dataset: { ribbonDisplayOption: option },
    role: 'menuitemradio',
    ariaChecked: checked,
    text: label,
  });
};

const RIBBON_DISPLAY_MENU_GAP = 4;
const RIBBON_DISPLAY_MENU_PAD = 4;
const RIBBON_DISPLAY_MENU_FALLBACK_WIDTH = 168;
const RIBBON_DISPLAY_MENU_FALLBACK_HEIGHT = 128;

const effectiveAxisScale = (renderedSize: number, layoutSize: number): number => {
  if (renderedSize > 0 && layoutSize > 0) {
    const scale = renderedSize / layoutSize;
    if (Number.isFinite(scale) && scale > 0) return scale;
  }
  return 1;
};

/** Position the display menu after it enters the document so its actual size
 * and the toggle's viewport rect determine the side and final fixed position. */
const positionRibbonDisplayMenu = (toggle: HTMLElement, menu: HTMLElement): void => {
  const ownerDocument = toggle.ownerDocument;
  const computed = ownerDocument.defaultView?.getComputedStyle(menu);
  const preferredBelow = computed?.top !== 'auto' && computed?.bottom === 'auto';
  const viewport = viewportSize(ownerDocument);
  const pad = RIBBON_DISPLAY_MENU_PAD;
  const gap = RIBBON_DISPLAY_MENU_GAP;
  const toggleRect = toggle.getBoundingClientRect();

  // The stylesheet anchors this menu with `top`/`bottom` and `right`. Fixed
  // positioning makes the final coordinates viewport-relative; the shared
  // clamp helper converts them when a transformed ancestor owns fixed layout.
  menu.style.position = 'fixed';
  menu.style.left = '0px';
  menu.style.top = '0px';
  menu.style.right = 'auto';
  menu.style.bottom = 'auto';
  menu.style.maxHeight = '';
  menu.style.maxWidth = '';
  menu.style.overflowY = '';
  menu.style.overscrollBehavior = '';
  menu.style.width = '';
  menu.style.minWidth = '';
  menu.style.boxSizing = 'border-box';

  const naturalSize = panelSize(
    menu,
    RIBBON_DISPLAY_MENU_FALLBACK_WIDTH,
    RIBBON_DISPLAY_MENU_FALLBACK_HEIGHT,
  );
  const naturalRect = menu.getBoundingClientRect();
  const scaleX = effectiveAxisScale(naturalRect.width, menu.offsetWidth);
  const scaleY = effectiveAxisScale(naturalRect.height, menu.offsetHeight);
  const availableWidth = Math.max(0, viewport.width - pad * 2);
  if (naturalSize.width > availableWidth) {
    menu.style.minWidth = '0px';
    menu.style.width = `${availableWidth / scaleX}px`;
    menu.style.maxWidth = `${availableWidth / scaleX}px`;
  }

  const contentSize = panelSize(menu, naturalSize.width, naturalSize.height);
  const aboveSpace = Math.max(0, toggleRect.top - pad - gap);
  const belowSpace = Math.max(0, viewport.height - pad - toggleRect.bottom - gap);
  const preferredSpace = preferredBelow ? belowSpace : aboveSpace;
  const alternateSpace = preferredBelow ? aboveSpace : belowSpace;
  const preferredFits = contentSize.height <= preferredSpace;
  const alternateFits = contentSize.height <= alternateSpace;
  const opensBelow = preferredFits
    ? preferredBelow
    : alternateFits
      ? !preferredBelow
      : belowSpace >= aboveSpace;
  const availableHeight = opensBelow ? belowSpace : aboveSpace;
  if (contentSize.height > availableHeight) {
    menu.style.maxHeight = `${Math.max(0, Math.floor(availableHeight / scaleY))}px`;
    menu.style.overflowY = 'auto';
    menu.style.overscrollBehavior = 'contain';
  }

  const size = panelSize(menu, contentSize.width, contentSize.height);
  const desiredX = toggleRect.right - size.width;
  const desiredY = opensBelow ? toggleRect.bottom + gap : toggleRect.top - size.height - gap;
  const position = clampPanelToViewport(menu, desiredX, desiredY, { pad });
  menu.style.left = `${Math.round(position.x)}px`;
  menu.style.top = `${Math.round(position.y)}px`;
};

export const createRenderRibbon = (ctx: RenderRibbonCtx): RenderRibbonApi => {
  const defaultMacMenus = createMacRibbonMenuFactory(ctx.ribbonLang);
  const playgroundFeatureFlags = (): FeatureFlags => ({
    viewToolbar: false,
    watchWindow: true,
    workbookObjects: true,
    formulaBar: ctx.state.getFormulaBarVisible(),
  });

  const ribbonSubmenuFactoryFor = (
    commandId: string,
    profile: RibbonProfile,
  ): (() => HTMLDivElement | null) | null => {
    if (profile === 'excel365Mac' && commandId.startsWith('mac.')) {
      const factory = ctx.macMenus ?? defaultMacMenus;
      return () => factory(commandId);
    }
    const menus = ctx.menus;
    if (!menus) return null;
    const routeKey = RIBBON_MENU_FACTORY_FOR_COMMAND[commandId] as keyof RibbonMenus | undefined;
    if (!routeKey) return null;
    const factory = menus[routeKey];
    return factory ? () => factory(commandId) : null;
  };

  const renderRibbon = (): void => {
    const ribbonRoot = ctx.ribbonRoot;
    if (!ribbonRoot) return;
    const ribbonText = ctx.ribbonText;
    const activeRibbonTab = ctx.state.getActiveTab();
    const ribbonDisplayMode = ctx.state.getDisplayMode();
    const {
      peek: ribbonPeek,
      autoHidePeek: ribbonAutoHidePeek,
      collapsed: ribbonCollapsed,
    } = ribbonShellDisplayState(ribbonDisplayMode, ctx.state.getAutoHidePeek());
    const backstageOpen = ctx.state.getBackstageOpen();
    const ribbonDisplayMenuOpen = ctx.state.getDisplayMenuOpen();
    const ribbonDisplayOptionsText = ctx.ribbonDisplayOptionsText;
    const { createSelect, createColor, createIcon, makeSvg, chevronPath } = ctx.helpers;
    const profile = ctx.getProfile?.() ?? ctx.profile ?? 'default';
    const model = buildRibbonModel(ctx.ribbonLang, { tabs: ctx.ribbonTabs, profile });
    const shell = document.createElement('div');
    shell.className = [
      'fc-tb__ribbon-shell',
      `fc-tb__ribbon-shell--${ribbonDisplayMode}`,
      ribbonPeek && 'fc-tb__ribbon-shell--peek',
      ribbonAutoHidePeek && 'fc-tb__ribbon-shell--autoHidePeek',
      ribbonCollapsed && 'fc-tb__ribbon-shell--collapsed',
    ]
      .filter(Boolean)
      .join(' ');
    shell.dataset.ribbonDisplayMode = ribbonDisplayMode;
    if (ribbonPeek) shell.dataset.ribbonPeek = 'true';
    if (ribbonAutoHidePeek) shell.dataset.ribbonAutoHidePeek = 'true';

    const tabs = document.createElement('div');
    tabs.className = 'fc-tb__ribbon-tabs';
    tabs.setAttribute('role', 'tablist');
    tabs.setAttribute('aria-label', ribbonText.ribbonTabs);
    tabs.dataset.ribbonCollapsed = ribbonCollapsed ? 'true' : 'false';
    for (const tab of model) {
      tabs.appendChild(createRibbonTabButton(tab, activeRibbonTab));
    }
    shell.appendChild(tabs);

    for (const tab of model) {
      const panel = document.createElement('div');
      panel.className = 'fc-tb__ribbon';
      panel.setAttribute('role', 'toolbar');
      panel.setAttribute('aria-label', `${tab.label} ${ribbonText.ribbon}`);
      panel.dataset.ribbonPanel = tab.id;
      panel.hidden = tab.id !== activeRibbonTab;

      for (const g of tab.groups) {
        const group = document.createElement('section');
        const groupClasses = ['fc-tb__ribbon-group'];
        if (g.variant) {
          groupClasses.push(`fc-tb__ribbon-group--${g.variant}`);
          if (TILE_LAYOUT_GROUP_VARIANTS.has(g.variant) && g.variant !== 'tiles') {
            groupClasses.push('fc-tb__ribbon-group--tiles');
          }
          if (STACKED_LAYOUT_GROUP_VARIANTS.has(g.variant)) {
            groupClasses.push('fc-tb__ribbon-group--stacked');
          }
          if (MIXED_LAYOUT_GROUP_VARIANTS.has(g.variant)) {
            groupClasses.push('fc-tb__ribbon-group--mixed');
          }
        }
        group.className = groupClasses.join(' ');
        group.setAttribute('aria-label', g.title);

        const tools = document.createElement('div');
        tools.className = 'fc-tb__ribbon-tools';
        for (const c of g.commands) {
          if (c.kind === 'break') {
            const rowBreak = document.createElement('div');
            rowBreak.className = 'fc-tb__rb-break';
            rowBreak.dataset.ribbonCommand = c.id;
            tools.appendChild(rowBreak);
            continue;
          }
          if (c.kind === 'select') {
            tools.appendChild(createSelect(c));
            continue;
          }
          if (c.kind === 'color') {
            tools.appendChild(createColor(c));
            continue;
          }
          const b = createRibbonCommandButton(c, {
            ribbonText,
            createIcon,
            makeSvg,
            chevronPath,
          });
          tools.appendChild(b);
          const submenu = ribbonSubmenuFactoryFor(c.id, profile);
          if (submenu) {
            const menu = submenu();
            if (menu) tools.appendChild(menu);
          }
        }

        const label = document.createElement('div');
        label.className = 'fc-tb__ribbon-label';
        label.textContent = g.title;
        group.appendChild(tools);
        group.appendChild(label);
        panel.appendChild(group);
      }

      shell.appendChild(panel);
    }

    let displayToggle: HTMLButtonElement | null = null;
    let displayMenu: HTMLDivElement | null = null;
    if (!backstageOpen) {
      const display = document.createElement('div');
      display.className = 'fc-tb__ribbon-display';
      const toggle = createRibbonDisplayToggleButton(
        ribbonDisplayOptionsText,
        ribbonDisplayMenuOpen,
      );
      displayToggle = toggle;
      display.appendChild(toggle);
      if (ribbonDisplayMenuOpen) {
        const menu = document.createElement('div');
        menu.className = 'fc-tb__ribbon-display-menu';
        menu.setAttribute('role', 'menu');
        const options: [string, boolean, string][] = [
          [ribbonDisplayOptionsText.expanded, ribbonDisplayMode === 'full', 'full'],
          [ribbonDisplayOptionsText.singleLine, ribbonDisplayMode === 'singleLine', 'singleLine'],
          [ribbonDisplayOptionsText.collapsed, ribbonDisplayMode === 'tabsOnly', 'tabsOnly'],
          [ribbonDisplayOptionsText.autoHide, ribbonDisplayMode === 'autoHide', 'autoHide'],
        ];
        for (const [label, checked, option] of options) {
          const item = createRibbonDisplayOptionButton(label, checked, option);
          menu.appendChild(item);
        }
        displayMenu = menu;
        display.appendChild(menu);
      }
      shell.appendChild(display);
    }

    ribbonRoot.replaceChildren(shell);
    if (backstageOpen) ribbonRoot.appendChild(ctx.createBackstageView());
    if (displayToggle && displayMenu) positionRibbonDisplayMenu(displayToggle, displayMenu);
    ctx.projectFormatToolbar();
  };

  return {
    renderRibbon,
    playgroundFeatureFlags,
    legacyCommandIds: LEGACY_COMMAND_IDS,
    RIBBON_SPLIT_BUTTON_COMMANDS: SPLIT_BUTTON_COMMANDS,
  };
};
