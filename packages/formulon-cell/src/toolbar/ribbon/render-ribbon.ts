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
  type RibbonTab,
  type ToolbarText,
} from '../ribbon-model.js';
import {
  RIBBON_MENU_FACTORY_FOR_COMMAND,
  RIBBON_SPLIT_BUTTON_COMMANDS,
  ribbonActivationForCommand,
} from './activation.js';
import { createRibbonButton } from './button.js';

export type RibbonDisplayMode = 'full' | 'singleLine' | 'tabsOnly' | 'autoHide';

/** Submenu factory invoked when the user clicks a split-button. Receives the
 *  ribbon command id so a single factory can serve multiple panels (e.g.
 *  `menu-autosum-home` vs. `menu-autosum-formulas`). */
export type RibbonMenuFactory = (commandId: string) => HTMLDivElement;

/** All known submenu slots. Missing entries are silently skipped — the
 *  split-button still renders but its menu is empty until the host wires it. */
export interface RibbonMenus {
  paste?: RibbonMenuFactory;
  pivotTable?: RibbonMenuFactory;
  definedNames?: RibbonMenuFactory;
  links?: RibbonMenuFactory;
  borders?: RibbonMenuFactory;
  underline?: RibbonMenuFactory;
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
  ribbonRoot: HTMLElement | null;
  state: RibbonRenderState;
  helpers: RibbonRenderHelpers;
  menus?: RibbonMenus;
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
    className: `demo__ribbon-tab${tab.id === 'file' ? ' demo__ribbon-tab--file' : ''}${
      tab.id === activeRibbonTab ? ' demo__ribbon-tab--active' : ''
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
  const layoutClass = command.layout === 'stacked' ? ' demo__rb--stacked' : '';
  const keyshortcuts = RIBBON_KEYSHORTCUTS[command.id];
  const activation = ribbonActivationForCommand(command.id);
  const legacyId = LEGACY_COMMAND_IDS[command.id];
  const button = createRibbonButton({
    className: `demo__rb${command.kind === 'large' ? ' demo__rb--large' : ''}${
      command.kind === 'wide' ? ' demo__rb--wide' : ''
    }${command.kind === 'mono' ? ' demo__rb--mono' : ''}${layoutClass}${
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
    const disabledReason = ctx.ribbonText.disabled;
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
    button.appendChild(ctx.makeSvg('0 0 12 12', ctx.chevronPath, 'demo__rb-split-chevron'));
  }
  return button;
};

const createRibbonDisplayToggleButton = (
  text: RibbonDisplayText,
  menuOpen: boolean,
): HTMLButtonElement => {
  return createRibbonButton({
    className: 'demo__ribbon-toggle',
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
    className: 'demo__ribbon-display-option',
    dataset: { ribbonDisplayOption: option },
    role: 'menuitemradio',
    ariaChecked: checked,
    text: label,
  });
};

export const createRenderRibbon = (ctx: RenderRibbonCtx): RenderRibbonApi => {
  const playgroundFeatureFlags = (): FeatureFlags => ({
    viewToolbar: false,
    watchWindow: true,
    workbookObjects: true,
    formulaBar: ctx.state.getFormulaBarVisible(),
  });

  const ribbonSubmenuFactoryFor = (commandId: string): (() => HTMLDivElement) | null => {
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
    const ribbonAutoHidePeek = ribbonDisplayMode === 'autoHide' && ctx.state.getAutoHidePeek();
    const ribbonCollapsed =
      ribbonDisplayMode === 'tabsOnly' || (ribbonDisplayMode === 'autoHide' && !ribbonAutoHidePeek);
    const backstageOpen = ctx.state.getBackstageOpen();
    const ribbonDisplayMenuOpen = ctx.state.getDisplayMenuOpen();
    const ribbonDisplayOptionsText = ctx.ribbonDisplayOptionsText;
    const { createSelect, createColor, createIcon, makeSvg, chevronPath } = ctx.helpers;
    const model = buildRibbonModel(ctx.ribbonLang, { tabs: ctx.ribbonTabs });
    const shell = document.createElement('div');
    shell.className = `demo__ribbon-shell app__ribbon-shell demo__ribbon-shell--${ribbonDisplayMode}${
      ribbonAutoHidePeek ? ' demo__ribbon-shell--autoHidePeek' : ''
    }${ribbonCollapsed ? ' demo__ribbon-shell--collapsed' : ''}`;
    shell.dataset.ribbonDisplayMode = ribbonDisplayMode;
    if (ribbonAutoHidePeek) shell.dataset.ribbonAutoHidePeek = 'true';

    const tabs = document.createElement('div');
    tabs.className = 'demo__ribbon-tabs';
    tabs.setAttribute('role', 'tablist');
    tabs.setAttribute('aria-label', ribbonText.ribbonTabs);
    tabs.dataset.ribbonCollapsed = ribbonCollapsed ? 'true' : 'false';
    for (const tab of model) {
      tabs.appendChild(createRibbonTabButton(tab, activeRibbonTab));
    }
    shell.appendChild(tabs);

    for (const tab of model) {
      const panel = document.createElement('div');
      panel.className = 'demo__ribbon';
      panel.setAttribute('role', 'toolbar');
      panel.setAttribute('aria-label', `${tab.label} ${ribbonText.ribbon}`);
      panel.dataset.ribbonPanel = tab.id;
      panel.hidden = tab.id !== activeRibbonTab;

      for (const g of tab.groups) {
        const group = document.createElement('section');
        const groupClasses = ['demo__ribbon-group'];
        if (g.variant) {
          groupClasses.push(`demo__ribbon-group--${g.variant}`);
          if (TILE_LAYOUT_GROUP_VARIANTS.has(g.variant) && g.variant !== 'tiles') {
            groupClasses.push('demo__ribbon-group--tiles');
          }
          if (STACKED_LAYOUT_GROUP_VARIANTS.has(g.variant)) {
            groupClasses.push('demo__ribbon-group--stacked');
          }
          if (MIXED_LAYOUT_GROUP_VARIANTS.has(g.variant)) {
            groupClasses.push('demo__ribbon-group--mixed');
          }
        }
        group.className = groupClasses.join(' ');
        group.setAttribute('aria-label', g.title);

        const tools = document.createElement('div');
        tools.className = 'demo__ribbon-tools';
        for (const c of g.commands) {
          if (c.kind === 'break') {
            const rowBreak = document.createElement('div');
            rowBreak.className = 'demo__rb-break';
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
          const submenu = ribbonSubmenuFactoryFor(c.id);
          if (submenu) tools.appendChild(submenu());
        }

        const label = document.createElement('div');
        label.className = 'demo__ribbon-label';
        label.textContent = g.title;
        group.appendChild(tools);
        group.appendChild(label);
        panel.appendChild(group);
      }

      shell.appendChild(panel);
    }

    if (!backstageOpen) {
      const display = document.createElement('div');
      display.className = 'demo__ribbon-display';
      const toggle = createRibbonDisplayToggleButton(
        ribbonDisplayOptionsText,
        ribbonDisplayMenuOpen,
      );
      display.appendChild(toggle);
      if (ribbonDisplayMenuOpen) {
        const menu = document.createElement('div');
        menu.className = 'demo__ribbon-display-menu';
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
        display.appendChild(menu);
      }
      shell.appendChild(display);
    }

    ribbonRoot.replaceChildren(shell);
    if (backstageOpen) ribbonRoot.appendChild(ctx.createBackstageView());
    ctx.projectFormatToolbar();
  };

  return {
    renderRibbon,
    playgroundFeatureFlags,
    legacyCommandIds: LEGACY_COMMAND_IDS,
    RIBBON_SPLIT_BUTTON_COMMANDS: SPLIT_BUTTON_COMMANDS,
  };
};
