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
import type { RibbonDisplayText, ToolbarMenuText } from '../menu-text.js';
import {
  buildRibbonModel,
  RIBBON_KEYSHORTCUTS,
  type RibbonCommand,
  type RibbonTab,
  type ToolbarText,
} from '../ribbon-model.js';

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
  sheetBackground?: RibbonMenuFactory;
  printTitles?: RibbonMenuFactory;
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
  RIBBON_SPLIT_BUTTON_COMMANDS: Set<string>;
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
export const SPLIT_BUTTON_COMMANDS = new Set<string>([
  'paste',
  'autosum',
  'autosumFormula',
  'addIn',
  'script',
  'currency',
  'pageTheme',
  'arrangeObjectsPageLayout',
]);

// Maps a ribbon command id (with optional group variant) to the matching
// RibbonMenus key. Lives here rather than inline so the dispatch shape is
// readable and additions show up in one place.
type MenuRoute = { key: keyof RibbonMenus; variant?: string };
const MENU_ROUTES: Record<string, MenuRoute> = {
  paste: { key: 'paste' },
  pivotTableInsert: { key: 'pivotTable' },
  namedRanges: { key: 'definedNames' },
  links: { key: 'links' },
  linksData: { key: 'links' },
  borders: { key: 'borders' },
  textOrientation: { key: 'textOrientation' },
  conditional: { key: 'conditional' },
  fillHome: { key: 'fill' },
  insertRows: { key: 'insertCells' },
  deleteRows: { key: 'deleteCells' },
  formatCellsHome: { key: 'formatCells' },
  autosum: { key: 'autoSum' },
  freeze: { key: 'freeze' },
  autosumFormula: { key: 'autoSum' },
  clearArrows: { key: 'clearArrows' },
  errorChecking: { key: 'errorChecking' },
  watch: { key: 'watch' },
  watchView: { key: 'watch' },
  deleteCommentReview: { key: 'reviewComments' },
  protectReview: { key: 'protect' },
  protect: { key: 'protect' },
  calcOptions: { key: 'calcOptions' },
  filter: { key: 'sort' },
  textToColumns: { key: 'textToColumns' },
  dataValidation: { key: 'dataValidation' },
  sortFilterHome: { key: 'sort' },
  findHome: { key: 'findSelect' },
  pictureInsert: { key: 'pictureInsert' },
  shapesInsert: { key: 'shapesInsert' },
  screenshotInsert: { key: 'screenshotInsert' },
  chartInsert: { key: 'chartInsert' },
  formatTableHome: { key: 'tableStyle' },
  formatTableInsert: { key: 'tableStyle' },
  cellStyles: { key: 'cellStyles' },
  currency: { key: 'currency' },
  pageTheme: { key: 'pageTheme' },
  arrangeObjectsPageLayout: { key: 'arrange' },
  printArea: { key: 'printArea' },
  pageBreaks: { key: 'pageBreaks' },
  sheetBackground: { key: 'sheetBackground' },
  printTitles: { key: 'printTitles' },
  symbolInsert: { key: 'symbol' },
  script: { key: 'script' },
  addIn: { key: 'addIn' },
  pdf: { key: 'pdf' },
  // 'clearFormat' is special: it only renders the clear menu in the
  // editing-variant group. See ribbonSubmenuFactoryFor.
};

export const createRenderRibbon = (ctx: RenderRibbonCtx): RenderRibbonApi => {
  const playgroundFeatureFlags = (): FeatureFlags => ({
    viewToolbar: false,
    watchWindow: true,
    workbookObjects: true,
    formulaBar: ctx.state.getFormulaBarVisible(),
  });

  const ribbonSubmenuFactoryFor = (
    commandId: string,
    groupVariant?: string,
  ): (() => HTMLDivElement) | null => {
    const menus = ctx.menus;
    if (!menus) return null;
    if (commandId === 'clearFormat') {
      const f = menus.clear;
      return groupVariant === 'editing' && f ? () => f(commandId) : null;
    }
    const route = MENU_ROUTES[commandId];
    if (!route) return null;
    const factory = menus[route.key];
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
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = `demo__ribbon-tab${tab.id === 'file' ? ' demo__ribbon-tab--file' : ''}${
        tab.id === activeRibbonTab ? ' demo__ribbon-tab--active' : ''
      }`;
      btn.setAttribute('role', 'tab');
      btn.setAttribute('aria-selected', tab.id === activeRibbonTab ? 'true' : 'false');
      btn.tabIndex = tab.id === activeRibbonTab ? 0 : -1;
      btn.dataset.ribbonTab = tab.id;
      btn.textContent = tab.label;
      tabs.appendChild(btn);
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
        group.className = `demo__ribbon-group${g.variant ? ` demo__ribbon-group--${g.variant}` : ''}`;
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
          const b = document.createElement('button');
          b.type = 'button';
          b.className = `demo__rb${c.kind === 'large' ? ' demo__rb--large' : ''}${
            c.kind === 'wide' ? ' demo__rb--wide' : ''
          }${c.kind === 'mono' ? ' demo__rb--mono' : ''}`;
          b.title = c.title;
          b.setAttribute('aria-label', c.title);
          const keyshortcuts = RIBBON_KEYSHORTCUTS[c.id];
          if (keyshortcuts) b.setAttribute('aria-keyshortcuts', keyshortcuts);
          b.dataset.ribbonCommand = c.id;
          const legacyId = LEGACY_COMMAND_IDS[c.id];
          if (legacyId) b.id = legacyId;
          b.disabled = !!c.disabled;
          const textOnly = !c.icon || c.kind === 'mono';
          const showLabel = textOnly || c.kind === 'wide' || c.kind === 'large';
          const icon = c.icon && c.kind !== 'mono' ? createIcon(c.icon) : null;
          if (icon) {
            b.appendChild(icon);
          }
          if (showLabel || (!icon && c.kind !== 'mono')) {
            const label = document.createElement('span');
            label.textContent = c.label;
            b.appendChild(label);
          }
          if (SPLIT_BUTTON_COMMANDS.has(c.id)) {
            b.setAttribute('aria-haspopup', 'menu');
            b.setAttribute('aria-expanded', 'false');
            b.appendChild(makeSvg('0 0 12 12', chevronPath, 'demo__rb-split-chevron'));
          }
          tools.appendChild(b);
          const submenu = ribbonSubmenuFactoryFor(c.id, g.variant);
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
      const toggle = document.createElement('button');
      toggle.type = 'button';
      toggle.className = 'demo__ribbon-toggle';
      toggle.dataset.ribbonToggle = 'true';
      toggle.setAttribute('aria-haspopup', 'menu');
      toggle.setAttribute('aria-expanded', ribbonDisplayMenuOpen ? 'true' : 'false');
      toggle.setAttribute('aria-label', ribbonDisplayOptionsText.label);
      toggle.title = toggle.getAttribute('aria-label') ?? '';
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
          const item = document.createElement('button');
          item.type = 'button';
          item.className = 'demo__ribbon-display-option';
          item.dataset.ribbonDisplayOption = option;
          item.setAttribute('role', 'menuitemradio');
          item.setAttribute('aria-checked', checked ? 'true' : 'false');
          item.textContent = label;
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
