// Ribbon DOM renderer extracted from main.ts. Owns the tab/panel layout, the
// per-command button rendering, the split-button chevron, the display-mode
// toggle, and the backstage hand-off. State (active tab, collapsed flag,
// backstage flag, display-menu flag) stays in the host; this factory reads
// them through getters so successive renders see the latest values.

import {
  buildRibbonModel,
  type FeatureFlags,
  RIBBON_KEYSHORTCUTS,
  type RibbonCommand,
  type RibbonDisplayText,
  type RibbonTab,
  type SpreadsheetInstance,
  type ToolbarMenuText,
  type ToolbarText,
} from '@libraz/formulon-cell';

export interface RenderRibbonCtx {
  getInst: () => SpreadsheetInstance | null;
  ribbonLang: 'ja' | 'en';
  ribbonText: ToolbarText;
  ribbonMenuText: ToolbarMenuText;
  ribbonDisplayOptionsText: RibbonDisplayText;
  ribbonRoot: HTMLElement | null;
  // Mutable state getters — the host holds the source of truth and the
  // renderer reads it fresh on every call so toggles don't have to thread
  // through here.
  getActiveRibbonTab: () => RibbonTab;
  getRibbonCollapsed: () => boolean;
  getBackstageOpen: () => boolean;
  getRibbonDisplayMenuOpen: () => boolean;
  getFormulaBarVisible: () => boolean;
  // Renderer helpers from select-color / control-dispatch.
  createRibbonSelect: (command: RibbonCommand) => HTMLDivElement;
  createRibbonColor: (command: RibbonCommand) => HTMLDivElement;
  createRibbonIcon: (name: string) => SVGSVGElement | null;
  makeSvg: (viewBox: string, pathData: string, className: string) => SVGSVGElement;
  RIBBON_CHEVRON_PATH: string;
  // Sub-menu factories (host wrappers that close over module state).
  createPasteMenu: () => HTMLDivElement;
  createPivotTableMenu: () => HTMLDivElement;
  createDefinedNamesMenu: (id: string) => HTMLDivElement;
  createLinksMenu: (id: string) => HTMLDivElement;
  createBordersMenu: () => HTMLDivElement;
  createTextOrientationMenu: () => HTMLDivElement;
  createConditionalMenu: () => HTMLDivElement;
  createFillMenu: () => HTMLDivElement;
  createInsertCellsMenu: () => HTMLDivElement;
  createDeleteCellsMenu: () => HTMLDivElement;
  createFormatCellsMenu: () => HTMLDivElement;
  createAutoSumMenu: (id: string) => HTMLDivElement;
  createFreezeMenu: () => HTMLDivElement;
  createClearArrowsMenu: () => HTMLDivElement;
  createErrorCheckingMenu: () => HTMLDivElement;
  createWatchMenu: (id: string) => HTMLDivElement;
  createReviewCommentsMenu: () => HTMLDivElement;
  createProtectMenu: (id: string) => HTMLDivElement;
  createCalcOptionsMenu: () => HTMLDivElement;
  createSortMenu: (id: string) => HTMLDivElement;
  createTextToColumnsMenu: () => HTMLDivElement;
  createDataValidationMenu: () => HTMLDivElement;
  createFindSelectMenu: () => HTMLDivElement;
  createPictureInsertMenu: () => HTMLDivElement;
  createShapesInsertMenu: () => HTMLDivElement;
  createScreenshotInsertMenu: () => HTMLDivElement;
  createChartInsertMenu: () => HTMLDivElement;
  createTableStyleMenu: (id: string) => HTMLDivElement;
  createCellStylesMenu: () => HTMLDivElement;
  createCurrencyMenu: () => HTMLDivElement;
  createPageThemeMenu: () => HTMLDivElement;
  createPrintAreaMenu: () => HTMLDivElement;
  createPageBreaksMenu: () => HTMLDivElement;
  createSheetBackgroundMenu: () => HTMLDivElement;
  createPrintTitlesMenu: () => HTMLDivElement;
  createSymbolMenu: () => HTMLDivElement;
  createScriptMenu: () => HTMLDivElement;
  createAddInMenu: () => HTMLDivElement;
  createPdfMenu: () => HTMLDivElement;
  createClearMenu: () => HTMLDivElement;
  // Backstage view factory + post-render projection hook.
  createBackstageView: () => HTMLElement;
  projectFormatToolbar: () => void;
}

export interface RenderRibbonApi {
  renderRibbon: () => void;
  playgroundFeatureFlags: () => FeatureFlags;
  legacyCommandIds: Record<string, string>;
  RIBBON_SPLIT_BUTTON_COMMANDS: Set<string>;
}

const LEGACY_COMMAND_IDS: Record<string, string> = {
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

// Split-button commands that need an extra chevron, aria-haspopup, and the
// open/close state on the primary button.
const SPLIT_BUTTON_COMMANDS = new Set<string>([
  'paste',
  'autosum',
  'autosumFormula',
  'addIn',
  'script',
  'currency',
  'pageTheme',
]);

export const createRenderRibbon = (ctx: RenderRibbonCtx): RenderRibbonApi => {
  const playgroundFeatureFlags = (): FeatureFlags => ({
    viewToolbar: false,
    watchWindow: true,
    workbookObjects: true,
    formulaBar: ctx.getFormulaBarVisible(),
  });

  // Lookup table for ribbon command-id → sub-menu DOM factory. Used by the
  // renderer to attach the right dropdown under each split-button. Kept as a
  // per-call thunk so factories that themselves close over module state (border
  // color, paste menu, etc.) stay lazy.
  const ribbonSubmenuFactoryFor = (
    commandId: string,
    groupVariant?: string,
  ): (() => HTMLDivElement) | null => {
    if (commandId === 'clearFormat') return groupVariant === 'editing' ? ctx.createClearMenu : null;
    switch (commandId) {
      case 'paste':
        return ctx.createPasteMenu;
      case 'pivotTableInsert':
        return ctx.createPivotTableMenu;
      case 'namedRangesInsert':
        return () => ctx.createDefinedNamesMenu('menu-defined-names-insert');
      case 'namedRanges':
        return () => ctx.createDefinedNamesMenu('menu-defined-names');
      case 'links':
        return () => ctx.createLinksMenu('menu-links-file');
      case 'linksInsert':
        return () => ctx.createLinksMenu('menu-links-insert');
      case 'linksData':
        return () => ctx.createLinksMenu('menu-links-data');
      case 'borders':
        return ctx.createBordersMenu;
      case 'textOrientation':
        return ctx.createTextOrientationMenu;
      case 'conditional':
        return ctx.createConditionalMenu;
      case 'fillHome':
        return ctx.createFillMenu;
      case 'insertRows':
        return ctx.createInsertCellsMenu;
      case 'deleteRows':
        return ctx.createDeleteCellsMenu;
      case 'formatCellsHome':
        return ctx.createFormatCellsMenu;
      case 'autosum':
        return () => ctx.createAutoSumMenu('menu-autosum-home');
      case 'freeze':
        return ctx.createFreezeMenu;
      case 'autosumFormula':
        return () => ctx.createAutoSumMenu('menu-autosum-formulas');
      case 'clearArrows':
        return ctx.createClearArrowsMenu;
      case 'errorChecking':
        return ctx.createErrorCheckingMenu;
      case 'watch':
        return () => ctx.createWatchMenu('menu-watch-formulas');
      case 'watchView':
        return () => ctx.createWatchMenu('menu-watch-view');
      case 'deleteCommentReview':
        return ctx.createReviewCommentsMenu;
      case 'protectReview':
        return () => ctx.createProtectMenu('menu-protect-review');
      case 'protect':
        return () => ctx.createProtectMenu('menu-protect-view');
      case 'calcOptions':
        return ctx.createCalcOptionsMenu;
      case 'filter':
        return () => ctx.createSortMenu('menu-sort');
      case 'textToColumns':
        return ctx.createTextToColumnsMenu;
      case 'dataValidation':
        return ctx.createDataValidationMenu;
      case 'sortFilterHome':
        return () => ctx.createSortMenu('menu-sort-home');
      case 'findHome':
        return ctx.createFindSelectMenu;
      case 'pictureInsert':
        return ctx.createPictureInsertMenu;
      case 'shapesInsert':
        return ctx.createShapesInsertMenu;
      case 'screenshotInsert':
        return ctx.createScreenshotInsertMenu;
      case 'chartInsert':
        return ctx.createChartInsertMenu;
      case 'formatTableHome':
        return () => ctx.createTableStyleMenu('menu-table-style-home');
      case 'formatTableInsert':
        return () => ctx.createTableStyleMenu('menu-table-style-insert');
      case 'cellStyles':
        return ctx.createCellStylesMenu;
      case 'currency':
        return ctx.createCurrencyMenu;
      case 'pageTheme':
        return ctx.createPageThemeMenu;
      case 'printArea':
        return ctx.createPrintAreaMenu;
      case 'pageBreaks':
        return ctx.createPageBreaksMenu;
      case 'sheetBackground':
        return ctx.createSheetBackgroundMenu;
      case 'printTitles':
        return ctx.createPrintTitlesMenu;
      case 'symbolInsert':
        return ctx.createSymbolMenu;
      case 'script':
        return ctx.createScriptMenu;
      case 'addIn':
        return ctx.createAddInMenu;
      case 'pdf':
        return ctx.createPdfMenu;
      default:
        return null;
    }
  };

  const renderRibbon = (): void => {
    const ribbonRoot = ctx.ribbonRoot;
    if (!ribbonRoot) return;
    const ribbonText = ctx.ribbonText;
    const activeRibbonTab = ctx.getActiveRibbonTab();
    const ribbonCollapsed = ctx.getRibbonCollapsed();
    const backstageOpen = ctx.getBackstageOpen();
    const ribbonDisplayMenuOpen = ctx.getRibbonDisplayMenuOpen();
    const ribbonDisplayOptionsText = ctx.ribbonDisplayOptionsText;
    const model = buildRibbonModel(ctx.ribbonLang);
    const shell = document.createElement('div');
    shell.className = `demo__ribbon-shell app__ribbon-shell${
      ribbonCollapsed ? ' demo__ribbon-shell--collapsed' : ''
    }`;

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
      panel.className = `demo__ribbon${tab.id === 'home' ? ' demo__ribbon--office365-home' : ''}`;
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
            tools.appendChild(ctx.createRibbonSelect(c));
            continue;
          }
          if (c.kind === 'color') {
            tools.appendChild(ctx.createRibbonColor(c));
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
          const icon = c.icon && c.kind !== 'mono' ? ctx.createRibbonIcon(c.icon) : null;
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
            b.appendChild(
              ctx.makeSvg('0 0 12 12', ctx.RIBBON_CHEVRON_PATH, 'demo__rb-split-chevron'),
            );
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
          [ribbonDisplayOptionsText.expanded, !ribbonCollapsed, 'expanded'],
          [ribbonDisplayOptionsText.collapsed, ribbonCollapsed, 'collapsed'],
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
