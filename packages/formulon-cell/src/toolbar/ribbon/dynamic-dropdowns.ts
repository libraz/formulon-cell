// Dynamic ribbon dropdown system: looks up which ribbon button owns which
// `.app__menu` panel via a single inverse map, then dispatches clicks inside
// open menus to the matching action handler. The host wires in every action
// callback through the factory; this module owns DOM open/close, focus, and
// the click dispatch table.

import type {
  CellStyleId,
  SessionChartKind,
  SpreadsheetInstance,
  TableStyle,
} from '@libraz/formulon-cell';
import type { SessionShapeKind } from '../illustration-types.js';
import { focusMenuItem } from '../menu-a11y.js';
import type { RibbonFillSeriesMode } from './fill-series.js';
import type { AutoSumFormulaName } from './menus/formulas.js';
import type { TableVariantId } from './menus/styles.js';

export type RibbonDropdownSpec = {
  menuId: string;
  command: string;
};

export type PrintTitlesAction = 'rows' | 'cols' | 'clear';
export type UiTheme = 'light' | 'dark' | 'contrast';

export interface DynamicDropdownsCtx {
  getInst: () => SpreadsheetInstance | null;
  // Menus that need refresh just before they open.
  updateCalcOptionsMenu: (menu: HTMLElement) => void;
  updateDefinedNamesMenu: (menu: HTMLElement) => void;
  // Sibling menu controllers that must close when a dynamic dropdown opens.
  closeBorderMenu: (restoreFocus?: boolean) => void;
  closeFreezeMenu: (restoreFocus?: boolean) => void;
  closePrintAreaMenu: (restoreFocus?: boolean) => void;
  closeSymbolMenu: (restoreFocus?: boolean) => void;
  // CF parent menu reference — used to avoid re-handling hover on the
  // statically-wired parent.
  getConditionalMenu: () => HTMLElement | null;
  // Action handlers used by DYNAMIC_DROPDOWN_HANDLERS, kept loosely typed so
  // each host can return either void or Promise<void>.
  applyRibbonPasteAction: (action: string) => void | Promise<void>;
  applyPivotTableAction: (action: string) => void | Promise<void>;
  applyDefinedNameAction: (action: string) => void | Promise<void>;
  applyLinksAction: (action: string) => void | Promise<void>;
  applyFillSeries: (mode?: RibbonFillSeriesMode) => void | Promise<void>;
  applyFillDirection: (direction: 'down' | 'right' | 'up' | 'left') => void;
  applyClearAction: (action: string) => void | Promise<void>;
  applyTextOrientationAction: (action: string) => void;
  applyCellInsertAction: (action: string) => void | Promise<void>;
  applyCellDeleteAction: (action: string) => void | Promise<void>;
  applyCellFormatAction: (action: string) => void | Promise<void>;
  applyPageBreakAction: (action: string) => void;
  applySheetBackgroundAction: (action: 'set' | 'clear') => void | Promise<void>;
  applyPrintTitlesAction: (action: PrintTitlesAction) => void;
  applyUiTheme: (theme: UiTheme) => void;
  focusSheet: () => void;
  applySortMenuAction: (action: string) => void;
  applyFindSelectAction: (action: string) => void;
  applyAutoSumFormula: (fn: AutoSumFormulaName) => void;
  applyFormulaAuditAction: (action: string) => void;
  applyWatchAction: (action: string) => void;
  applyReviewCommentAction: (action: string) => void;
  applyProtectAction: (action: string) => void | Promise<void>;
  applyCalcOptionAction: (action: string) => void;
  createRecommendedChartFromSelection: () => void | Promise<void>;
  createChartFromSelection: (kind: SessionChartKind) => void;
  chartKindFromAction: (action: string) => SessionChartKind;
  insertPictureFromRibbon: (action: string) => void | Promise<void>;
  insertShapeFromRibbon: (shape: SessionShapeKind) => void;
  insertScreenshotFromRibbon: () => void;
  applyScriptAction: (action: string) => void | Promise<void>;
  applyPdfAction: (action: string) => void | Promise<void>;
  createTableFromSelection: (
    style?: TableStyle,
    color?: string,
    variant?: TableVariantId,
  ) => void | Promise<void>;
  openTableStyleFooterAction: (action: string) => void | Promise<void>;
  applyCellStyleFromRibbon: (id: CellStyleId) => void;
  openCellStyleFooterAction: (action: string) => void | Promise<void>;
  applyCurrencyPreset: (symbol: string) => void;
  openCurrencyFooterAction: (action: string) => void;
  splitTextToColumns: (delimiter: string) => void | Promise<void>;
  splitTextToColumnsCustom: () => void | Promise<void>;
  applyDataValidationAction: (action: string) => void;
  applyAddInAction: (action: string) => void | Promise<void>;
  applyConditionalMenuAction: (action: string, panel?: string) => void | Promise<void>;
}

export interface DynamicDropdownsApi {
  DYNAMIC_RIBBON_DROPDOWN_IDS: ReadonlySet<string>;
  dynamicDropdownSpecForButton: (button: HTMLButtonElement) => RibbonDropdownSpec | null;
  dynamicDropdownSpecForMenu: (menu: HTMLElement) => RibbonDropdownSpec | null;
  dynamicDropdownButtonForSpec: (spec: RibbonDropdownSpec) => HTMLButtonElement | null;
  openDynamicRibbonDropdown: (spec: RibbonDropdownSpec, button?: HTMLButtonElement | null) => void;
  closeDynamicRibbonDropdown: (spec: RibbonDropdownSpec, restoreFocus?: boolean) => void;
  closeAllDynamicRibbonDropdowns: (exceptMenuId?: string) => void;
  closeDynamicConditionalSubmenus: (menu: HTMLElement) => void;
  openDynamicConditionalSubmenu: (menu: HTMLElement, key: string, trigger: HTMLElement) => void;
  dynamicRibbonDropdownClick: (event: MouseEvent) => boolean;
}

type DynamicDropdownHandler = (
  value: string,
  ctx: { menu: HTMLElement; button: HTMLButtonElement },
) => void | Promise<void>;

const DYNAMIC_RIBBON_DROPDOWN_IDS: ReadonlySet<string> = new Set([
  'menu-paste',
  'menu-pivot-table',
  'menu-defined-names-insert',
  'menu-defined-names',
  'menu-links-file',
  'menu-links-insert',
  'menu-links-data',
  'menu-conditional',
  'menu-fill',
  'menu-clear',
  'menu-text-orientation',
  'menu-insert-cells',
  'menu-delete-cells',
  'menu-format-cells',
  'menu-page-theme',
  'menu-page-breaks',
  'menu-sheet-background',
  'menu-print-titles',
  'menu-sort-home',
  'menu-sort',
  'menu-find-select',
  'menu-autosum-home',
  'menu-autosum-formulas',
  'menu-clear-arrows',
  'menu-error-checking',
  'menu-watch-formulas',
  'menu-watch-view',
  'menu-review-comments',
  'menu-protect-review',
  'menu-protect-view',
  'menu-calc-options',
  'menu-chart-insert',
  'menu-picture-insert',
  'menu-shapes-insert',
  'menu-screenshot-insert',
  'menu-script',
  'menu-table-style-home',
  'menu-table-style-insert',
  'menu-cell-styles-home',
  'menu-currency-home',
  'menu-text-to-columns',
  'menu-data-validation',
  'menu-add-ins',
  'menu-pdf',
]);

// Ribbon command id ↔ menu DOM id (the one-and-only inverse-mapping table).
// Both spec lookups read this — the click-handler walks command→menuId, the
// menu reader walks menuId→command via the derived inverse map below.
const RIBBON_DROPDOWN_MENU_FOR_COMMAND: Readonly<Record<string, string>> = {
  paste: 'menu-paste',
  pivotTableInsert: 'menu-pivot-table',
  namedRangesInsert: 'menu-defined-names-insert',
  namedRanges: 'menu-defined-names',
  links: 'menu-links-file',
  linksInsert: 'menu-links-insert',
  linksData: 'menu-links-data',
  conditional: 'menu-conditional',
  fillHome: 'menu-fill',
  textOrientation: 'menu-text-orientation',
  insertRows: 'menu-insert-cells',
  deleteRows: 'menu-delete-cells',
  formatCellsHome: 'menu-format-cells',
  pageTheme: 'menu-page-theme',
  pageBreaks: 'menu-page-breaks',
  sheetBackground: 'menu-sheet-background',
  printTitles: 'menu-print-titles',
  sortFilterHome: 'menu-sort-home',
  filter: 'menu-sort',
  findHome: 'menu-find-select',
  autosum: 'menu-autosum-home',
  autosumFormula: 'menu-autosum-formulas',
  clearArrows: 'menu-clear-arrows',
  errorChecking: 'menu-error-checking',
  watch: 'menu-watch-formulas',
  watchView: 'menu-watch-view',
  deleteCommentReview: 'menu-review-comments',
  protectReview: 'menu-protect-review',
  protect: 'menu-protect-view',
  calcOptions: 'menu-calc-options',
  chartInsert: 'menu-chart-insert',
  pictureInsert: 'menu-picture-insert',
  shapesInsert: 'menu-shapes-insert',
  screenshotInsert: 'menu-screenshot-insert',
  script: 'menu-script',
  formatTableHome: 'menu-table-style-home',
  formatTableInsert: 'menu-table-style-insert',
  cellStyles: 'menu-cell-styles-home',
  currency: 'menu-currency-home',
  textToColumns: 'menu-text-to-columns',
  dataValidation: 'menu-data-validation',
  addIn: 'menu-add-ins',
  pdf: 'menu-pdf',
};
const RIBBON_DROPDOWN_COMMAND_FOR_MENU: Readonly<Record<string, string>> = Object.fromEntries(
  Object.entries(RIBBON_DROPDOWN_MENU_FOR_COMMAND).map(([command, menuId]) => [menuId, command]),
);

export const createDynamicDropdowns = (ctx: DynamicDropdownsCtx): DynamicDropdownsApi => {
  const dynamicDropdownSpecForButton = (button: HTMLButtonElement): RibbonDropdownSpec | null => {
    const command = button.dataset.ribbonCommand ?? '';
    // `clearFormat` is dropdown only inside the Editing group — elsewhere it is
    // a plain "clear formatting" button so we explicitly exclude those.
    if (command === 'clearFormat' && button.closest<HTMLElement>('.demo__ribbon-group--editing')) {
      return { command, menuId: 'menu-clear' };
    }
    const menuId = RIBBON_DROPDOWN_MENU_FOR_COMMAND[command];
    return menuId ? { command, menuId } : null;
  };

  const dynamicDropdownSpecForMenu = (menu: HTMLElement): RibbonDropdownSpec | null => {
    if (menu.id === 'menu-clear') return { command: 'clearFormat', menuId: menu.id };
    const command = RIBBON_DROPDOWN_COMMAND_FOR_MENU[menu.id];
    return command ? { command, menuId: menu.id } : null;
  };

  const dynamicDropdownButtonForSpec = (spec: RibbonDropdownSpec): HTMLButtonElement | null => {
    if (spec.menuId === 'menu-clear') {
      return document.querySelector<HTMLButtonElement>(
        '.demo__ribbon-group--editing button[data-ribbon-command="clearFormat"]',
      );
    }
    return document.querySelector<HTMLButtonElement>(
      `button[data-ribbon-command="${spec.command}"]`,
    );
  };

  const closeDynamicConditionalSubmenus = (menu: HTMLElement): void => {
    menu.querySelectorAll<HTMLElement>('[data-cf-panel]').forEach((panel) => {
      panel.hidden = true;
    });
    menu.querySelectorAll<HTMLElement>('[data-cf-submenu]').forEach((trigger) => {
      trigger.classList.remove('app__menu-item--active');
    });
  };

  const closeDynamicRibbonDropdown = (spec: RibbonDropdownSpec, restoreFocus = false): void => {
    const menu = document.getElementById(spec.menuId) as HTMLDivElement | null;
    const button = dynamicDropdownButtonForSpec(spec);
    if (!menu) return;
    menu.hidden = true;
    if (menu.id === 'menu-conditional') closeDynamicConditionalSubmenus(menu);
    button?.setAttribute('aria-expanded', 'false');
    if (restoreFocus) button?.focus();
  };

  const closeAllDynamicRibbonDropdowns = (exceptMenuId?: string): void => {
    for (const menu of document.querySelectorAll<HTMLDivElement>('.app__menu')) {
      if (!DYNAMIC_RIBBON_DROPDOWN_IDS.has(menu.id) || menu.id === exceptMenuId) continue;
      const spec = dynamicDropdownSpecForMenu(menu);
      if (spec) closeDynamicRibbonDropdown(spec);
    }
  };

  const openDynamicRibbonDropdown = (
    spec: RibbonDropdownSpec,
    button: HTMLButtonElement | null = dynamicDropdownButtonForSpec(spec),
  ): void => {
    const menu = document.getElementById(spec.menuId) as HTMLDivElement | null;
    if (!menu || !button) return;
    if (spec.menuId === 'menu-calc-options') ctx.updateCalcOptionsMenu(menu);
    if (spec.menuId === 'menu-defined-names' || spec.menuId === 'menu-defined-names-insert') {
      ctx.updateDefinedNamesMenu(menu);
    }
    closeAllDynamicRibbonDropdowns(spec.menuId);
    ctx.closeBorderMenu();
    ctx.closeFreezeMenu();
    ctx.closePrintAreaMenu();
    ctx.closeSymbolMenu();
    menu.hidden = false;
    button.setAttribute('aria-haspopup', 'menu');
    button.setAttribute('aria-expanded', 'true');
    focusMenuItem(menu);
  };

  const openDynamicConditionalSubmenu = (
    menu: HTMLElement,
    key: string,
    trigger: HTMLElement,
  ): void => {
    closeDynamicConditionalSubmenus(menu);
    const panel = menu.querySelector<HTMLElement>(`[data-cf-panel="${key}"]`);
    if (!panel) return;
    const menuRect = menu.getBoundingClientRect();
    const triggerRect = trigger.getBoundingClientRect();
    panel.style.top = `${Math.max(0, triggerRect.top - menuRect.top - 4)}px`;
    panel.hidden = false;
    trigger.classList.add('app__menu-item--active');
  };

  // Each entry binds a `data-<attr>` button inside an open ribbon dropdown to
  // the matching action helper. The dispatcher closes the dropdown and calls
  // the handler with the attribute's value — handlers that need other dataset
  // bits (table variant, color) pull them off ctx.button.
  const DYNAMIC_DROPDOWN_HANDLERS: ReadonlyArray<{
    attr: string;
    handler: DynamicDropdownHandler;
  }> = [
    { attr: 'paste-action', handler: (v) => ctx.applyRibbonPasteAction(v) },
    { attr: 'pivot-table-action', handler: (v) => ctx.applyPivotTableAction(v) },
    { attr: 'defined-name-action', handler: (v) => ctx.applyDefinedNameAction(v) },
    { attr: 'link-action', handler: (v) => ctx.applyLinksAction(v) },
    {
      attr: 'fill',
      handler: (v) => {
        if (v === 'series') return ctx.applyFillSeries();
        if (v === 'days' || v === 'weekdays' || v === 'months' || v === 'years')
          return ctx.applyFillSeries(v);
        ctx.applyFillDirection(v as 'down' | 'right' | 'up' | 'left');
      },
    },
    { attr: 'clear', handler: (v) => ctx.applyClearAction(v) },
    { attr: 'text-orientation', handler: (v) => ctx.applyTextOrientationAction(v) },
    { attr: 'cell-insert', handler: (v) => ctx.applyCellInsertAction(v) },
    { attr: 'cell-delete', handler: (v) => ctx.applyCellDeleteAction(v) },
    { attr: 'cell-format', handler: (v) => ctx.applyCellFormatAction(v) },
    { attr: 'page-break-action', handler: (v) => ctx.applyPageBreakAction(v) },
    {
      attr: 'sheet-background-action',
      handler: (v) => ctx.applySheetBackgroundAction(v === 'clear' ? 'clear' : 'set'),
    },
    {
      attr: 'print-titles-action',
      handler: (v) => ctx.applyPrintTitlesAction(v as PrintTitlesAction),
    },
    {
      attr: 'page-theme-action',
      handler: (v) => {
        ctx.applyUiTheme(v as UiTheme);
        ctx.focusSheet();
      },
    },
    { attr: 'sort', handler: (v) => ctx.applySortMenuAction(v) },
    { attr: 'find-select', handler: (v) => ctx.applyFindSelectAction(v) },
    { attr: 'autosum-fn', handler: (v) => ctx.applyAutoSumFormula(v as AutoSumFormulaName) },
    { attr: 'formula-audit-action', handler: (v) => ctx.applyFormulaAuditAction(v) },
    { attr: 'watch-action', handler: (v) => ctx.applyWatchAction(v) },
    { attr: 'comment-action', handler: (v) => ctx.applyReviewCommentAction(v) },
    { attr: 'protect-action', handler: (v) => ctx.applyProtectAction(v) },
    { attr: 'calc-option', handler: (v) => ctx.applyCalcOptionAction(v) },
    {
      attr: 'chart-insert',
      handler: (v) => {
        if (v === 'recommended') return ctx.createRecommendedChartFromSelection();
        ctx.createChartFromSelection(ctx.chartKindFromAction(v));
      },
    },
    { attr: 'picture-insert', handler: (v) => ctx.insertPictureFromRibbon(v) },
    { attr: 'shape-insert', handler: (v) => ctx.insertShapeFromRibbon(v as SessionShapeKind) },
    { attr: 'screenshot-insert', handler: () => ctx.insertScreenshotFromRibbon() },
    { attr: 'script-action', handler: (v) => ctx.applyScriptAction(v) },
    { attr: 'pdf-action', handler: (v) => ctx.applyPdfAction(v) },
    {
      attr: 'table-style',
      handler: (v, { button }) => {
        const variant = (button.dataset.tableVariant as TableVariantId | undefined) ?? 'banded';
        return ctx.createTableFromSelection(v as TableStyle, button.dataset.tableColor, variant);
      },
    },
    { attr: 'table-style-footer', handler: (v) => ctx.openTableStyleFooterAction(v) },
    { attr: 'cell-style', handler: (v) => ctx.applyCellStyleFromRibbon(v as CellStyleId) },
    { attr: 'cell-style-footer', handler: (v) => ctx.openCellStyleFooterAction(v) },
    { attr: 'currency-preset', handler: (v) => ctx.applyCurrencyPreset(v) },
    { attr: 'currency-footer', handler: (v) => ctx.openCurrencyFooterAction(v) },
    {
      attr: 'text-to-columns-delimiter',
      handler: (v) => {
        if (v === 'custom') return ctx.splitTextToColumnsCustom();
        ctx.splitTextToColumns(v === '\\t' ? '\t' : v);
      },
    },
    { attr: 'validation-action', handler: (v) => ctx.applyDataValidationAction(v) },
    { attr: 'add-in-action', handler: (v) => ctx.applyAddInAction(v) },
  ];

  const dynamicRibbonDropdownClick = (event: MouseEvent): boolean => {
    const target = event.target as Element | null;
    const menu = target?.closest<HTMLElement>('.app__menu');
    if (!menu || !DYNAMIC_RIBBON_DROPDOWN_IDS.has(menu.id)) return false;
    const spec = dynamicDropdownSpecForMenu(menu);
    if (!spec) return false;

    // CF submenus open another pane *without* closing the parent dropdown, so
    // they live outside the table-driven loop.
    const cfSubmenu = target?.closest<HTMLElement>('[data-cf-submenu]');
    if (cfSubmenu && menu.id === 'menu-conditional') {
      event.preventDefault();
      event.stopPropagation();
      openDynamicConditionalSubmenu(menu, cfSubmenu.dataset.cfSubmenu ?? '', cfSubmenu);
      return true;
    }
    const cfItem = target?.closest<HTMLButtonElement>('[data-cf-action]');
    const cfAction = cfItem?.dataset.cfAction;
    if (cfAction && menu.id === 'menu-conditional' && !cfAction.startsWith('submenu-')) {
      event.preventDefault();
      event.stopPropagation();
      const panel = cfItem?.closest<HTMLElement>('[data-cf-panel]')?.dataset.cfPanel;
      closeDynamicRibbonDropdown(spec);
      void ctx.applyConditionalMenuAction(cfAction, panel);
      return true;
    }

    for (const entry of DYNAMIC_DROPDOWN_HANDLERS) {
      const button = target?.closest<HTMLButtonElement>(`[data-${entry.attr}]`);
      if (!button) continue;
      const datasetKey = entry.attr.replace(/-([a-z])/g, (_, c: string) => c.toUpperCase());
      const value = button.dataset[datasetKey];
      if (value === undefined) continue;
      event.preventDefault();
      event.stopPropagation();
      closeDynamicRibbonDropdown(spec);
      void entry.handler(value, { menu, button });
      return true;
    }

    return false;
  };

  return {
    DYNAMIC_RIBBON_DROPDOWN_IDS,
    dynamicDropdownSpecForButton,
    dynamicDropdownSpecForMenu,
    dynamicDropdownButtonForSpec,
    openDynamicRibbonDropdown,
    closeDynamicRibbonDropdown,
    closeAllDynamicRibbonDropdowns,
    closeDynamicConditionalSubmenus,
    openDynamicConditionalSubmenu,
    dynamicRibbonDropdownClick,
  };
};
