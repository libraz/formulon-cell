// Dynamic ribbon dropdown system: looks up which ribbon button owns which
// `.app__menu` panel via a single inverse map, then dispatches clicks inside
// open menus to the matching action handler. The host wires in every action
// callback through the factory; this module owns DOM open/close, focus, and
// the click dispatch table.

import type { SessionChartKind, SpreadsheetInstance } from '@libraz/formulon-cell';
import { clamp, viewportSize } from '../../interact/overlay-position.js';
import type { SessionShapeKind } from '../illustration-types.js';
import { focusMenuItem } from '../menu-a11y.js';
import { RIBBON_DROPDOWN_MENU_FOR_COMMAND } from './activation.js';
import type { RibbonFillSeriesMode } from './fill-series.js';
import type { AutoSumFormulaName } from './menus/formulas.js';
import type { TableVariantId } from './menus/styles.js';

export type RibbonDropdownSpec = {
  menuId: string;
  command: string;
};

export type PrintAreaAction = 'set' | 'add' | 'clear';
export type ArrangeAction =
  | 'bring-forward'
  | 'send-backward'
  | 'bring-front'
  | 'send-back'
  | 'selection-pane';
export type UiTheme = 'light' | 'dark' | 'contrast';

export interface DynamicDropdownsCtx {
  getInst: () => SpreadsheetInstance | null;
  // Menus that need refresh just before they open.
  updateCalcOptionsMenu: (menu: HTMLElement) => void;
  updateCellDeleteMenu: (menu: HTMLElement) => void;
  updateCellInsertMenu: (menu: HTMLElement) => void;
  updateCellStylesMenu: (menu: HTMLElement) => void;
  updateClearMenu: (menu: HTMLElement) => void;
  updateClearArrowsMenu: (menu: HTMLElement) => void;
  updateCurrencyMenu: (menu: HTMLElement) => void;
  updateDataValidationMenu: (menu: HTMLElement) => void;
  updateDefinedNamesMenu: (menu: HTMLElement) => void;
  updateErrorCheckingMenu: (menu: HTMLElement) => void;
  updateFillMenu: (menu: HTMLElement) => void;
  updateFormatCellsMenu: (menu: HTMLElement) => void;
  updateFreezeMenu: (menu: HTMLElement) => void;
  updateLinksMenu: (menu: HTMLElement) => void;
  updatePasteMenu: (menu: HTMLElement) => void;
  updateArrangeMenu: (menu: HTMLElement) => void;
  updatePageBreaksMenu: (menu: HTMLElement) => void;
  updatePrintAreaMenu: (menu: HTMLElement) => void;
  updateProtectMenu: (menu: HTMLElement) => void;
  updatePageThemeMenu: (menu: HTMLElement) => void;
  updateReviewCommentsMenu: (menu: HTMLElement) => void;
  updateSortMenu: (menu: HTMLElement) => void;
  updateTableStylesMenu: (menu: HTMLElement) => void;
  updateTextOrientationMenu: (menu: HTMLElement) => void;
  updateWatchMenu: (menu: HTMLElement) => void;
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
  applyUnderlineAction: (action: string) => void | Promise<void>;
  applyMergeAction: (action: string) => void;
  applyFreezeAction: (action: string) => void;
  applyTextOrientationAction: (action: string) => void;
  applyCellInsertAction: (action: string) => void | Promise<void>;
  applyCellDeleteAction: (action: string) => void | Promise<void>;
  applyCellFormatAction: (action: string) => void | Promise<void>;
  applyPageBreakAction: (action: string) => void;
  applySheetBackgroundAction: (action: 'set' | 'clear') => void | Promise<void>;
  applyPrintAreaAction: (action: PrintAreaAction) => void;
  applyArrangeAction: (action: ArrangeAction) => void;
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
  insertScreenshotFromRibbon: (action?: string) => void | Promise<void>;
  applyScriptAction: (action: string) => void | Promise<void>;
  applyPdfAction: (action: string) => void | Promise<void>;
  createTableFromSelection: (
    style?: string,
    color?: string,
    variant?: TableVariantId,
  ) => void | Promise<void>;
  openTableStyleFooterAction: (action: string) => void | Promise<void>;
  applyPivotTableStyleFromRibbon: (styleId: string) => void | Promise<void>;
  applyCellStyleFromRibbon: (id: string) => void;
  openCellStyleFooterAction: (action: string) => void | Promise<void>;
  applyCurrencyPreset: (symbol: string) => void;
  openCurrencyFooterAction: (action: string) => void;
  splitTextToColumns: (delimiter: string) => void | Promise<void>;
  splitTextToColumnsCustom: () => void | Promise<void>;
  applyDataValidationAction: (action: string) => void;
  applyAddInAction: (action: string) => void | Promise<void>;
  applyConditionalMenuAction: (action: string, panel?: string) => void | Promise<void>;
  applySymbolAction: (symbol: string) => void | Promise<void>;
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
  dynamicRibbonDropdownPointerDown: (event: MouseEvent) => boolean;
  dynamicRibbonDropdownFocusIn: (event: FocusEvent) => boolean;
  dynamicRibbonDropdownHover: (event: MouseEvent) => boolean;
  dynamicRibbonDropdownKeydown: (event: KeyboardEvent) => boolean;
}

export const DYNAMIC_RIBBON_DROPDOWN_HANDLER_ATTRS = [
  'paste-action',
  'pivot-table-action',
  'defined-name-action',
  'link-action',
  'fill',
  'clear',
  'underline-action',
  'merge-action',
  'freeze',
  'text-orientation',
  'cell-insert',
  'cell-delete',
  'cell-format',
  'page-break-action',
  'print-area-action',
  'arrange-action',
  'page-theme-action',
  'sort',
  'find-select',
  'autosum-fn',
  'formula-audit-action',
  'watch-action',
  'comment-action',
  'protect-action',
  'calc-option',
  'chart-insert',
  'picture-insert',
  'shape-insert',
  'screenshot-insert',
  'symbol',
  'symbol-action',
  'script-action',
  'pdf-action',
  'table-style',
  'table-style-footer',
  'pivot-table-style',
  'cell-style',
  'cell-style-footer',
  'currency-preset',
  'currency-footer',
  'text-to-columns-delimiter',
  'validation-action',
  'add-in-action',
] as const;

type DynamicDropdownHandlerAttr = (typeof DYNAMIC_RIBBON_DROPDOWN_HANDLER_ATTRS)[number];

type DynamicDropdownHandler = (
  value: string,
  ctx: { menu: HTMLElement; button: HTMLButtonElement },
) => void | Promise<void>;

const datasetKeyForAttr = (attr: string): string =>
  attr.replace(/-([a-z])/g, (_, c: string) => c.toUpperCase());

const eventElement = (event: Event): Element | null =>
  event.target instanceof Element ? event.target : null;

const isDisabledMenuControl = (element: Element | null): boolean =>
  element instanceof HTMLButtonElement &&
  (element.disabled || element.getAttribute('aria-disabled') === 'true');

const RIBBON_DROPDOWN_VIEWPORT_PAD = 8;
const RIBBON_DROPDOWN_MIN_SCROLL_HEIGHT = 80;

const applyVerticalViewportLimit = (el: HTMLElement, contentHeight: number, maxHeight: number): void => {
  const height = Math.round(
    Math.max(RIBBON_DROPDOWN_MIN_SCROLL_HEIGHT, Math.min(contentHeight, maxHeight)),
  );
  el.style.maxHeight = `${height}px`;
  if (contentHeight > height) {
    el.style.overflowY = 'auto';
    el.style.overscrollBehavior = 'contain';
  } else {
    el.style.overflowY = '';
    el.style.overscrollBehavior = '';
  }
};

export const DYNAMIC_RIBBON_DROPDOWN_HANDLER_DATASET_KEYS: ReadonlySet<string> = new Set([
  ...DYNAMIC_RIBBON_DROPDOWN_HANDLER_ATTRS.map(datasetKeyForAttr),
  'cfAction',
  'cfSubmenu',
]);

export type DynamicDropdownMenuRefresherKey = {
  [K in keyof DynamicDropdownsCtx]: K extends `update${string}Menu` ? K : never;
}[keyof DynamicDropdownsCtx];

export const DYNAMIC_RIBBON_DROPDOWN_MENU_REFRESHERS: Readonly<
  Record<string, DynamicDropdownMenuRefresherKey>
> = {
  'menu-arrange-objects': 'updateArrangeMenu',
  'menu-calc-options': 'updateCalcOptionsMenu',
  'menu-cell-styles-home': 'updateCellStylesMenu',
  'menu-clear': 'updateClearMenu',
  'menu-clear-arrows': 'updateClearArrowsMenu',
  'menu-currency-home': 'updateCurrencyMenu',
  'menu-data-validation': 'updateDataValidationMenu',
  'menu-defined-names': 'updateDefinedNamesMenu',
  'menu-delete-cells': 'updateCellDeleteMenu',
  'menu-error-checking': 'updateErrorCheckingMenu',
  'menu-fill': 'updateFillMenu',
  'menu-format-cells': 'updateFormatCellsMenu',
  'menu-freeze': 'updateFreezeMenu',
  'menu-insert-cells': 'updateCellInsertMenu',
  'menu-links-data': 'updateLinksMenu',
  'menu-page-breaks': 'updatePageBreaksMenu',
  'menu-page-theme': 'updatePageThemeMenu',
  'menu-paste': 'updatePasteMenu',
  'menu-print-area': 'updatePrintAreaMenu',
  'menu-protect-review': 'updateProtectMenu',
  'menu-protect-view': 'updateProtectMenu',
  'menu-review-comments': 'updateReviewCommentsMenu',
  'menu-sort': 'updateSortMenu',
  'menu-sort-home': 'updateSortMenu',
  'menu-table-style-home': 'updateTableStylesMenu',
  'menu-text-orientation': 'updateTextOrientationMenu',
  'menu-watch-formulas': 'updateWatchMenu',
  'menu-watch-view': 'updateWatchMenu',
};

const DYNAMIC_RIBBON_DROPDOWN_IDS: ReadonlySet<string> = new Set(
  Object.values(RIBBON_DROPDOWN_MENU_FOR_COMMAND),
);

export { RIBBON_DROPDOWN_MENU_FOR_COMMAND } from './activation.js';

export const ribbonDropdownMenuIdForCommand = (commandId: string): string | null =>
  RIBBON_DROPDOWN_MENU_FOR_COMMAND[commandId] ?? null;

const RIBBON_DROPDOWN_COMMAND_FOR_MENU: Readonly<Record<string, string>> = Object.fromEntries(
  Object.entries(RIBBON_DROPDOWN_MENU_FOR_COMMAND).map(([command, menuId]) => [menuId, command]),
);

export const createDynamicDropdowns = (ctx: DynamicDropdownsCtx): DynamicDropdownsApi => {
  const menuRefreshers: Readonly<Record<string, (menu: HTMLElement) => void>> = Object.fromEntries(
    Object.entries(DYNAMIC_RIBBON_DROPDOWN_MENU_REFRESHERS).map(([menuId, key]) => [
      menuId,
      (menu: HTMLElement) => ctx[key](menu),
    ]),
  );

  const dynamicDropdownSpecForButton = (button: HTMLButtonElement): RibbonDropdownSpec | null => {
    const command = button.dataset.ribbonCommand ?? '';
    const menuId = RIBBON_DROPDOWN_MENU_FOR_COMMAND[command];
    return menuId ? { command, menuId } : null;
  };

  const dynamicDropdownSpecForMenu = (menu: HTMLElement): RibbonDropdownSpec | null => {
    const command = RIBBON_DROPDOWN_COMMAND_FOR_MENU[menu.id];
    return command ? { command, menuId: menu.id } : null;
  };

  const dynamicDropdownButtonForSpec = (spec: RibbonDropdownSpec): HTMLButtonElement | null => {
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
      trigger.setAttribute('aria-expanded', 'false');
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

  const firstOpenDynamicDropdownSpec = (): RibbonDropdownSpec | null => {
    for (const menu of document.querySelectorAll<HTMLDivElement>('.app__menu')) {
      if (menu.hidden || !DYNAMIC_RIBBON_DROPDOWN_IDS.has(menu.id)) continue;
      const spec = dynamicDropdownSpecForMenu(menu);
      if (spec) return spec;
    }
    return null;
  };

  const positionDynamicRibbonDropdown = (menu: HTMLElement, button: HTMLElement): void => {
    const buttonRect = button.getBoundingClientRect();
    const { width, height } = viewportSize();
    const pad = RIBBON_DROPDOWN_VIEWPORT_PAD;
    const menuWidth = menu.offsetWidth || 216;
    const menuHeight = menu.offsetHeight || 260;
    const left = clamp(buttonRect.left, pad, Math.max(pad, width - menuWidth - pad));
    const gap = 3;
    const belowTop = buttonRect.bottom + gap;
    const belowSpace = Math.max(0, height - pad - belowTop);
    const aboveSpace = Math.max(0, buttonRect.top - pad - gap);
    const opensBelow = belowSpace >= menuHeight || belowSpace >= aboveSpace;
    const availableHeight = opensBelow ? belowSpace : aboveSpace;
    const top = opensBelow
      ? belowTop
      : Math.max(pad, buttonRect.top - Math.min(menuHeight, availableHeight) - gap);
    menu.style.position = 'fixed';
    menu.style.left = `${Math.round(left)}px`;
    menu.style.top = `${Math.round(top)}px`;
    applyVerticalViewportLimit(menu, menuHeight, availableHeight || height - pad * 2);
  };

  const openDynamicRibbonDropdown = (
    spec: RibbonDropdownSpec,
    button: HTMLButtonElement | null = dynamicDropdownButtonForSpec(spec),
  ): void => {
    const menu = document.getElementById(spec.menuId) as HTMLDivElement | null;
    if (!menu || !button) return;
    menuRefreshers[spec.menuId]?.(menu);
    closeAllDynamicRibbonDropdowns(spec.menuId);
    ctx.closeBorderMenu();
    ctx.closeFreezeMenu();
    ctx.closePrintAreaMenu();
    ctx.closeSymbolMenu();
    menu.hidden = false;
    positionDynamicRibbonDropdown(menu, button);
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
    const panelRect = panel.getBoundingClientRect();
    const panelWidth = panelRect.width || panel.offsetWidth || 260;
    const panelHeight = panelRect.height || panel.offsetHeight || 260;
    const { width, height } = viewportSize();
    const pad = RIBBON_DROPDOWN_VIEWPORT_PAD;
    const fitsRight = menuRect.right + panelWidth <= width - pad;
    const desiredTop = Math.max(0, triggerRect.top - menuRect.top - 4);
    const maxTop = Math.max(0, height - pad - panelHeight - menuRect.top);
    const top = Math.min(desiredTop, maxTop);
    panel.style.left = fitsRight
      ? `${Math.max(0, menuRect.width - 1)}px`
      : `${-Math.max(panelWidth - 1, menuRect.width)}px`;
    panel.style.right = '';
    panel.style.top = `${Math.round(top)}px`;
    applyVerticalViewportLimit(panel, panelHeight, height - pad - menuRect.top - top);
    panel.hidden = false;
    trigger.classList.add('app__menu-item--active');
    trigger.setAttribute('aria-expanded', 'true');
  };

  // Each entry binds a `data-<attr>` button inside an open ribbon dropdown to
  // the matching action helper. The dispatcher closes the dropdown and calls
  // the handler with the attribute's value — handlers that need other dataset
  // bits (table variant, color) pull them off ctx.button.
  const DYNAMIC_DROPDOWN_HANDLERS: ReadonlyArray<{
    attr: DynamicDropdownHandlerAttr;
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
    { attr: 'underline-action', handler: (v) => ctx.applyUnderlineAction(v) },
    { attr: 'merge-action', handler: (v) => ctx.applyMergeAction(v) },
    { attr: 'freeze', handler: (v) => ctx.applyFreezeAction(v) },
    { attr: 'text-orientation', handler: (v) => ctx.applyTextOrientationAction(v) },
    { attr: 'cell-insert', handler: (v) => ctx.applyCellInsertAction(v) },
    { attr: 'cell-delete', handler: (v) => ctx.applyCellDeleteAction(v) },
    { attr: 'cell-format', handler: (v) => ctx.applyCellFormatAction(v) },
    { attr: 'page-break-action', handler: (v) => ctx.applyPageBreakAction(v) },
    {
      attr: 'print-area-action',
      handler: (v) =>
        ctx.applyPrintAreaAction(v === 'add' ? 'add' : v === 'clear' ? 'clear' : 'set'),
    },
    {
      attr: 'arrange-action',
      handler: (v) => ctx.applyArrangeAction(v as ArrangeAction),
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
    { attr: 'screenshot-insert', handler: (v) => ctx.insertScreenshotFromRibbon(v) },
    { attr: 'symbol', handler: (v) => ctx.applySymbolAction(v) },
    { attr: 'symbol-action', handler: (v) => ctx.applySymbolAction(v) },
    { attr: 'script-action', handler: (v) => ctx.applyScriptAction(v) },
    { attr: 'pdf-action', handler: (v) => ctx.applyPdfAction(v) },
    {
      attr: 'table-style',
      handler: (v, { button }) => {
        const variant = (button.dataset.tableVariant as TableVariantId | undefined) ?? 'banded';
        return ctx.createTableFromSelection(v, button.dataset.tableColor, variant);
      },
    },
    { attr: 'table-style-footer', handler: (v) => ctx.openTableStyleFooterAction(v) },
    { attr: 'pivot-table-style', handler: (v) => ctx.applyPivotTableStyleFromRibbon(v) },
    { attr: 'cell-style', handler: (v) => ctx.applyCellStyleFromRibbon(v) },
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
    const target = eventElement(event);
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
      if (isDisabledMenuControl(cfSubmenu)) return true;
      openDynamicConditionalSubmenu(menu, cfSubmenu.dataset.cfSubmenu ?? '', cfSubmenu);
      return true;
    }
    const cfItem = target?.closest<HTMLButtonElement>('[data-cf-action]');
    const cfAction = cfItem?.dataset.cfAction;
    if (cfAction && menu.id === 'menu-conditional' && !cfAction.startsWith('submenu-')) {
      event.preventDefault();
      event.stopPropagation();
      if (isDisabledMenuControl(cfItem)) return true;
      const panel = cfItem?.closest<HTMLElement>('[data-cf-panel]')?.dataset.cfPanel;
      closeDynamicRibbonDropdown(spec);
      void ctx.applyConditionalMenuAction(cfAction, panel);
      return true;
    }

    for (const entry of DYNAMIC_DROPDOWN_HANDLERS) {
      const button = target?.closest<HTMLButtonElement>(`[data-${entry.attr}]`);
      if (!button) continue;
      if (isDisabledMenuControl(button)) {
        event.preventDefault();
        event.stopPropagation();
        return true;
      }
      const datasetKey = datasetKeyForAttr(entry.attr);
      const value = button.dataset[datasetKey];
      if (value === undefined) continue;
      event.preventDefault();
      event.stopPropagation();
      // Restore focus to the menu's opener before invoking the handler so any
      // dialog the handler opens captures the opener as its `restoreFocusEl`.
      closeDynamicRibbonDropdown(spec, true);
      void entry.handler(value, { menu, button });
      return true;
    }

    return false;
  };

  const dynamicRibbonDropdownPointerDown = (event: MouseEvent): boolean => {
    const target = eventElement(event);
    if (!target) return false;
    const menu = target.closest<HTMLElement>('.app__menu');
    if (menu && DYNAMIC_RIBBON_DROPDOWN_IDS.has(menu.id)) return false;
    const button = target.closest<HTMLButtonElement>('[data-ribbon-command]');
    if (button && dynamicDropdownSpecForButton(button)) return false;
    closeAllDynamicRibbonDropdowns();
    return true;
  };

  const dynamicRibbonDropdownFocusIn = (event: FocusEvent): boolean => {
    const openSpec = firstOpenDynamicDropdownSpec();
    if (!openSpec) return false;
    const target = eventElement(event);
    if (!target) return false;
    const menu = target.closest<HTMLElement>('.app__menu');
    if (menu && DYNAMIC_RIBBON_DROPDOWN_IDS.has(menu.id)) return false;
    const button = target.closest<HTMLButtonElement>('[data-ribbon-command]');
    if (button && dynamicDropdownSpecForButton(button)) return false;
    closeDynamicRibbonDropdown(openSpec);
    closeAllDynamicRibbonDropdowns();
    return true;
  };

  const dynamicRibbonDropdownHover = (event: MouseEvent): boolean => {
    const target = eventElement(event);
    const menu = target?.closest<HTMLElement>('.app__menu');
    if (!menu || menu.id !== 'menu-conditional') return false;
    if (menu.hidden) return false;
    const trigger = target?.closest<HTMLElement>('[data-cf-submenu]');
    if (!trigger) return false;
    if (isDisabledMenuControl(trigger)) return true;
    openDynamicConditionalSubmenu(menu, trigger.dataset.cfSubmenu ?? '', trigger);
    return true;
  };

  const dynamicRibbonDropdownKeydown = (event: KeyboardEvent): boolean => {
    const target = eventElement(event);
    const menu = target?.closest<HTMLElement>('.app__menu');
    if (event.key === 'Escape') {
      if (menu && DYNAMIC_RIBBON_DROPDOWN_IDS.has(menu.id) && !menu.hidden) {
        event.preventDefault();
        event.stopPropagation();
        const spec = dynamicDropdownSpecForMenu(menu);
        if (spec) closeDynamicRibbonDropdown(spec, true);
        else closeAllDynamicRibbonDropdowns();
        return true;
      }
      const button = target?.closest<HTMLButtonElement>('[data-ribbon-command]');
      const spec = button ? dynamicDropdownSpecForButton(button) : null;
      if (spec) {
        const targetMenu = document.getElementById(spec.menuId) as HTMLElement | null;
        if (targetMenu && !targetMenu.hidden) {
          event.preventDefault();
          event.stopPropagation();
          closeDynamicRibbonDropdown(spec, true);
          return true;
        }
      }
      const openSpec = firstOpenDynamicDropdownSpec();
      if (openSpec) {
        event.preventDefault();
        event.stopPropagation();
        closeDynamicRibbonDropdown(openSpec, true);
        closeAllDynamicRibbonDropdowns();
        return true;
      }
    }
    if (!menu || menu.id !== 'menu-conditional' || menu.hidden) return false;
    const trigger = target?.closest<HTMLElement>('[data-cf-submenu]');
    if (trigger && (event.key === 'ArrowRight' || event.key === 'Enter' || event.key === ' ')) {
      event.preventDefault();
      event.stopPropagation();
      if (isDisabledMenuControl(trigger)) return true;
      const key = trigger.dataset.cfSubmenu ?? '';
      openDynamicConditionalSubmenu(menu, key, trigger);
      const panel = menu.querySelector<HTMLElement>(`[data-cf-panel="${key}"]`);
      if (panel) focusMenuItem(panel);
      return true;
    }
    const panel = target?.closest<HTMLElement>('[data-cf-panel]');
    if (panel && event.key === 'ArrowLeft') {
      event.preventDefault();
      event.stopPropagation();
      const triggerForPanel = menu.querySelector<HTMLElement>(
        `[data-cf-submenu="${panel.dataset.cfPanel ?? ''}"]`,
      );
      closeDynamicConditionalSubmenus(menu);
      triggerForPanel?.focus();
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
    dynamicRibbonDropdownPointerDown,
    dynamicRibbonDropdownFocusIn,
    dynamicRibbonDropdownHover,
    dynamicRibbonDropdownKeydown,
  };
};
