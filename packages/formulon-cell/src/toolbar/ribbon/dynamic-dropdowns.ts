// Dynamic ribbon dropdown system: looks up which ribbon button owns which
// `.fc-tb__menu` panel via a single inverse map, then dispatches clicks inside
// open menus to the matching action handler. The host wires in every action
// callback through the factory; this module owns DOM open/close, focus, and
// the click dispatch table.

import type { RibbonFillAction } from '../../commands/fill.js';
import type { SessionChartKind, SpreadsheetInstance } from '../../index.js';
import { clamp, viewportSize } from '../../interact/overlay-position.js';
import type { SessionShapeKind } from '../illustration-types.js';
import { focusMenuItem, handleMenuKeydown } from '../menu-a11y.js';
import { RIBBON_DROPDOWN_MENU_FOR_COMMAND } from './activation.js';
import type { RibbonFillSeriesMode } from './fill-series.js';
import {
  MAC_FORMULAS_MORE_MENU_ID,
  MAC_RIBBON_MENU_COMMAND_SET,
  macRibbonMenuIdForCommand,
} from './mac/model.js';
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
export type UiTheme = 'paper' | 'ink' | 'contrast';

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
  applyCopyAction: (action: string) => void | Promise<void>;
  applyRibbonPasteAction: (action: string) => void | Promise<void>;
  applyPivotTableAction: (action: string) => void | Promise<void>;
  applyDefinedNameAction: (action: string) => void | Promise<void>;
  applyLinksAction: (action: string) => void | Promise<void>;
  applyFillSeries: (mode?: RibbonFillSeriesMode) => void | Promise<void>;
  applyFillDirection: (
    direction: Extract<RibbonFillAction, 'down' | 'right' | 'up' | 'left' | 'flash'>,
  ) => void;
  applyClearAction: (action: string) => void | Promise<void>;
  applyUnderlineAction: (action: string) => void | Promise<void>;
  applyWrapAction: (action: string) => void | Promise<void>;
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
  'copy-action',
  'paste-action',
  'pivot-table-action',
  'defined-name-action',
  'link-action',
  'fill',
  'clear',
  'underline-action',
  'wrap-action',
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

const applyVerticalViewportLimit = (
  el: HTMLElement,
  contentHeight: number | null,
  maxHeight: number,
): void => {
  if (contentHeight === null) {
    el.style.maxHeight = '';
    el.style.removeProperty('overflow-y');
    el.style.removeProperty('overscroll-behavior');
    return;
  }
  const cap = Math.max(0, maxHeight);
  const height = Math.round(
    Math.min(
      cap,
      Math.max(Math.min(RIBBON_DROPDOWN_MIN_SCROLL_HEIGHT, cap), Math.min(contentHeight, cap)),
    ),
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
  'formatSubmenu',
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

const allDropdownMenus: Readonly<Record<string, string>> = {
  ...RIBBON_DROPDOWN_MENU_FOR_COMMAND,
  ...Object.fromEntries(
    [...MAC_RIBBON_MENU_COMMAND_SET].flatMap((command) => {
      const menuId = macRibbonMenuIdForCommand(command);
      return menuId ? [[command, menuId]] : [];
    }),
  ),
};

const DYNAMIC_RIBBON_DROPDOWN_IDS: ReadonlySet<string> = new Set(
  Object.values(RIBBON_DROPDOWN_MENU_FOR_COMMAND),
);
const ALL_DROPDOWN_IDS: ReadonlySet<string> = new Set(Object.values(allDropdownMenus));

export { RIBBON_DROPDOWN_MENU_FOR_COMMAND } from './activation.js';

export const ribbonDropdownMenuIdForCommand = (commandId: string): string | null =>
  allDropdownMenus[commandId] ?? null;

const RIBBON_DROPDOWN_COMMAND_FOR_MENU: Readonly<Record<string, string>> = Object.fromEntries(
  Object.entries(allDropdownMenus).map(([command, menuId]) => [menuId, command]),
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
    const menuId = allDropdownMenus[command];
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

  const closeDynamicSubmenus = (
    menu: HTMLElement,
    panelSelector: string,
    triggerSelector: string,
  ): void => {
    menu.querySelectorAll<HTMLElement>(panelSelector).forEach((panel) => {
      panel.hidden = true;
    });
    menu.querySelectorAll<HTMLElement>(triggerSelector).forEach((trigger) => {
      trigger.classList.remove('fc-tb__menu-item--active');
      trigger.setAttribute('aria-expanded', 'false');
    });
  };

  const closeDynamicConditionalSubmenus = (menu: HTMLElement): void => {
    closeDynamicSubmenus(menu, '[data-cf-panel]', '[data-cf-submenu]');
  };

  const closeDynamicFormatSubmenus = (menu: HTMLElement): void => {
    closeDynamicSubmenus(menu, '[data-format-panel]', '[data-format-submenu]');
  };

  const closeDynamicMoreSubmenus = (menu: HTMLElement): void => {
    closeDynamicSubmenus(
      menu,
      '[data-function-category-panel]',
      '[data-function-category-submenu]',
    );
  };

  const visibleMoreSubmenuForKeyEvent = (
    menu: HTMLElement,
    target: Element | null,
  ): HTMLElement | null => {
    const targetPanel = target?.closest<HTMLElement>('[data-function-category-panel]');
    if (targetPanel && !targetPanel.hidden) return targetPanel;

    const targetTrigger = target?.closest<HTMLElement>('[data-function-category-submenu]');
    if (targetTrigger?.getAttribute('aria-expanded') === 'true') {
      const key = targetTrigger.dataset.functionCategorySubmenu ?? '';
      const panel = menu.querySelector<HTMLElement>(`[data-function-category-panel="${key}"]`);
      if (panel && !panel.hidden) return panel;
    }

    const visiblePanels = Array.from(
      menu.querySelectorAll<HTMLElement>('[data-function-category-panel]'),
    ).filter((panel) => !panel.hidden);
    return visiblePanels.length === 1 ? (visiblePanels[0] ?? null) : null;
  };

  const closeDynamicRibbonDropdown = (spec: RibbonDropdownSpec, restoreFocus = false): void => {
    const menu = document.getElementById(spec.menuId) as HTMLDivElement | null;
    const button = dynamicDropdownButtonForSpec(spec);
    if (!menu) return;
    menu.hidden = true;
    if (menu.id === 'menu-conditional') closeDynamicConditionalSubmenus(menu);
    if (menu.id === 'menu-format-cells') closeDynamicFormatSubmenus(menu);
    if (menu.id === MAC_FORMULAS_MORE_MENU_ID) closeDynamicMoreSubmenus(menu);
    button?.setAttribute('aria-expanded', 'false');
    if (restoreFocus) button?.focus();
  };

  const closeAllDynamicRibbonDropdowns = (exceptMenuId?: string): void => {
    for (const menu of document.querySelectorAll<HTMLDivElement>('.fc-tb__menu')) {
      if (!ALL_DROPDOWN_IDS.has(menu.id) || menu.id === exceptMenuId) continue;
      const spec = dynamicDropdownSpecForMenu(menu);
      if (spec) closeDynamicRibbonDropdown(spec);
    }
  };

  const firstOpenDynamicDropdownSpec = (): RibbonDropdownSpec | null => {
    for (const menu of document.querySelectorAll<HTMLDivElement>('.fc-tb__menu')) {
      if (menu.hidden || !ALL_DROPDOWN_IDS.has(menu.id)) continue;
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

  const openDynamicSubmenu = (
    menu: HTMLElement,
    key: string,
    trigger: HTMLElement,
    options: {
      close: (menu: HTMLElement) => void;
      panelAttr: string;
      fallbackWidth: number;
      fallbackHeight: number;
    },
  ): void => {
    options.close(menu);
    const panel = menu.querySelector<HTMLElement>(`[data-${options.panelAttr}="${key}"]`);
    if (!panel) return;

    // A submenu keeps the last viewport limit in inline styles after it is
    // closed. Clear that state and make it measurable before reading its
    // geometry; otherwise a hidden panel reports only the fallback height and
    // its natural content can paint outside the viewport without scrolling.
    panel.style.left = '';
    panel.style.right = '';
    panel.style.top = '';
    panel.style.height = '';
    applyVerticalViewportLimit(panel, null, 0);
    panel.hidden = false;

    const menuRect = menu.getBoundingClientRect();
    const triggerRect = trigger.getBoundingClientRect();
    const panelRect = panel.getBoundingClientRect();
    const panelWidth = Math.ceil(
      panelRect.width || panel.offsetWidth || panel.scrollWidth || options.fallbackWidth,
    );
    // Once visible, getBoundingClientRect includes the panel's natural border
    // box. Hidden panels report zero, so fall back to scrollHeight/offsetHeight
    // for the first measurement instead of using a fixed height.
    const panelHeight = Math.ceil(
      panelRect.height || panel.offsetHeight || panel.scrollHeight || options.fallbackHeight,
    );
    const { width, height } = viewportSize();
    const pad = RIBBON_DROPDOWN_VIEWPORT_PAD;
    const rightX = menuRect.right - 1;
    const leftX = menuRect.left - panelWidth + 1;
    const fitsRight = rightX + panelWidth <= width - pad;
    const fitsLeft = leftX >= pad;
    let panelX: number;
    let top: number;
    let availableHeight: number;
    if (fitsRight || fitsLeft) {
      panelX = fitsRight ? rightX : leftX;
      const desiredTop = Math.max(0, triggerRect.top - menuRect.top - 4);
      const maxTop = Math.max(0, height - pad - panelHeight - menuRect.top);
      top = Math.min(desiredTop, maxTop);
      availableHeight = height - pad - menuRect.top - top;
    } else {
      panelX = clamp(menuRect.left, pad, width - panelWidth - pad);
      const belowTop = menuRect.bottom + 3;
      const belowAvailable = Math.max(0, height - pad - belowTop);
      const aboveAvailable = Math.max(0, menuRect.top - pad - 3);
      const opensBelow = belowAvailable >= aboveAvailable;
      availableHeight = opensBelow ? belowAvailable : aboveAvailable;
      const projectedHeight = Math.min(panelHeight, availableHeight);
      top = opensBelow
        ? belowTop - menuRect.top
        : menuRect.top - 3 - projectedHeight - menuRect.top;
    }
    panel.style.left = `${Math.round(panelX - menuRect.left)}px`;
    panel.style.right = '';
    panel.style.top = `${Math.round(top)}px`;
    applyVerticalViewportLimit(panel, panelHeight, availableHeight);
    trigger.classList.add('fc-tb__menu-item--active');
    trigger.setAttribute('aria-expanded', 'true');
  };

  const openDynamicConditionalSubmenu = (
    menu: HTMLElement,
    key: string,
    trigger: HTMLElement,
  ): void => {
    openDynamicSubmenu(menu, key, trigger, {
      close: closeDynamicConditionalSubmenus,
      panelAttr: 'cf-panel',
      fallbackWidth: 260,
      fallbackHeight: 260,
    });
  };

  const openDynamicFormatSubmenu = (menu: HTMLElement, key: string, trigger: HTMLElement): void => {
    openDynamicSubmenu(menu, key, trigger, {
      close: closeDynamicFormatSubmenus,
      panelAttr: 'format-panel',
      fallbackWidth: 178,
      fallbackHeight: 180,
    });
  };

  const openDynamicMoreSubmenu = (menu: HTMLElement, key: string, trigger: HTMLElement): void => {
    openDynamicSubmenu(menu, key, trigger, {
      close: closeDynamicMoreSubmenus,
      panelAttr: 'function-category-panel',
      fallbackWidth: 260,
      fallbackHeight: 300,
    });
  };

  // Each entry binds a `data-<attr>` button inside an open ribbon dropdown to
  // the matching action helper. The dispatcher closes the dropdown and calls
  // the handler with the attribute's value — handlers that need other dataset
  // bits (table variant, color) pull them off ctx.button.
  const DYNAMIC_DROPDOWN_HANDLERS: ReadonlyArray<{
    attr: DynamicDropdownHandlerAttr;
    handler: DynamicDropdownHandler;
  }> = [
    { attr: 'copy-action', handler: (v) => ctx.applyCopyAction(v) },
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
        if (v === 'down' || v === 'right' || v === 'up' || v === 'left' || v === 'flash')
          ctx.applyFillDirection(v);
      },
    },
    { attr: 'clear', handler: (v) => ctx.applyClearAction(v) },
    { attr: 'underline-action', handler: (v) => ctx.applyUnderlineAction(v) },
    { attr: 'wrap-action', handler: (v) => ctx.applyWrapAction(v) },
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
    const menu = target?.closest<HTMLElement>('.fc-tb__menu');
    if (!menu || !ALL_DROPDOWN_IDS.has(menu.id)) return false;
    const spec = dynamicDropdownSpecForMenu(menu);
    if (!spec) return false;

    const moreSubmenu = target?.closest<HTMLElement>('[data-function-category-submenu]');
    if (moreSubmenu && menu.id === MAC_FORMULAS_MORE_MENU_ID) {
      event.preventDefault();
      event.stopPropagation();
      if (isDisabledMenuControl(moreSubmenu)) return true;
      openDynamicMoreSubmenu(menu, moreSubmenu.dataset.functionCategorySubmenu ?? '', moreSubmenu);
      return true;
    }

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

    const formatSubmenu = target?.closest<HTMLElement>('[data-format-submenu]');
    if (formatSubmenu && menu.id === 'menu-format-cells') {
      event.preventDefault();
      event.stopPropagation();
      if (isDisabledMenuControl(formatSubmenu)) return true;
      openDynamicFormatSubmenu(menu, formatSubmenu.dataset.formatSubmenu ?? '', formatSubmenu);
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
    const menu = target.closest<HTMLElement>('.fc-tb__menu');
    if (menu && ALL_DROPDOWN_IDS.has(menu.id)) return false;
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
    const menu = target.closest<HTMLElement>('.fc-tb__menu');
    if (menu && ALL_DROPDOWN_IDS.has(menu.id)) {
      if (menu.id === MAC_FORMULAS_MORE_MENU_ID) {
        const trigger = target.closest<HTMLElement>('[data-function-category-submenu]');
        if (trigger) closeDynamicMoreSubmenus(menu);
      }
      return false;
    }
    const button = target.closest<HTMLButtonElement>('[data-ribbon-command]');
    if (button && dynamicDropdownSpecForButton(button)) return false;
    closeDynamicRibbonDropdown(openSpec);
    closeAllDynamicRibbonDropdowns();
    return true;
  };

  const dynamicRibbonDropdownHover = (event: MouseEvent): boolean => {
    const target = eventElement(event);
    const menu = target?.closest<HTMLElement>('.fc-tb__menu');
    if (menu?.id === MAC_FORMULAS_MORE_MENU_ID) {
      if (menu.hidden) return false;
      const trigger = target?.closest<HTMLElement>('[data-function-category-submenu]');
      if (!trigger) return false;
      if (isDisabledMenuControl(trigger)) return true;
      openDynamicMoreSubmenu(menu, trigger.dataset.functionCategorySubmenu ?? '', trigger);
      return true;
    }
    if (menu?.id !== 'menu-conditional' && menu?.id !== 'menu-format-cells') return false;
    if (menu.hidden) return false;
    if (menu.id === 'menu-format-cells') {
      const trigger = target?.closest<HTMLElement>('[data-format-submenu]');
      if (!trigger) return false;
      if (isDisabledMenuControl(trigger)) return true;
      openDynamicFormatSubmenu(menu, trigger.dataset.formatSubmenu ?? '', trigger);
      return true;
    }
    const trigger = target?.closest<HTMLElement>('[data-cf-submenu]');
    if (!trigger) return false;
    if (isDisabledMenuControl(trigger)) return true;
    openDynamicConditionalSubmenu(menu, trigger.dataset.cfSubmenu ?? '', trigger);
    return true;
  };

  const dynamicRibbonDropdownKeydown = (event: KeyboardEvent): boolean => {
    const target = eventElement(event);
    const menu = target?.closest<HTMLElement>('.fc-tb__menu');
    if (menu?.id === MAC_FORMULAS_MORE_MENU_ID && !menu.hidden) {
      const targetPanel = target?.closest<HTMLElement>('[data-function-category-panel]');
      const panelForDismiss =
        event.key === 'Escape' || event.key === 'ArrowLeft'
          ? visibleMoreSubmenuForKeyEvent(menu, target)
          : null;
      if (panelForDismiss) {
        event.preventDefault();
        event.stopPropagation();
        const trigger = menu.querySelector<HTMLElement>(
          `[data-function-category-submenu="${panelForDismiss.dataset.functionCategoryPanel ?? ''}"]`,
        );
        closeDynamicMoreSubmenus(menu);
        trigger?.focus();
        return true;
      }
      const trigger = target?.closest<HTMLElement>('[data-function-category-submenu]');
      if (trigger && (event.key === 'ArrowRight' || event.key === 'Enter' || event.key === ' ')) {
        event.preventDefault();
        event.stopPropagation();
        if (isDisabledMenuControl(trigger)) return true;
        const key = trigger.dataset.functionCategorySubmenu ?? '';
        openDynamicMoreSubmenu(menu, key, trigger);
        const targetPanel = menu.querySelector<HTMLElement>(
          `[data-function-category-panel="${key}"]`,
        );
        if (targetPanel) focusMenuItem(targetPanel);
        return true;
      }
      if (targetPanel) {
        handleMenuKeydown(event, targetPanel, {
          close: () => {
            closeDynamicMoreSubmenus(menu);
          },
        });
        return event.defaultPrevented;
      }
    }
    if (event.key === 'Escape') {
      if (menu && ALL_DROPDOWN_IDS.has(menu.id) && !menu.hidden) {
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
    if (menu?.classList.contains('fc-tb__menu--mac') && !menu.hidden) {
      const spec = dynamicDropdownSpecForMenu(menu);
      handleMenuKeydown(event, menu, {
        close: (restore) => {
          if (spec) closeDynamicRibbonDropdown(spec, restore);
        },
      });
      return event.defaultPrevented;
    }
    if ((menu?.id !== 'menu-conditional' && menu?.id !== 'menu-format-cells') || menu.hidden) {
      return false;
    }
    if (menu.id === 'menu-format-cells') {
      const trigger = target?.closest<HTMLElement>('[data-format-submenu]');
      if (trigger && (event.key === 'ArrowRight' || event.key === 'Enter' || event.key === ' ')) {
        event.preventDefault();
        event.stopPropagation();
        if (isDisabledMenuControl(trigger)) return true;
        const key = trigger.dataset.formatSubmenu ?? '';
        openDynamicFormatSubmenu(menu, key, trigger);
        const panel = menu.querySelector<HTMLElement>(`[data-format-panel="${key}"]`);
        if (panel) focusMenuItem(panel);
        return true;
      }
      const panel = target?.closest<HTMLElement>('[data-format-panel]');
      if (panel && event.key === 'ArrowLeft') {
        event.preventDefault();
        event.stopPropagation();
        const triggerForPanel = menu.querySelector<HTMLElement>(
          `[data-format-submenu="${panel.dataset.formatPanel ?? ''}"]`,
        );
        closeDynamicFormatSubmenus(menu);
        triggerForPanel?.focus();
        return true;
      }
      return false;
    }
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
