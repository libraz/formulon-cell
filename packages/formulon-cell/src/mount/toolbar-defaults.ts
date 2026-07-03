// Default factories the toolbar uses when the host doesn't supply its own.
//
// `Spreadsheet.mountToolbar(host, instance, opts)` only requires `helpers`,
// `menus`, and `hooks` because those used to be host-specific (the playground
// owned them). Now the core ships sensible defaults so React/Vue
// adapters can mount the ribbon with two lines of code; they merge their own
// app-specific overrides on top (e.g. a custom review-comment flow).
//
// What each factory provides:
//  - `createDefaultRibbonHelpers` composes the control-dispatch + select-color
//    factories that render the inline select / color / icon DOM.
//  - `createDefaultRibbonMenus` instantiates every menu category whose factory
//    lives in core. Categories that need app glue (borders color closure,
//    conditional menu open-callbacks) are stubbed with no-op closures the
//    host can override by spreading its own `menus` on top.
//  - `createDefaultRibbonHooks` wires the hook categories whose behavior is
//    fully derivable from the instance (clipboard, drawing, autosum). Hooks
//    that require app dialogs (sort/protect/review/automation) stay
//    undefined — hosts opt in by passing their own implementation.

import { listCustomCellStyles } from '../commands/cell-styles.js';
import { listCustomPivotTableStyles, listCustomTableStyles } from '../commands/format-as-table.js';
import {
  collapseColGroup,
  collapseRowGroup,
  expandColGroup,
  expandRowGroup,
  groupCols,
  groupRows,
  ungroupCols,
  ungroupRows,
} from '../commands/outline.js';
import { deleteSheetView, saveSheetView } from '../commands/sheet-views.js';
import { setSheetZoom } from '../commands/structure.js';
import { dictionaries } from '../i18n/strings.js';
import { reportDialogLabels, showReport } from '../toolbar/dialogs/report.js';
import { showZoomDialog } from '../toolbar/dialogs/zoom.js';
import { backstageMenuText, pageScaleMenuText, toolbarMenuText } from '../toolbar/menu-text.js';
import {
  analyzeAccessibilityCells,
  analyzeSpellingCells,
  buildTranslationReviewItems,
  type RibbonReportLang,
  reviewCellsFromState,
} from '../toolbar/review-tools.js';
import type { RibbonHooks } from '../toolbar/ribbon/apply-ribbon-command.js';
import { createControlDispatch } from '../toolbar/ribbon/control-dispatch.js';
import { shouldShowFontOption } from '../toolbar/ribbon/font-availability.js';
import { createBordersMenu } from '../toolbar/ribbon/menus/borders.js';
import { createConditionalMenu } from '../toolbar/ribbon/menus/conditional.js';
import { createFormulasMenuFactories } from '../toolbar/ribbon/menus/formulas.js';
import { createHomeMenuFactories } from '../toolbar/ribbon/menus/home.js';
import { createInsertMenuFactories } from '../toolbar/ribbon/menus/insert.js';
import { createPageLayoutMenuFactories } from '../toolbar/ribbon/menus/page-layout.js';
import { createPasteMenu } from '../toolbar/ribbon/menus/paste.js';
import { createReviewMenuFactories } from '../toolbar/ribbon/menus/review.js';
import { createStylesMenuFactories } from '../toolbar/ribbon/menus/styles.js';
import { createTextOrientationMenu } from '../toolbar/ribbon/menus/text-orientation.js';
import type { RibbonMenus, RibbonRenderHelpers } from '../toolbar/ribbon/render-ribbon.js';
import { createSelectColorRibbon } from '../toolbar/ribbon/select-color.js';
import { toolbarText } from '../toolbar/ribbon-model.js';
import { formatA1Range } from '../wrappers/toolbar-a1.js';
import { dispatchHostClipboard, handleAutoSum } from '../wrappers/toolbar-actions.js';
import {
  createDefaultDynamicDropdownsCtx,
  showCreateTableDialog,
} from './dynamic-dropdowns-defaults.js';
import type { SpreadsheetInstance } from './types.js';

/** Options shared by every default factory. */
export interface ToolbarDefaultsOptions {
  /** Language for built-in labels. Defaults to `instance.i18n.locale === 'en' ? 'en' : 'ja'`. */
  lang?: 'ja' | 'en';
  /** Called when a control-dispatch flow needs to push focus back to the
   *  sheet (after a font/page-setup change, e.g.). Defaults to focusing
   *  `instance.host`. */
  focusSheet?: () => void;
  /** Called after a control change that mutates cells — gives the host a
   *  chance to refresh its cached cells layer. Defaults to a no-op because
   *  the store subscription already triggers a redraw. */
  refreshCells?: () => void;
  /** Called after a control change so the toolbar can re-project active
   *  state. Wired by `mountToolbar` to its own `projectFormatToolbar`. */
  projectFormatToolbar?: () => void;
  /** Called after a zoom change so the toolbar/status indicator can refresh. */
  refreshZoom?: () => void;
  /** Used by the borders submenu to read / write the currently picked color.
   *  Defaults to a closure over `'#000000'`. Wired by `mountToolbar` to its
   *  own borderColor state. */
  getBorderColor?: () => string;
  setBorderColor?: (color: string) => void;
}

/** Shape of the readable helpers the toolbar embeds into every command
 *  button. Returned by [[createDefaultRibbonHelpers]] but also re-exported as
 *  `RibbonRenderHelpers` so callers can ship a partial helper bundle. */
export type DefaultRibbonHelpers = RibbonRenderHelpers;

const resolveLang = (instance: SpreadsheetInstance, opts: ToolbarDefaultsOptions): 'ja' | 'en' =>
  opts.lang ?? (instance.i18n.locale === 'en' ? 'en' : 'ja');

export function createDefaultRibbonHelpers(
  instance: SpreadsheetInstance,
  opts: ToolbarDefaultsOptions = {},
): DefaultRibbonHelpers {
  const lang = resolveLang(instance, opts);
  const ribbonText = toolbarText(lang);
  const pageScaleText = pageScaleMenuText(lang);
  const focusSheet = opts.focusSheet ?? ((): void => instance.host.focus());
  const refreshCells = opts.refreshCells ?? ((): void => undefined);
  const projectFormatToolbar = opts.projectFormatToolbar ?? ((): void => undefined);
  const getInst = (): SpreadsheetInstance => instance;

  const dispatch = createControlDispatch({
    getInst,
    ribbonLang: lang,
    ribbonText,
    pageScaleText,
    sheetEl: instance.host,
    focusSheet,
    refreshWorkbookCells: refreshCells,
    projectFormatToolbar,
  });
  const sc = createSelectColorRibbon({
    ribbonLang: lang,
    ribbonText,
    pageScaleText,
    getInst,
    applyRibbonControl: dispatch.applyRibbonControl,
    currentRibbonControlValue: dispatch.currentRibbonControlValue,
    shouldShowFontOption,
    createRibbonIcon: dispatch.createRibbonIcon,
  });

  return {
    createSelect: sc.createRibbonSelect,
    createColor: sc.createRibbonColor,
    createIcon: dispatch.createRibbonIcon,
    makeSvg: sc.makeSvg,
    chevronPath: sc.RIBBON_CHEVRON_PATH,
  };
}

/** Mirrors `ToolbarDefaultsOptions` because menu construction needs the same
 *  border-color closures as the helpers — keeping them in one bag avoids two
 *  separate places for the host to wire identical glue. */
export function createDefaultRibbonMenus(
  instance: SpreadsheetInstance,
  opts: ToolbarDefaultsOptions = {},
): RibbonMenus {
  const lang = resolveLang(instance, opts);
  const ribbonText = toolbarText(lang);
  const ribbonMenuText = toolbarMenuText(lang);
  const getBorderColor = opts.getBorderColor ?? ((): string => '#000000');
  const setBorderColor = opts.setBorderColor ?? ((): void => undefined);

  const insertFactories = createInsertMenuFactories(ribbonMenuText);
  const pageLayoutFactories = createPageLayoutMenuFactories(ribbonMenuText);
  const formulaFactories = createFormulasMenuFactories(ribbonMenuText, lang);
  const reviewFactories = createReviewMenuFactories(ribbonMenuText);
  const homeFactories = createHomeMenuFactories({
    ribbonLang: lang,
    ribbonMenuText,
    ribbonText,
    formatDialog: dictionaries[lang].formatDialog,
    sheetTabs: dictionaries[lang].sheetTabs,
    viewToolbar: dictionaries[lang].viewToolbar,
  });
  const styleFactories = createStylesMenuFactories({
    ribbonLang: lang,
    ribbonMenuText,
    ribbonText,
    customCellStyles: () => listCustomCellStyles(instance.store.getState()),
    customTableStyles: () => listCustomTableStyles(instance.store.getState()),
    customPivotTableStyles: () => listCustomPivotTableStyles(instance.store.getState()),
  });
  const buildBorders = (): HTMLDivElement =>
    createBordersMenu({
      ribbonText,
      getBorderColor,
      onPickColor: setBorderColor,
    });
  const buildConditional = (): HTMLDivElement => createConditionalMenu(lang);
  const buildPaste = (): HTMLDivElement => createPasteMenu(dictionaries[lang]);
  const buildTextOrientation = (): HTMLDivElement => createTextOrientationMenu(ribbonMenuText);

  return {
    paste: buildPaste,
    copy: homeFactories.createCopyMenu,
    borders: buildBorders,
    underline: homeFactories.createUnderlineMenu,
    wrap: homeFactories.createWrapMenu,
    merge: homeFactories.createMergeMenu,
    textOrientation: buildTextOrientation,
    conditional: buildConditional,

    // Insert tab
    pivotTable: insertFactories.createPivotTableMenu,
    definedNames: insertFactories.createDefinedNamesMenu,
    links: insertFactories.createLinksMenu,
    pictureInsert: insertFactories.createPictureInsertMenu,
    shapesInsert: insertFactories.createShapesInsertMenu,
    screenshotInsert: insertFactories.createScreenshotInsertMenu,
    chartInsert: insertFactories.createChartInsertMenu,
    symbol: insertFactories.createSymbolMenu,
    script: insertFactories.createScriptMenu,
    addIn: insertFactories.createAddInMenu,
    pdf: insertFactories.createPdfMenu,
    dataValidation: insertFactories.createDataValidationMenu,

    // Page Layout tab
    pageTheme: pageLayoutFactories.createPageThemeMenu,
    arrange: pageLayoutFactories.createArrangeMenu,
    printArea: pageLayoutFactories.createPrintAreaMenu,
    pageBreaks: pageLayoutFactories.createPageBreaksMenu,

    // Formulas tab
    autoSum: formulaFactories.createAutoSumMenu,
    calcOptions: formulaFactories.createCalcOptionsMenu,
    clearArrows: formulaFactories.createClearArrowsMenu,
    errorChecking: formulaFactories.createErrorCheckingMenu,

    // Review tab
    watch: reviewFactories.createWatchMenu,
    reviewComments: reviewFactories.createReviewCommentsMenu,
    protect: reviewFactories.createProtectMenu,

    // Home tab
    fill: homeFactories.createFillMenu,
    clear: homeFactories.createClearMenu,
    freeze: homeFactories.createFreezeMenu,
    insertCells: homeFactories.createInsertCellsMenu,
    deleteCells: homeFactories.createDeleteCellsMenu,
    formatCells: homeFactories.createFormatCellsMenu,
    sort: homeFactories.createSortMenu,
    textToColumns: homeFactories.createTextToColumnsMenu,
    findSelect: homeFactories.createFindSelectMenu,

    // Styles
    tableStyle: styleFactories.createTableStyleMenu,
    cellStyles: styleFactories.createCellStylesMenu,
    currency: styleFactories.createCurrencyMenu,
  };
}

/** Hooks the toolbar can satisfy from instance methods alone — clipboard,
 *  drawing, autoSum. Hosts merge their own categories (review, protection,
 *  automation, …) on top because those involve app-specific dialogs. */
export function createDefaultRibbonHooks(
  instance: SpreadsheetInstance,
  opts: ToolbarDefaultsOptions = {},
): RibbonHooks {
  const dropdowns = createDefaultDynamicDropdownsCtx(
    instance as unknown as Parameters<typeof createDefaultDynamicDropdownsCtx>[0],
    {
      focusSheet: opts.focusSheet,
      projectFormatToolbar: opts.projectFormatToolbar,
      refreshCells: opts.refreshCells,
    },
  );
  const reportLang: RibbonReportLang = instance.i18n.locale === 'ja' ? 'ja' : 'en';
  const showSharedReport = async (
    title: string,
    items: { severity: 'info' | 'warning'; label: string; detail: string }[],
  ): Promise<void> => {
    const strings = instance.i18n.strings;
    await showReport({
      title,
      items,
      ...reportDialogLabels(strings),
    });
    instance.host.focus();
  };
  const showReviewReport = (
    title: string,
    buildItems: (
      cells: ReturnType<typeof reviewCellsFromState>,
    ) => { severity: 'info' | 'warning'; label: string; detail: string }[],
  ): void => {
    const state = instance.store.getState();
    const cells = reviewCellsFromState(state, state.data.sheetIndex, state.selection.range);
    void showSharedReport(title, buildItems(cells));
  };
  const applyOutline = (action: 'group' | 'ungroup' | 'show-detail' | 'hide-detail'): void => {
    const range = instance.store.getState().selection.range;
    const useRows = range.r1 > range.r0 || range.c0 === range.c1;
    if (action === 'group') {
      if (useRows)
        groupRows(instance.store, instance.history, range.r0, range.r1, instance.workbook);
      else groupCols(instance.store, instance.history, range.c0, range.c1, instance.workbook);
    } else if (action === 'ungroup') {
      if (useRows)
        ungroupRows(instance.store, instance.history, range.r0, range.r1, instance.workbook);
      else ungroupCols(instance.store, instance.history, range.c0, range.c1, instance.workbook);
    } else if (action === 'hide-detail') {
      if (useRows)
        collapseRowGroup(instance.store, instance.history, range.r0, range.r1, instance.workbook);
      else
        collapseColGroup(instance.store, instance.history, range.c0, range.c1, instance.workbook);
    } else if (useRows) {
      expandRowGroup(instance.store, instance.history, range.r0, range.r1, instance.workbook);
    } else {
      expandColGroup(instance.store, instance.history, range.c0, range.c1, instance.workbook);
    }
    instance.host.focus();
  };
  return {
    clipboard: {
      copy: () => {
        dispatchHostClipboard(instance, 'copy');
      },
      cut: () => {
        dispatchHostClipboard(instance, 'cut');
      },
      paste: () => {
        dispatchHostClipboard(instance, 'paste');
      },
    },
    formula: {
      autoSum: () => {
        handleAutoSum(instance, 'SUM');
      },
      errorChecking: () => {
        void dropdowns.applyFormulaAuditAction('error-checking');
      },
    },
    sortFilter: {
      sort: (direction) => void dropdowns.applySortMenuAction(direction),
      customSort: () => dropdowns.applySortMenuAction('custom'),
      openFilter: () => {
        void dropdowns.applySortMenuAction('filter');
      },
      removeDuplicates: () => {
        void dropdowns.applySortMenuAction('dedupe');
      },
      splitTextToColumns: (sep) => void dropdowns.splitTextToColumns(sep),
      splitTextToColumnsCustom: () => void dropdowns.splitTextToColumnsCustom(),
    },
    insert: {
      createTable: () => dropdowns.createTableFromSelection('medium'),
      createTableDialog: () =>
        showCreateTableDialog(instance as unknown as Parameters<typeof showCreateTableDialog>[0]),
      createChart: () => {
        dropdowns.createChartFromSelection('column');
      },
      createRecommendedChart: () => dropdowns.createRecommendedChartFromSelection(),
      insertPicture: (source) => dropdowns.insertPictureFromRibbon(source),
      insertShape: (shape) => {
        dropdowns.insertShapeFromRibbon(shape);
      },
      insertScreenshot: () => {
        dropdowns.insertScreenshotFromRibbon();
      },
      insertSymbol: (symbol) => dropdowns.applySymbolAction(symbol),
    },
    page: {
      pageBreak: () => {
        dropdowns.applyPageBreakAction('insert');
      },
      sheetBackground: (action) => {
        const sheet = instance.store.getState().data.sheetIndex;
        const nextAction =
          action === 'set' && instance.store.getState().ui.sheetBackgroundImages.has(sheet)
            ? 'clear'
            : action;
        return dropdowns.applySheetBackgroundAction(nextAction);
      },
      pdf: (action) => {
        void dropdowns.applyPdfAction(action);
      },
      inspect: () => {
        instance.openWorkbookObjects();
      },
      outline: applyOutline,
    },
    review: {
      spelling: () => {
        showReviewReport(instance.i18n.strings.ribbon.spelling, (cells) =>
          analyzeSpellingCells(cells, reportLang),
        );
      },
      translate: () => {
        showReviewReport(instance.i18n.strings.ribbon.translate, (cells) =>
          buildTranslationReviewItems(cells, reportLang),
        );
      },
      accessibility: () => {
        showReviewReport(instance.i18n.strings.ribbon.accessibility, (cells) =>
          analyzeAccessibilityCells(cells, reportLang),
        );
      },
      deleteComment: () => {
        dropdowns.applyReviewCommentAction('delete-active');
      },
      selectComment: () => undefined,
    },
    protection: {
      runSheet: () =>
        dropdowns.applyProtectAction(
          instance.isSheetProtected() ? 'unprotect-sheet' : 'protect-sheet',
        ),
      runWorkbook: (protect) =>
        dropdowns.applyProtectAction(protect ? 'protect-workbook' : 'unprotect-workbook'),
      allowEditRanges: () => dropdowns.applyProtectAction('allow-edit-ranges'),
    },
    automation: {
      runScript: () => dropdowns.applyScriptAction('custom'),
      recordActions: () => {
        const strings = instance.i18n.strings.ribbonMenu;
        const range = instance.store.getState().selection.range;
        void showSharedReport(strings.recordActionsStatus, [
          {
            severity: 'info',
            label: strings.recordActionsStatus,
            detail: `${strings.recordActionsEmpty} (${formatA1Range(range)})`,
          },
        ]);
      },
      allScripts: () => {
        const strings = instance.i18n.strings.ribbonMenu;
        void showSharedReport(strings.automationScriptsTitle, [
          {
            severity: 'info',
            label: strings.automationBuiltInScriptsLabel,
            detail: strings.automationBuiltInScriptsDetail,
          },
          {
            severity: 'info',
            label: strings.automationRecentRunsLabel,
            detail: strings.automationNoRuns,
          },
        ]);
      },
      addInManager: () => {
        void dropdowns.applyAddInAction('manage');
      },
    },
    sheetView: {
      save: () => {
        const state = instance.store.getState();
        const count =
          state.sheetViews.views.filter((view) => view.sheet === state.data.sheetIndex).length + 1;
        saveSheetView(
          instance.store,
          `view-${state.data.sheetIndex}-${count}`,
          `${instance.i18n.strings.viewToolbar.views} ${count}`,
          undefined,
          instance.history,
        );
        instance.host.focus();
      },
      deleteActive: () => {
        const id = instance.store.getState().sheetViews.activeViewId;
        if (id) deleteSheetView(instance.store, id, instance.history);
        instance.host.focus();
      },
      zoomDialog: async () => {
        const strings = instance.i18n.strings.ribbon;
        const current = Math.round(instance.store.getState().viewport.zoom * 100);
        const next = await showZoomDialog({
          title: strings.zoomDialogTitle,
          label: strings.zoomDialogPercent,
          initial: current,
          okLabel: instance.i18n.strings.formatDialog.ok,
          cancelLabel: instance.i18n.strings.formatDialog.cancel,
          invalidMessage: strings.zoomDialogInvalid,
        });
        if (next === null) return;
        setSheetZoom(instance.store, next / 100, instance.workbook);
        opts.refreshZoom?.();
        instance.host.focus();
      },
    },
  };
}

/** Default backstage view — an empty div. Hosts that want a full file menu
 *  pass their own `createBackstageView` to `mountToolbar`. */
export function createDefaultBackstageView(_instance: SpreadsheetInstance): HTMLElement {
  const placeholder = document.createElement('div');
  placeholder.className = 'demo__backstage demo__backstage--placeholder';
  return placeholder;
}

/** Re-exports so consumers can grab text dictionaries without reaching into
 *  the toolbar subfolder by hand. */
export { backstageMenuText, pageScaleMenuText, toolbarMenuText, toolbarText };
