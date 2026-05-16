// Ribbon command dispatcher. Extracted from main.ts. The function fans out a
// ribbon command id to the matching handler — first via the four declarative
// dispatch tables (dialog openers, function-arg openers, format mutators, zoom
// presets, view modes, border-draw modes), then via the big switch for ids
// that need bespoke wiring.
//
// The dispatcher reaches into ~50 module-scoped helpers in main.ts. Rather
// than build a factory that destructures all of them once, we pass them in
// per-call via this deps struct: it keeps the call site tagged with the exact
// surface area and avoids a stale closure if any of those helpers get
// recreated. Most fields are 0-arg thunks that already close over `inst`.

import {
  applyMerge,
  applyUnmerge,
  type CellBorderStyle,
  deleteCols,
  deleteRows,
  type FeatureFlags,
  getPageSetup,
  hideCols,
  hideRows,
  insertCols,
  insertRows,
  isWorkbookStructureProtected,
  type SpreadsheetInstance,
  setFont,
  setGridlinesVisible,
  setHeadingsVisible,
  setPrintGridlines,
  setPrintHeadings,
  setR1C1ReferenceStyle,
  setSheetZoom,
  setShowFormulas,
  setWorkbookView,
  type ToolbarMenuText,
  type ToolbarText,
} from '@libraz/formulon-cell';
import {
  RIBBON_BORDER_DRAW_MODES,
  RIBBON_DIALOG_OPENERS,
  RIBBON_FORMAT_MUTATORS,
  RIBBON_FUNCTION_ARG_OPENERS,
  RIBBON_VIEW_MODES,
  RIBBON_ZOOM_PRESETS,
  type RibbonFormatMutator,
} from './command-tables.js';
import type { AutoSumFormulaName } from './menus/formulas.js';

type UiTheme = 'light' | 'dark' | 'contrast';

/** Dependencies threaded into `applyRibbonCommand` on each call. */
export interface ApplyRibbonCommandDeps {
  // Core handles + i18n -------------------------------------------------------
  inst: SpreadsheetInstance | null;
  ribbonText: ToolbarText;
  ribbonMenuText: ToolbarMenuText;
  uiTheme: UiTheme;
  selectedBorderStyle: CellBorderStyle;
  selectedBorderColor: string;
  formulaBarVisible: boolean;

  // Module-scope shared helpers ----------------------------------------------
  applyRibbonFormat: (fn: RibbonFormatMutator) => void;
  applyUiTheme: (theme: UiTheme) => void;
  focusSheet: () => void;
  projectFormatToolbar: () => void;
  refreshWorkbookCells: () => void;
  refreshZoom: () => void;
  selectedRowCount: () => number;
  selectedColCount: () => number;
  setFormulaBarVisible: (next: boolean) => void;
  playgroundFeatureFlags: () => FeatureFlags;
  showMessage: (opts: { title: string; message: string }) => Promise<void> | void;

  // Clipboard ----------------------------------------------------------------
  copySelectionToClipboard: () => Promise<unknown> | unknown;
  cutSelectionToClipboard: () => Promise<unknown> | unknown;
  pasteClipboardIntoSelection: () => Promise<unknown> | unknown;

  // Sort / filter / data -----------------------------------------------------
  sortSelection: (dir: 'asc' | 'desc') => void;
  customSortSelection: () => Promise<unknown> | unknown;
  openFilterForSelection: () => void;
  removeDuplicateRows: () => void;
  splitTextToColumns: (sep: string) => void;

  // Insert -------------------------------------------------------------------
  createTableFromSelection: (variant: 'medium') => Promise<unknown> | unknown;
  createChartFromSelection: () => void;
  insertPictureFromRibbon: (source: 'online') => Promise<unknown> | unknown;
  insertShapeFromRibbon: (shape: 'rectangle') => void;
  insertScreenshotFromRibbon: () => void;

  // Page / outline / backstage ------------------------------------------------
  applyPageBreakAction: () => void;
  applySheetBackgroundAction: (action: 'set') => Promise<unknown> | unknown;
  applyPdfAction: (action: 'create') => void;
  inspectWorkbookFromBackstage: () => void;
  applyOutlineAction: (action: 'group' | 'ungroup' | 'show-detail' | 'hide-detail') => void;

  // Drawing / ink ------------------------------------------------------------
  setDrawInkMode: (mode: 'pen' | 'erase') => void;

  // Review / proofing --------------------------------------------------------
  runSpellingReview: () => void;
  openTranslateReview: () => void;
  runAccessibilityCheck: () => void;
  deleteActiveReviewComment: () => void;
  selectReviewComment: (direction: 1 | -1) => void;

  // Protection ---------------------------------------------------------------
  runSheetProtectionFlow: () => Promise<unknown> | unknown;
  runWorkbookProtectionFlow: (protect: boolean) => Promise<unknown> | unknown;
  applyProtectAction: (action: 'allow-edit-ranges') => Promise<unknown> | unknown;

  // Automation / scripts -----------------------------------------------------
  runPlaygroundScript: () => Promise<unknown> | unknown;
  recordSelectedActions: () => void;
  openAllScripts: () => void;
  openAddInManager: () => void;

  // Formula auditing ---------------------------------------------------------
  applyAutoSumFormula: (fn: AutoSumFormulaName) => void;
  runFormulaErrorChecking: () => void;

  // Sheet views / zoom -------------------------------------------------------
  saveCurrentSheetViewFromRibbon: () => Promise<unknown> | unknown;
  deleteActiveSheetViewFromRibbon: () => void;
  showZoomDialogFromRibbon: () => Promise<unknown> | unknown;
}

/**
 * Dispatches a ribbon command id through the playground's mix of declarative
 * tables and switch-driven handlers. Returns `true` when the id matched a
 * handler so the caller can `preventDefault()` on the originating click,
 * `false` when nothing matched and the caller should fall through.
 */
export const applyRibbonCommand = (id: string, deps: ApplyRibbonCommandDeps): boolean => {
  const i = deps.inst;
  if (!i) return false;
  const state = i.store.getState();
  const range = state.selection.range;
  const dialogOpener = RIBBON_DIALOG_OPENERS[id];
  if (dialogOpener) {
    dialogOpener(i);
    return true;
  }
  const fnName = RIBBON_FUNCTION_ARG_OPENERS[id];
  if (fnName) {
    i.openFunctionArguments(fnName);
    return true;
  }
  const formatMutator = RIBBON_FORMAT_MUTATORS[id];
  if (formatMutator) {
    deps.applyRibbonFormat(formatMutator);
    return true;
  }
  const zoomLevel = RIBBON_ZOOM_PRESETS[id];
  if (zoomLevel !== undefined) {
    setSheetZoom(i.store, zoomLevel, i.workbook);
    deps.refreshZoom();
    deps.focusSheet();
    return true;
  }
  const viewMode = RIBBON_VIEW_MODES[id];
  if (viewMode) {
    setWorkbookView(i.store, viewMode);
    deps.projectFormatToolbar();
    deps.focusSheet();
    return true;
  }
  const borderMode = RIBBON_BORDER_DRAW_MODES[id];
  if (borderMode) {
    if (i.borderDraw?.getMode() === borderMode) i.borderDraw.deactivate();
    else i.borderDraw?.activate(borderMode, deps.selectedBorderStyle, deps.selectedBorderColor);
    deps.projectFormatToolbar();
    deps.focusSheet();
    return true;
  }
  switch (id) {
    case 'pageBreaks':
      deps.applyPageBreakAction();
      return true;
    case 'pageTheme':
      deps.applyUiTheme(deps.uiTheme === 'dark' ? 'light' : 'dark');
      deps.focusSheet();
      return true;
    case 'sheetBackground':
      void deps.applySheetBackgroundAction('set');
      return true;
    case 'print':
    case 'printPageLayout':
      i.print('print');
      return true;
    case 'pdf':
      deps.applyPdfAction('create');
      return true;
    case 'inspect':
      deps.inspectWorkbookFromBackstage();
      return true;
    case 'paste':
      void deps.pasteClipboardIntoSelection();
      return true;
    case 'cut':
      void deps.cutSelectionToClipboard();
      return true;
    case 'copy':
      void deps.copySelectionToClipboard();
      return true;
    case 'undoHome':
      if (i.undo()) deps.focusSheet();
      return true;
    case 'redoHome':
      if (i.redo()) deps.focusSheet();
      return true;
    case 'fontGrow':
      deps.applyRibbonFormat((s, store) => {
        const a = s.selection.active;
        const f = s.format.formats.get(`${a.sheet}:${a.row}:${a.col}`);
        setFont(s, store, { fontSize: (f?.fontSize ?? 11) + 1 });
      });
      return true;
    case 'fontShrink':
      deps.applyRibbonFormat((s, store) => {
        const a = s.selection.active;
        const f = s.format.formats.get(`${a.sheet}:${a.row}:${a.col}`);
        setFont(s, store, { fontSize: Math.max(1, (f?.fontSize ?? 11) - 1) });
      });
      return true;
    case 'merge': {
      const anchorAt0 = state.merges.byAnchor.get(`${range.sheet}:${range.r0}:${range.c0}`);
      const isExactMerge =
        anchorAt0 &&
        range.r0 === anchorAt0.r0 &&
        range.c0 === anchorAt0.c0 &&
        range.r1 === anchorAt0.r1 &&
        range.c1 === anchorAt0.c1;
      if (isExactMerge) applyUnmerge(i.store, i.workbook, i.history, range);
      else applyMerge(i.store, i.workbook, i.history, range);
      deps.focusSheet();
      return true;
    }
    case 'formatPainter': {
      const fp = i.formatPainter;
      if (!fp) return true;
      if (fp.isActive()) fp.deactivate();
      else fp.activate(false);
      deps.projectFormatToolbar();
      deps.focusSheet();
      return true;
    }
    case 'formatTableHome':
      void deps.createTableFromSelection('medium');
      return true;
    case 'insertRows':
      insertRows(i.store, i.workbook, i.history, range.r0, deps.selectedRowCount());
      deps.refreshWorkbookCells();
      deps.focusSheet();
      return true;
    case 'deleteRows':
      deleteRows(i.store, i.workbook, i.history, range.r0, deps.selectedRowCount());
      deps.refreshWorkbookCells();
      deps.focusSheet();
      return true;
    case 'insertCols':
      insertCols(i.store, i.workbook, i.history, range.c0, deps.selectedColCount());
      deps.refreshWorkbookCells();
      deps.focusSheet();
      return true;
    case 'deleteCols':
      deleteCols(i.store, i.workbook, i.history, range.c0, deps.selectedColCount());
      deps.refreshWorkbookCells();
      deps.focusSheet();
      return true;
    case 'sortAscHome':
    case 'sortAsc':
    case 'sortFilterHome':
      deps.sortSelection('asc');
      return true;
    case 'sortDesc':
      deps.sortSelection('desc');
      return true;
    case 'sortData':
      void deps.customSortSelection();
      return true;
    case 'filterHome':
      deps.openFilterForSelection();
      return true;
    case 'outlineGroup':
      deps.applyOutlineAction('group');
      return true;
    case 'outlineUngroup':
      deps.applyOutlineAction('ungroup');
      return true;
    case 'outlineShowDetail':
      deps.applyOutlineAction('show-detail');
      return true;
    case 'outlineHideDetail':
      deps.applyOutlineAction('hide-detail');
      return true;
    case 'drawPen':
      i.borderDraw?.deactivate();
      deps.setDrawInkMode('pen');
      return true;
    case 'drawErase':
      i.borderDraw?.deactivate();
      deps.setDrawInkMode('erase');
      return true;
    case 'findHome':
    case 'findReview':
      i.openFindReplace();
      return true;
    case 'spellingReview':
      deps.runSpellingReview();
      return true;
    case 'translateReview':
      deps.openTranslateReview();
      return true;
    case 'accessibility':
      deps.runAccessibilityCheck();
      return true;
    case 'formatTableInsert':
      void deps.createTableFromSelection('medium');
      return true;
    case 'removeDupesInsert':
    case 'removeDupes':
      deps.removeDuplicateRows();
      return true;
    case 'textToColumns':
      deps.splitTextToColumns(',');
      return true;
    case 'dataValidation':
      i.openDataValidationDialog();
      return true;
    case 'chartInsert':
      deps.createChartFromSelection();
      return true;
    case 'pictureInsert':
      void deps.insertPictureFromRibbon('online');
      return true;
    case 'shapesInsert':
      deps.insertShapeFromRibbon('rectangle');
      return true;
    case 'screenshotInsert':
      deps.insertScreenshotFromRibbon();
      return true;
    case 'autosum':
    case 'autosumFormula': {
      deps.applyAutoSumFormula('SUM');
      return true;
    }
    case 'precedents':
      if (i.tracePrecedents() === 0) {
        void deps.showMessage({
          title: deps.ribbonText.formulaAuditing,
          message: deps.ribbonMenuText.traceNoPrecedents,
        });
      }
      return true;
    case 'dependents':
      if (i.traceDependents() === 0) {
        void deps.showMessage({
          title: deps.ribbonText.formulaAuditing,
          message: deps.ribbonMenuText.traceNoDependents,
        });
      }
      return true;
    case 'clearArrows':
      i.clearTraces();
      return true;
    case 'errorChecking':
      deps.runFormulaErrorChecking();
      return true;
    case 'showFormulasFormula':
      setShowFormulas(i.store, !i.store.getState().ui.showFormulas);
      deps.projectFormatToolbar();
      deps.focusSheet();
      return true;
    case 'recalcNow':
      i.recalc();
      deps.focusSheet();
      return true;
    case 'watch':
    case 'watchView':
      i.toggleWatchWindow();
      return true;
    case 'viewGridlines':
    case 'pageLayoutGridlinesView':
      setGridlinesVisible(i.store, i.store.getState().ui.showGridLines === false);
      deps.projectFormatToolbar();
      deps.focusSheet();
      return true;
    case 'pageLayoutGridlinesPrint': {
      const sheet = i.store.getState().data.sheetIndex;
      const setup = getPageSetup(i.store.getState(), sheet);
      setPrintGridlines(i.store, sheet, !setup.showGridlines, i.history);
      deps.projectFormatToolbar();
      deps.focusSheet();
      return true;
    }
    case 'viewHeadings':
    case 'pageLayoutHeadingsView':
      setHeadingsVisible(i.store, i.store.getState().ui.showHeaders === false);
      deps.projectFormatToolbar();
      deps.focusSheet();
      return true;
    case 'pageLayoutHeadingsPrint': {
      const sheet = i.store.getState().data.sheetIndex;
      const setup = getPageSetup(i.store.getState(), sheet);
      setPrintHeadings(i.store, sheet, !setup.showHeadings, i.history);
      deps.projectFormatToolbar();
      deps.focusSheet();
      return true;
    }
    case 'sheetViewSave':
      void deps.saveCurrentSheetViewFromRibbon();
      return true;
    case 'sheetViewDelete':
      deps.deleteActiveSheetViewFromRibbon();
      return true;
    case 'viewFormulas':
      setShowFormulas(i.store, !i.store.getState().ui.showFormulas);
      deps.projectFormatToolbar();
      deps.focusSheet();
      return true;
    case 'viewFormulaBar':
      deps.setFormulaBarVisible(!deps.formulaBarVisible);
      i.setFeatures(deps.playgroundFeatureFlags());
      deps.projectFormatToolbar();
      deps.focusSheet();
      return true;
    case 'viewR1C1':
      setR1C1ReferenceStyle(i.store, !i.store.getState().ui.r1c1);
      deps.projectFormatToolbar();
      deps.focusSheet();
      return true;
    case 'hideRows':
      hideRows(i.store, i.history, range.r0, range.r1, i.workbook);
      deps.focusSheet();
      return true;
    case 'hideCols':
      hideCols(i.store, i.history, range.c0, range.c1, i.workbook);
      deps.focusSheet();
      return true;
    case 'deleteCommentReview':
      deps.deleteActiveReviewComment();
      return true;
    case 'previousCommentReview':
      deps.selectReviewComment(-1);
      return true;
    case 'nextCommentReview':
      deps.selectReviewComment(1);
      return true;
    case 'protectReview':
    case 'protect':
      void deps.runSheetProtectionFlow();
      return true;
    case 'protectWorkbookReview':
      void deps.runWorkbookProtectionFlow(!isWorkbookStructureProtected(i.store.getState()));
      return true;
    case 'protectionReview':
      void deps.applyProtectAction('allow-edit-ranges');
      return true;
    case 'script':
      void deps.runPlaygroundScript();
      return true;
    case 'recordActions':
      deps.recordSelectedActions();
      return true;
    case 'allScripts':
      deps.openAllScripts();
      return true;
    case 'addIn':
      deps.openAddInManager();
      return true;
    case 'zoomSelection': {
      const selected = i.store.getState().selection.range;
      const rowCount = Math.max(1, selected.r1 - selected.r0 + 1);
      const colCount = Math.max(1, selected.c1 - selected.c0 + 1);
      const scaleForRows = 20 / rowCount;
      const scaleForCols = 12 / colCount;
      setSheetZoom(i.store, Math.max(0.5, Math.min(4, scaleForRows, scaleForCols)), i.workbook);
      deps.refreshZoom();
      deps.focusSheet();
      return true;
    }
    case 'zoomDialog':
      void deps.showZoomDialogFromRibbon();
      return true;
    default:
      return false;
  }
};
