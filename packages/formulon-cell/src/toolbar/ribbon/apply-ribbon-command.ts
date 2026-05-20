// Ribbon command dispatcher. Fans out a ribbon command id to the matching
// handler — first via the declarative dispatch tables (dialog openers,
// function-arg openers, format mutators, zoom presets, view modes, border-draw
// modes), then via the big switch for ids that need bespoke wiring.
//
// Deps are split into three groups so callers can wire them deliberately:
//  - flat required core (inst, text, ui state)
//  - `runtime` for toolbar-internal re-render / focus / format glue
//  - `hooks` for optional feature integrations the host opts into; missing
//    hook → the matching command id silently no-ops so consumers can ship a
//    minimal toolbar without every feature wired up.

import { setFont } from '../../commands/format.js';
import { applyMerge, applyUnmerge } from '../../commands/merge.js';
import { setPrintGridlines, setPrintHeadings } from '../../commands/page-setup.js';
import { isWorkbookStructureProtected } from '../../commands/protection.js';
import {
  deleteCols,
  deleteRows,
  hideCols,
  hideRows,
  insertCols,
  insertRows,
  setSheetZoom,
} from '../../commands/structure.js';
import {
  setGridlinesVisible,
  setHeadingsVisible,
  setR1C1ReferenceStyle,
  setShowFormulas,
  setWorkbookView,
} from '../../commands/view.js';
import type { FeatureFlags } from '../../extensions/index.js';
import type { SpreadsheetInstance } from '../../mount/types.js';
import { getPageSetup } from '../../store/store.js';
import type { CellBorderStyle } from '../../store/types.js';
import type { SessionShapeKind } from '../illustration-types.js';
import type { ToolbarMenuText } from '../menu-text.js';
import type { ToolbarText } from '../ribbon-model.js';
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

/** Toolbar-owned UI state read on each dispatch. Kept as plain values rather
 *  than getters because the dispatcher only reads them once per call. */
export interface RibbonUiState {
  theme: UiTheme;
  borderStyle: CellBorderStyle;
  borderColor: string;
  formulaBarVisible: boolean;
}

/** Toolbar-internal runtime callbacks. These are not optional — the toolbar
 *  can't function without re-render / focus / format-projection hooks. */
export interface RibbonRuntime {
  focusSheet: () => void;
  refreshCells: () => void;
  refreshZoom: () => void;
  projectFormatToolbar: () => void;
  applyRibbonFormat: (fn: RibbonFormatMutator) => void;
  applyUiTheme: (theme: UiTheme) => void;
  setFormulaBarVisible: (next: boolean) => void;
  featureFlags: () => FeatureFlags;
  showMessage: (opts: { title: string; message: string }) => Promise<void> | void;
}

/** Optional feature hooks grouped by category. Each group is independently
 *  opt-in: omitting a group (or any single field within a group) makes the
 *  matching command ids no-op. The dispatcher accesses every field through
 *  optional chaining, so partial wiring is safe. */
export interface RibbonHooks {
  clipboard?: {
    copy?: () => Promise<unknown> | unknown;
    cut?: () => Promise<unknown> | unknown;
    paste?: () => Promise<unknown> | unknown;
  };
  sortFilter?: {
    sort?: (dir: 'asc' | 'desc') => void;
    customSort?: () => Promise<unknown> | unknown;
    openFilter?: () => void;
    removeDuplicates?: () => void;
    splitTextToColumns?: (sep: string) => void;
    splitTextToColumnsCustom?: () => Promise<unknown> | unknown;
  };
  insert?: {
    createTable?: (variant: 'medium') => Promise<unknown> | unknown;
    createTableDialog?: () => Promise<unknown> | unknown;
    createChart?: () => void;
    createRecommendedChart?: () => Promise<unknown> | unknown;
    insertPicture?: (source: 'online') => Promise<unknown> | unknown;
    insertShape?: (shape: SessionShapeKind) => void;
    insertScreenshot?: () => void;
    insertSymbol?: (symbol: string) => Promise<unknown> | unknown;
  };
  page?: {
    pageBreak?: () => void;
    sheetBackground?: (action: 'set') => Promise<unknown> | unknown;
    pdf?: (action: 'create') => void;
    inspect?: () => void;
    outline?: (action: 'group' | 'ungroup' | 'show-detail' | 'hide-detail') => void;
  };
  drawing?: {
    setInkMode?: (mode: 'pen' | 'erase') => void;
  };
  review?: {
    spelling?: () => void;
    translate?: () => void;
    accessibility?: () => void;
    deleteComment?: () => void;
    selectComment?: (direction: 1 | -1) => void;
  };
  protection?: {
    runSheet?: () => Promise<unknown> | unknown;
    runWorkbook?: (protect: boolean) => Promise<unknown> | unknown;
    allowEditRanges?: () => Promise<unknown> | unknown;
  };
  automation?: {
    runScript?: () => Promise<unknown> | unknown;
    recordActions?: () => void;
    allScripts?: () => void;
    addInManager?: () => void;
  };
  formula?: {
    autoSum?: (fn: AutoSumFormulaName) => void;
    errorChecking?: () => void;
  };
  sheetView?: {
    save?: () => Promise<unknown> | unknown;
    deleteActive?: () => void;
    zoomDialog?: () => Promise<unknown> | unknown;
  };
}

/** Dependencies threaded into `applyRibbonCommand` on each call. */
export interface ApplyRibbonCommandDeps {
  inst: SpreadsheetInstance | null;
  text: ToolbarText;
  menuText: ToolbarMenuText;
  ui: RibbonUiState;
  runtime: RibbonRuntime;
  hooks?: RibbonHooks;
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
  const { ui, runtime, hooks } = deps;
  const state = i.store.getState();
  const range = state.selection.range;
  const rowCount = Math.max(1, range.r1 - range.r0 + 1);
  const colCount = Math.max(1, range.c1 - range.c0 + 1);
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
    runtime.applyRibbonFormat(formatMutator);
    return true;
  }
  const zoomLevel = RIBBON_ZOOM_PRESETS[id];
  if (zoomLevel !== undefined) {
    setSheetZoom(i.store, zoomLevel, i.workbook);
    runtime.refreshZoom();
    runtime.focusSheet();
    return true;
  }
  const viewMode = RIBBON_VIEW_MODES[id];
  if (viewMode) {
    setWorkbookView(i.store, viewMode);
    runtime.projectFormatToolbar();
    runtime.focusSheet();
    return true;
  }
  const borderMode = RIBBON_BORDER_DRAW_MODES[id];
  if (borderMode) {
    if (i.borderDraw?.getMode() === borderMode) i.borderDraw.deactivate();
    else i.borderDraw?.activate(borderMode, ui.borderStyle, ui.borderColor);
    runtime.projectFormatToolbar();
    runtime.focusSheet();
    return true;
  }
  switch (id) {
    case 'pageBreaks':
      hooks?.page?.pageBreak?.();
      return true;
    case 'pageTheme':
      runtime.applyUiTheme(ui.theme === 'dark' ? 'light' : 'dark');
      runtime.focusSheet();
      return true;
    case 'sheetBackground':
      void hooks?.page?.sheetBackground?.('set');
      return true;
    case 'print':
    case 'printPageLayout':
      i.print('print');
      return true;
    case 'pdf':
      hooks?.page?.pdf?.('create');
      return true;
    case 'inspect':
      hooks?.page?.inspect?.();
      return true;
    case 'paste':
      void hooks?.clipboard?.paste?.();
      return true;
    case 'cut':
      void hooks?.clipboard?.cut?.();
      return true;
    case 'copy':
      void hooks?.clipboard?.copy?.();
      return true;
    case 'undoHome':
      if (i.undo()) runtime.focusSheet();
      return true;
    case 'redoHome':
      if (i.redo()) runtime.focusSheet();
      return true;
    case 'fontGrow':
      runtime.applyRibbonFormat((s, store) => {
        const a = s.selection.active;
        const f = s.format.formats.get(`${a.sheet}:${a.row}:${a.col}`);
        setFont(s, store, { fontSize: (f?.fontSize ?? 11) + 1 });
      });
      return true;
    case 'fontShrink':
      runtime.applyRibbonFormat((s, store) => {
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
      runtime.focusSheet();
      return true;
    }
    case 'formatPainter': {
      const fp = i.formatPainter;
      if (!fp) return true;
      if (fp.isActive()) fp.deactivate();
      else fp.activate(false);
      runtime.projectFormatToolbar();
      runtime.focusSheet();
      return true;
    }
    case 'formatTableHome':
      void hooks?.insert?.createTable?.('medium');
      return true;
    case 'insertRows':
      insertRows(i.store, i.workbook, i.history, range.r0, rowCount);
      runtime.refreshCells();
      runtime.focusSheet();
      return true;
    case 'deleteRows':
      deleteRows(i.store, i.workbook, i.history, range.r0, rowCount);
      runtime.refreshCells();
      runtime.focusSheet();
      return true;
    case 'insertCols':
      insertCols(i.store, i.workbook, i.history, range.c0, colCount);
      runtime.refreshCells();
      runtime.focusSheet();
      return true;
    case 'deleteCols':
      deleteCols(i.store, i.workbook, i.history, range.c0, colCount);
      runtime.refreshCells();
      runtime.focusSheet();
      return true;
    case 'sortAscHome':
    case 'sortAsc':
    case 'sortFilterHome':
      hooks?.sortFilter?.sort?.('asc');
      return true;
    case 'sortDesc':
      hooks?.sortFilter?.sort?.('desc');
      return true;
    case 'sortData':
      void hooks?.sortFilter?.customSort?.();
      return true;
    case 'filterHome':
    case 'filter':
      hooks?.sortFilter?.openFilter?.();
      return true;
    case 'outlineGroup':
      hooks?.page?.outline?.('group');
      return true;
    case 'outlineUngroup':
      hooks?.page?.outline?.('ungroup');
      return true;
    case 'outlineShowDetail':
      hooks?.page?.outline?.('show-detail');
      return true;
    case 'outlineHideDetail':
      hooks?.page?.outline?.('hide-detail');
      return true;
    case 'drawPen':
      i.borderDraw?.deactivate();
      hooks?.drawing?.setInkMode?.('pen');
      return true;
    case 'drawErase':
      i.borderDraw?.deactivate();
      hooks?.drawing?.setInkMode?.('erase');
      return true;
    case 'findHome':
    case 'findReview':
      i.openFindReplace();
      return true;
    case 'spellingReview':
      hooks?.review?.spelling?.();
      return true;
    case 'translateReview':
      hooks?.review?.translate?.();
      return true;
    case 'accessibility':
      hooks?.review?.accessibility?.();
      return true;
    case 'formatTableInsert':
      if (hooks?.insert?.createTableDialog) void hooks.insert.createTableDialog();
      else void hooks?.insert?.createTable?.('medium');
      return true;
    case 'removeDupes':
      hooks?.sortFilter?.removeDuplicates?.();
      return true;
    case 'textToColumns':
      if (hooks?.sortFilter?.splitTextToColumnsCustom)
        void hooks.sortFilter.splitTextToColumnsCustom();
      else hooks?.sortFilter?.splitTextToColumns?.(',');
      return true;
    case 'chartInsert':
      if (hooks?.insert?.createRecommendedChart) void hooks.insert.createRecommendedChart();
      else hooks?.insert?.createChart?.();
      return true;
    case 'pictureInsert':
      void hooks?.insert?.insertPicture?.('online');
      return true;
    case 'shapesInsert':
      hooks?.insert?.insertShape?.('rectangle');
      return true;
    case 'screenshotInsert':
      hooks?.insert?.insertScreenshot?.();
      return true;
    case 'symbolInsert':
      void hooks?.insert?.insertSymbol?.('more');
      return true;
    case 'autosum':
    case 'autosumFormula':
      hooks?.formula?.autoSum?.('SUM');
      return true;
    case 'precedents':
      if (i.tracePrecedents() === 0) {
        void runtime.showMessage({
          title: deps.text.formulaAuditing,
          message: deps.menuText.traceNoPrecedents,
        });
      }
      return true;
    case 'dependents':
      if (i.traceDependents() === 0) {
        void runtime.showMessage({
          title: deps.text.formulaAuditing,
          message: deps.menuText.traceNoDependents,
        });
      }
      return true;
    case 'clearArrows':
      i.clearTraces();
      return true;
    case 'errorChecking':
      hooks?.formula?.errorChecking?.();
      return true;
    case 'showFormulasFormula':
      setShowFormulas(i.store, !i.store.getState().ui.showFormulas);
      runtime.projectFormatToolbar();
      runtime.focusSheet();
      return true;
    case 'recalcNow':
      i.recalc();
      runtime.focusSheet();
      return true;
    case 'watch':
    case 'watchView':
      i.toggleWatchWindow();
      return true;
    case 'viewGridlines':
    case 'pageLayoutGridlinesView':
      setGridlinesVisible(i.store, i.store.getState().ui.showGridLines === false);
      runtime.projectFormatToolbar();
      runtime.focusSheet();
      return true;
    case 'pageLayoutGridlinesPrint': {
      const sheet = i.store.getState().data.sheetIndex;
      const setup = getPageSetup(i.store.getState(), sheet);
      setPrintGridlines(i.store, sheet, !setup.showGridlines, i.history);
      runtime.projectFormatToolbar();
      runtime.focusSheet();
      return true;
    }
    case 'viewHeadings':
    case 'pageLayoutHeadingsView':
      setHeadingsVisible(i.store, i.store.getState().ui.showHeaders === false);
      runtime.projectFormatToolbar();
      runtime.focusSheet();
      return true;
    case 'pageLayoutHeadingsPrint': {
      const sheet = i.store.getState().data.sheetIndex;
      const setup = getPageSetup(i.store.getState(), sheet);
      setPrintHeadings(i.store, sheet, !setup.showHeadings, i.history);
      runtime.projectFormatToolbar();
      runtime.focusSheet();
      return true;
    }
    case 'sheetViewSave':
      void hooks?.sheetView?.save?.();
      return true;
    case 'sheetViewDelete':
      hooks?.sheetView?.deleteActive?.();
      return true;
    case 'viewFormulas':
      setShowFormulas(i.store, !i.store.getState().ui.showFormulas);
      runtime.projectFormatToolbar();
      runtime.focusSheet();
      return true;
    case 'viewFormulaBar':
      runtime.setFormulaBarVisible(!ui.formulaBarVisible);
      i.setFeatures(runtime.featureFlags());
      runtime.projectFormatToolbar();
      runtime.focusSheet();
      return true;
    case 'viewR1C1':
      setR1C1ReferenceStyle(i.store, !i.store.getState().ui.r1c1);
      runtime.projectFormatToolbar();
      runtime.focusSheet();
      return true;
    case 'hideRows':
      hideRows(i.store, i.history, range.r0, range.r1, i.workbook);
      runtime.focusSheet();
      return true;
    case 'hideCols':
      hideCols(i.store, i.history, range.c0, range.c1, i.workbook);
      runtime.focusSheet();
      return true;
    case 'deleteCommentReview':
      hooks?.review?.deleteComment?.();
      return true;
    case 'previousCommentReview':
      hooks?.review?.selectComment?.(-1);
      return true;
    case 'nextCommentReview':
      hooks?.review?.selectComment?.(1);
      return true;
    case 'protectReview':
    case 'protect':
      void hooks?.protection?.runSheet?.();
      return true;
    case 'protectWorkbookReview':
      void hooks?.protection?.runWorkbook?.(!isWorkbookStructureProtected(i.store.getState()));
      return true;
    case 'protectionReview':
      void hooks?.protection?.allowEditRanges?.();
      return true;
    case 'script':
      void hooks?.automation?.runScript?.();
      return true;
    case 'recordActions':
      hooks?.automation?.recordActions?.();
      return true;
    case 'allScripts':
      hooks?.automation?.allScripts?.();
      return true;
    case 'addIn':
      hooks?.automation?.addInManager?.();
      return true;
    case 'zoomSelection': {
      const selected = i.store.getState().selection.range;
      const selectedRows = Math.max(1, selected.r1 - selected.r0 + 1);
      const selectedCols = Math.max(1, selected.c1 - selected.c0 + 1);
      const scaleForRows = 20 / selectedRows;
      const scaleForCols = 12 / selectedCols;
      setSheetZoom(i.store, Math.max(0.5, Math.min(4, scaleForRows, scaleForCols)), i.workbook);
      runtime.refreshZoom();
      runtime.focusSheet();
      return true;
    }
    case 'zoomDialog':
      void hooks?.sheetView?.zoomDialog?.();
      return true;
    default:
      return false;
  }
};
