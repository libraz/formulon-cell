// Dispatch tables for the playground ribbon — each map is one row per ribbon
// command id, replacing what used to be a `case 'id': handler; return true;`
// triplet in `applyRibbonCommand`. Pulled out of main.ts so the dispatcher can
// stay short and the data sits in a single grep-friendly location.

import {
  bumpDecimals,
  bumpIndent,
  clearVisualFormat,
  cycleCurrency,
  cyclePercent,
  setAlign,
  setNumFmt,
  setRotation,
  setVAlign,
  toggleBold,
  toggleItalic,
  toggleStrike,
  toggleUnderline,
  toggleWrap,
} from '../../commands/format.js';
import type { SpreadsheetInstance } from '../../mount/types.js';
import type { SpreadsheetStore } from '../../store/store.js';

type RibbonState = ReturnType<SpreadsheetStore['getState']>;
export type RibbonFormatMutator = (state: RibbonState, store: SpreadsheetStore) => void;

/** Plain "open this dialog / panel" buttons. Each value is a thunk over the
 *  instance to keep the lookup callable with one argument from the dispatcher. */
export const RIBBON_DIALOG_OPENERS: Readonly<Record<string, (i: SpreadsheetInstance) => void>> = {
  pageSetup: (i) => i.openPageSetup(),
  pageSetupAdvanced: (i) => i.openPageSetup(),
  printTitles: (i) => i.openPageSetup(),
  formatCells: (i) => i.openFormatDialog(),
  formatCellsHome: (i) => i.openFormatDialog(),
  moreBorders: (i) => i.openFormatDialog(),
  windowVisibility: (i) => i.openFormatDialog('more'),
  gotoSpecial: (i) => i.openGoToSpecial(),
  gotoSpecialHome: (i) => i.openGoToSpecial(),
  hyperlinkInsert: (i) => i.openHyperlinkDialog(),
  commentInsert: (i) => i.openCommentDialog(),
  newCommentReview: (i) => i.openCommentDialog(),
  links: (i) => i.openExternalLinksDialog(),
  linksData: (i) => i.openExternalLinksDialog(),
  conditional: (i) => i.openConditionalDialog(),
  cellStyles: (i) => i.openCellStylesGallery(),
  rules: (i) => i.openCfRulesDialog(),
  pivotTableInsert: (i) => i.openPivotTableDialog(),
  namedRanges: (i) => i.openNamedRangeDialog(),
  evaluateFormula: (i) => i.openEvaluateFormulaDialog(),
  fxInsert: (i) => i.openFunctionArguments(),
  fx: (i) => i.openFunctionArguments(),
  workbookObjectsView: (i) => i.openWorkbookObjects(),
  arrangeObjectsPageLayout: (i) => i.openWorkbookObjects(),
  selectionPanePageLayout: (i) => i.openWorkbookObjects(),
  pivotFieldListView: (i) => {
    if (!i.openActivePivotFieldList()) i.openWorkbookObjects();
  },
  calcOptions: (i) => i.openIterativeDialog(),
};

/** Ribbon command id → spreadsheet function name to prefill the Function
 *  Arguments dialog with. Kept separate from [[RIBBON_DIALOG_OPENERS]] so the
 *  dispatcher can pass the captured string straight through. */
export const RIBBON_FUNCTION_ARG_OPENERS: Readonly<Record<string, string>> = {
  sum: 'SUM',
  avg: 'AVERAGE',
  ifFormula: 'IF',
  xlookupFormula: 'XLOOKUP',
  concatFormula: 'CONCAT',
  todayFormula: 'TODAY',
  pmtFormula: 'PMT',
  roundFormula: 'ROUND',
};

/** Format-toggle buttons. The dispatcher wraps these in `applyRibbonFormat`
 *  so the mutation lands in one undoable history entry and refocuses the
 *  sheet on the way out — entries here stay declarative. */
export const RIBBON_FORMAT_MUTATORS: Readonly<Record<string, RibbonFormatMutator>> = {
  bold: (s, store) => toggleBold(s, store),
  italic: (s, store) => toggleItalic(s, store),
  underline: (s, store) => toggleUnderline(s, store),
  strike: (s, store) => toggleStrike(s, store),
  currency: (s, store) => cycleCurrency(s, store),
  percent: (s, store) => cyclePercent(s, store),
  comma: (s, store) => setNumFmt(s, store, { kind: 'fixed', decimals: 2, thousands: true }),
  alignL: (s, store) => setAlign(s, store, 'left'),
  alignC: (s, store) => setAlign(s, store, 'center'),
  alignR: (s, store) => setAlign(s, store, 'right'),
  top: (s, store) => setVAlign(s, store, 'top'),
  middle: (s, store) => setVAlign(s, store, 'middle'),
  bottomAlign: (s, store) => setVAlign(s, store, 'bottom'),
  decUp: (s, store) => bumpDecimals(s, store, 1),
  decDown: (s, store) => bumpDecimals(s, store, -1),
  wrap: (s, store) => toggleWrap(s, store),
  clearFormat: (s, store) => clearVisualFormat(s, store),
  general: (s, store) => setNumFmt(s, store, { kind: 'general' }),
  textOrientation: (s, store) => setRotation(s, store, 45),
  indentDecrease: (s, store) => bumpIndent(s, store, -1),
  indentIncrease: (s, store) => bumpIndent(s, store, 1),
};

/** Zoom-preset buttons — the dispatcher applies the value via setSheetZoom and
 *  refreshes the zoom indicator. */
export const RIBBON_ZOOM_PRESETS: Readonly<Record<string, number>> = {
  zoom75: 0.75,
  zoom100: 1,
  zoom125: 1.25,
};

export type RibbonViewMode = 'normal' | 'pageLayout' | 'pageBreakPreview';

/** Workbook-view selector buttons. */
export const RIBBON_VIEW_MODES: Readonly<Record<string, RibbonViewMode>> = {
  viewNormal: 'normal',
  viewPageLayout: 'pageLayout',
  viewPageBreakPreview: 'pageBreakPreview',
};

export type RibbonBorderDrawMode = 'draw' | 'grid' | 'erase';

/** Border-draw toggle buttons. The dispatcher flips the mode off when it is
 *  already active so each button is a toggle. */
export const RIBBON_BORDER_DRAW_MODES: Readonly<Record<string, RibbonBorderDrawMode>> = {
  drawBorder: 'draw',
  drawBorderGrid: 'grid',
  drawGrid: 'grid',
  eraseBorder: 'erase',
};
