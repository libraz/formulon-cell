import type { Range } from '../engine/types.js';
import type { SpreadsheetStore } from '../store/store.js';
import { type BuiltInCommandId, isBuiltInCommandId } from './built-in-command-ids.js';
import { getRecentFunctions } from './function-history.js';
import { interactionControllerFor } from './interaction-controller.js';
import { columnShiftBand, rowShiftBand } from './interaction-effects.js';
import type {
  InteractionOperation,
  InteractionOrigin,
  OperationEffect,
  OperationIntent,
  PermissionDecision,
} from './interaction-policy.js';
import {
  refusedOnProtectedSheet,
  sheetPermits,
  unlockedCellsOption,
  workbookStructurePermits,
} from './protection.js';
import { FUNCTION_SIGNATURES } from './refs.js';

/** A built-in command id, or a Mac function-menu leaf naming a spreadsheet function. */
export type BuiltInCommandRequest = BuiltInCommandId | `mac.function.${string}`;

/** The operation each built-in command performs, or `null` for a command that never mutates the
 * workbook and is always allowed. Total over {@link BuiltInCommandId}, so a new command cannot
 * ship without a classification, and an unknown id is never executable. */
export const BUILT_IN_COMMAND_OPERATION: Readonly<
  Record<BuiltInCommandId, InteractionOperation | null>
> = {
  paste: 'paste',
  pasteSpecial: 'paste',
  pasteAll: 'paste',
  pasteValues: 'paste',
  pasteValuesNumFmt: 'paste',
  pasteTranspose: 'paste',
  pasteFormatsOnly: 'format',
  pasteFormulas: 'formulaEdit',
  pasteFormulasNumFmt: 'formulaEdit',
  cut: 'moveCells',
  clear: 'clear',
  insertCells: 'insertRows',
  insertCopiedCells: 'moveCells',
  deleteCells: 'deleteRows',
  insertRows: 'insertRows',
  deleteRows: 'deleteRows',
  rowInsertAbove: 'insertRows',
  rowInsertBelow: 'insertRows',
  rowDelete: 'deleteRows',
  colInsertLeft: 'insertColumns',
  colInsertRight: 'insertColumns',
  colDelete: 'deleteColumns',
  rowHeight: 'resizeRows',
  colWidth: 'resizeColumns',
  bold: 'format',
  italic: 'format',
  underline: 'format',
  strike: 'format',
  alignLeft: 'format',
  alignCenter: 'format',
  alignRight: 'format',
  // Home ribbon format mutators and the inline font/number controls.
  currency: 'format',
  percent: 'format',
  comma: 'format',
  alignL: 'format',
  alignC: 'format',
  alignR: 'format',
  top: 'format',
  middle: 'format',
  bottomAlign: 'format',
  decUp: 'format',
  decDown: 'format',
  wrap: 'format',
  textOrientation: 'format',
  indentDecrease: 'format',
  indentIncrease: 'format',
  fontGrow: 'format',
  fontShrink: 'format',
  fontFamily: 'format',
  fontSize: 'format',
  fontColor: 'format',
  fillColor: 'format',
  numberFormat: 'format',
  borders: 'format',
  formatCells: 'format',
  defineName: 'namedRange',
  filterClear: 'filter',
  filterReapply: 'filter',
  filterByValue: 'filter',
  sortAsc: 'sort',
  sortDesc: 'sort',
  insertComment: 'comment',
  deleteComment: 'comment',
  insertHyperlink: 'hyperlink',
  showPhonetic: 'format',
  editPhonetic: 'format',
  protect: 'protection',
  protectReview: 'protection',
  pageSetup: 'pageSetup',
  print: 'print',
  printArea: 'pageSetup',
  pageBreaks: 'pageSetup',
  dataValidation: 'validation',
  namedRanges: 'namedRange',
  formatTableInsert: 'table',
  formatTableHome: 'table',
  pivotTableInsert: 'pivot',
  tableName: 'table',
  tableHeaderRow: 'table',
  tableTotalRow: 'table',
  tableBandedRows: 'table',
  tableFirstColumn: 'table',
  tableLastColumn: 'table',
  tableBandedColumns: 'table',
  chartInsert: 'object',
  pictureInsert: 'object',
  shapesInsert: 'object',
  screenshotInsert: 'object',
  workbookObjectsView: null,
  chartAddElement: 'object',
  chartColors: 'object',
  chartSwitchRowColumn: 'object',
  chartSelectData: 'object',
  chartChangeType: 'object',
  chartFormatSelection: 'object',
  objectBringForward: 'object',
  objectSendBackward: 'object',
  objectSelectionPane: null,
  shapeFill: 'object',
  shapeOutline: 'object',
  scaleWidth: 'pageSetup',
  scaleHeight: 'pageSetup',
  sheetViewSelect: 'filter',

  // Excel 365 Mac ribbon actions. Keep these aliases explicit so a restricted
  // embed cannot gain a new mutation merely by adopting the Mac profile.
  'mac.insert.pivotTable': 'pivot',
  'mac.insert.recommendedPivotTable': 'pivot',
  'mac.insert.table': 'table',
  'mac.insert.photo': 'object',
  'mac.insert.shapeRectangle': 'object',
  'mac.insert.shapeArrow': 'object',
  'mac.insert.shapeLine': 'object',
  'mac.insert.shapeRoundedRectangle': 'object',
  'mac.insert.shapeOval': 'object',
  'mac.insert.shapeTriangle': 'object',
  'mac.insert.shapeDiamond': 'object',
  'mac.insert.screenshot': 'object',
  'mac.insert.recommendedChart': 'object',
  'mac.insert.chartColumn': 'object',
  'mac.insert.chartColumnStacked': 'object',
  'mac.insert.chartColumn100': 'object',
  'mac.insert.chartBarStacked': 'object',
  'mac.insert.chartBar100': 'object',
  'mac.insert.chartLineStacked': 'object',
  'mac.insert.chartLine100': 'object',
  'mac.insert.chartAreaStacked': 'object',
  'mac.insert.chartArea100': 'object',
  'mac.insert.chartBar': 'object',
  'mac.insert.chartLine': 'object',
  'mac.insert.chartArea': 'object',
  'mac.insert.chartPie': 'object',
  'mac.insert.chartDoughnut': 'object',
  'mac.insert.chartScatter': 'object',
  'mac.insert.chartHierarchy': 'object',
  'mac.insert.chartStatistical': 'object',
  'mac.insert.chartWaterfall': 'object',
  'mac.insert.chartCombo': 'object',
  'mac.insert.chartMap': 'object',
  'mac.insert.sparkline': 'object',
  'mac.insert.slicer': 'table',
  'mac.insert.link': 'hyperlink',
  'mac.insert.newComment': 'comment',
  'mac.insert.headerFooter': 'pageSetup',
  'mac.home.addins': null,
  'mac.draw.toggle': 'object',
  'mac.draw.penBlack': 'object',
  'mac.draw.penRed': 'object',
  'mac.draw.pencil': 'object',
  'mac.draw.highlighter': 'object',
  'mac.draw.trackpad': 'object',
  'mac.page.theme.office': 'format',
  'mac.page.theme.gray': 'format',
  'mac.page.color.office': 'format',
  'mac.page.color.gray': 'format',
  'mac.page.font.office': 'format',
  'mac.page.font.gray': 'format',
  'mac.page.background': 'pageSetup',
  'mac.page.printTitles': 'pageSetup',
  'mac.page.pageSetup': 'pageSetup',
  'mac.page.showGridlines': null,
  'mac.page.printGridlines': 'pageSetup',
  'mac.page.showHeadings': null,
  'mac.page.printHeadings': 'pageSetup',
  'mac.formulas.insertFunction': 'formulaEdit',
  'mac.autosum.SUM': 'formulaEdit',
  'mac.autosum.AVERAGE': 'formulaEdit',
  'mac.autosum.COUNT': 'formulaEdit',
  'mac.autosum.MAX': 'formulaEdit',
  'mac.autosum.MIN': 'formulaEdit',
  'mac.formulas.namesManager': 'namedRange',
  'mac.formulas.defineName': 'namedRange',
  'mac.formulas.useInFormula': 'formulaEdit',
  'mac.formulas.createFromSelection': 'namedRange',
  'mac.formulas.removeArrows.all': null,
  'mac.formulas.removeArrows.precedents': null,
  'mac.formulas.removeArrows.dependents': null,
  'mac.formulas.errorCheck.trace': null,
  'mac.formulas.errorCheck.ignore': 'formulaEdit',
  'mac.formulas.precedents': null,
  'mac.formulas.dependents': null,
  'mac.formulas.recalc': null,
  'mac.formulas.sheetRecalc': null,
  'mac.page.margins.normal': 'pageSetup',
  'mac.page.margins.wide': 'pageSetup',
  'mac.page.margins.narrow': 'pageSetup',
  'mac.page.margins.custom': 'pageSetup',
  'mac.page.orientation.portrait': 'pageSetup',
  'mac.page.orientation.landscape': 'pageSetup',
  'mac.page.size.a4': 'pageSetup',
  'mac.page.size.a3': 'pageSetup',
  'mac.page.size.letter': 'pageSetup',
  'mac.page.size.legal': 'pageSetup',
  'mac.page.printArea.set': 'pageSetup',
  'mac.page.printArea.add': 'pageSetup',
  'mac.page.printArea.clear': 'pageSetup',
  'mac.page.break.insert': 'pageSetup',
  'mac.page.break.remove': 'pageSetup',
  'mac.page.break.reset': 'pageSetup',
  'mac.formulas.createNames.topRow': 'namedRange',
  'mac.formulas.createNames.bottomRow': 'namedRange',
  'mac.formulas.createNames.leftColumn': 'namedRange',
  'mac.formulas.createNames.rightColumn': 'namedRange',
  'mac.formulas.calc.auto': null,
  'mac.formulas.calc.autoNoTable': null,
  'mac.formulas.calc.manual': null,
  'mac.formulas.calc.iterative': null,
  'mac.data.sortAsc': 'sort',
  'mac.data.sortDesc': 'sort',
  'mac.data.sortCustom': 'sort',
  'mac.data.filter': 'filter',
  'mac.data.clear': 'filter',
  'mac.data.reapply': 'filter',
  'mac.data.advancedFilter': 'filter',
  'mac.data.textToColumns': 'valueEdit',
  'mac.data.flashFill': 'fill',
  'mac.data.removeDuplicates': 'valueEdit',
  'mac.data.validation.settings': 'validation',
  'mac.data.validation.circleInvalid': 'validation',
  'mac.data.validation.clearCircles': 'validation',
  'mac.data.validation.clearRules': 'validation',
  'mac.data.consolidate': 'valueEdit',
  'mac.data.whatIf': null,
  'mac.data.goalSeek': 'valueEdit',
  'mac.data.dataTable': 'valueEdit',
  'mac.data.subtotal': 'valueEdit',
  'mac.data.group': 'resizeRows',
  'mac.data.ungroup': 'resizeRows',
  'mac.data.showDetail': 'resizeRows',
  'mac.data.hideDetail': 'resizeRows',
  'mac.review.newComment': 'comment',
  'mac.review.deleteComment': 'comment',
  'mac.review.newNote': 'comment',
  'mac.review.protectSheet': 'protection',
  'mac.review.protectWorkbook': 'protection',
  'mac.view.sheetViewSave': 'sheet',
  'mac.view.sheetViewDelete': 'sheet',
  'mac.automate.showScripts': null,
  'mac.automate.allRowsColumns': 'resizeRows',
  'mac.automate.makeSubtable': 'table',
  'mac.automate.removeHyperlinks': 'clear',
  'mac.automate.countEmptyRows': null,
  'mac.automate.tableToJson': 'export',
  'mac.automate.newPivotTable': 'pivot',

  // Never mutate the workbook: navigation, view state, zoom, read-only panes, menu roots and
  // launchers whose own writes go through the controller. Selection-moving ids are listed here
  // too and follow SELECTING_NAVIGATION_COMMANDS.
  accessibility: null,
  addIn: null,
  allScripts: null,
  arrangeObjectsPageLayout: null,
  calcOptions: null,
  cellStyles: null,
  conditional: null,
  copy: null,
  drawErase: null,
  errorChecking: null,
  evaluateFormula: null,
  fillHome: null,
  findHome: null,
  findReview: null,
  formatCellsHome: null,
  freeze: null,
  inspect: null,
  linksData: null,
  'mac.automate.freezeSelection': null,
  'mac.automate.gallery': null,
  'mac.data.dataTypes': null,
  'mac.data.pictureFromData': null,
  'mac.data.validation': null,
  'mac.data.workbookLinks': null,
  'mac.draw.add': null,
  'mac.draw.eraser': null,
  'mac.formulas.autoSum': null,
  'mac.formulas.calcOptions': null,
  'mac.formulas.category.all': null,
  'mac.formulas.category.recent': null,
  'mac.formulas.dateTime': null,
  'mac.formulas.errorCheck': null,
  'mac.formulas.errorCheck.run': null,
  'mac.formulas.financial': null,
  'mac.formulas.logical': null,
  'mac.formulas.lookup': null,
  'mac.formulas.math': null,
  'mac.formulas.more': null,
  'mac.formulas.recent': null,
  'mac.formulas.removeArrows': null,
  'mac.formulas.showFormulas': null,
  'mac.formulas.text': null,
  'mac.formulas.watch': null,
  'mac.insert.shapes': null,
  'mac.page.color': null,
  'mac.page.font': null,
  'mac.page.margins': null,
  'mac.page.orientation': null,
  'mac.page.pageBreaks': null,
  'mac.page.printArea': null,
  'mac.page.size': null,
  'mac.page.theme': null,
  'mac.review.accessibility': null,
  'mac.review.nextComment': null,
  'mac.review.nextNote': null,
  'mac.review.notes': null,
  'mac.review.previousComment': null,
  'mac.review.previousNote': null,
  'mac.review.showComments': null,
  'mac.review.showHideNote': null,
  'mac.review.showNotes': null,
  'mac.review.spelling': null,
  'mac.review.stats': null,
  'mac.view.formulaBar': null,
  'mac.view.freeze': null,
  'mac.view.freeze.firstColumn': null,
  'mac.view.freeze.firstRow': null,
  'mac.view.gridlines': null,
  'mac.view.headings': null,
  'mac.view.pageBreakPreview': null,
  'mac.view.pageLayout': null,
  'mac.view.split': null,
  'mac.view.standard': null,
  'mac.view.zeros': null,
  'mac.view.zoom': null,
  'mac.view.zoom100': null,
  'mac.view.zoomSelection': null,
  nextCommentReview: null,
  openHyperlink: null,
  outlineSymbols: null,
  pageLayoutGridlinesView: null,
  pageLayoutHeadingsView: null,
  pageTheme: null,
  pivotBlankRows: null,
  pivotClear: null,
  pivotFieldsItemsSets: null,
  pivotGrandTotals: null,
  pivotReportLayout: null,
  pivotSelect: null,
  pivotSelectData: null,
  pivotSelectEnable: null,
  pivotSelectLabels: null,
  pivotSelectLabelsAndData: null,
  pivotSelectTable: null,
  pivotSubtotals: null,
  previousCommentReview: null,
  recalcNow: null,
  recalcSheet: null,
  recordActions: null,
  script: null,
  selectAll: null,
  selectionPanePageLayout: null,
  showAllNotesReview: null,
  showFormulasFormula: null,
  showHideNote: null,
  showHideNoteReview: null,
  spellingReview: null,
  split: null,
  toggleWatch: null,
  translateReview: null,
  viewFormulaBar: null,
  viewFormulas: null,
  viewGridlines: null,
  viewHeadings: null,
  viewNormal: null,
  viewPageBreakPreview: null,
  viewPageLayout: null,
  viewR1C1: null,
  viewRightToLeft: null,
  viewZeros: null,
  watch: null,
  watchView: null,
  whatIf: null,
  zoom100: null,
  zoom125: null,
  zoom75: null,
  zoomDialog: null,
  zoomSelection: null,

  // Generic ribbon, context-menu and pivot commands.
  autosum: 'formulaEdit',
  autosumFormula: 'formulaEdit',
  avg: 'formulaEdit',
  clearArrows: null,
  clearFormat: 'format',
  colGroup: 'resizeColumns',
  colHide: 'resizeColumns',
  colUngroup: 'resizeColumns',
  colUnhide: 'resizeColumns',
  commentInsert: 'comment',
  concatFormula: 'formulaEdit',
  deleteCommentReview: 'comment',
  dependents: null,
  drawGrid: 'format',
  drawPen: 'object',
  filter: 'filter',
  formatPainter: 'format',
  fx: 'formulaEdit',
  hyperlinkInsert: 'hyperlink',
  ifFormula: 'formulaEdit',
  insertTableColumns: 'insertColumns',
  insertTableRows: 'insertRows',
  'mac.insert.recommendedChart.apply': 'object',
  'mac.insert.symbol': 'valueEdit',
  marginsPreset: 'pageSetup',
  merge: 'merge',
  newCommentReview: 'comment',
  orientationPreset: 'pageSetup',
  outlineGroup: 'resizeRows',
  outlineHideDetail: 'resizeRows',
  outlineShowDetail: 'resizeRows',
  outlineSummaryBelow: 'resizeRows',
  outlineSummaryRight: 'resizeColumns',
  outlineUngroup: 'resizeRows',
  pageLayoutGridlinesPrint: 'pageSetup',
  pageLayoutHeadingsPrint: 'pageSetup',
  pageSetupAdvanced: 'pageSetup',
  paperSizePreset: 'pageSetup',
  pasteColumnWidths: 'resizeColumns',
  pasteKeepSource: 'paste',
  pasteLink: 'formulaEdit',
  pasteNoBorders: 'paste',
  pdf: 'export',
  pivotActiveField: 'pivot',
  pivotCalcField: 'pivot',
  pivotCalcItem: 'pivot',
  pivotColHeaders: 'pivot',
  pivotColStripes: 'pivot',
  pivotCollapse: 'pivot',
  pivotCollapseField: 'pivot',
  pivotCollapseTo: 'pivot',
  pivotDelete: 'pivot',
  pivotExpand: 'pivot',
  pivotExpandField: 'pivot',
  pivotFieldList: 'pivot',
  pivotFieldListView: 'pivot',
  pivotFieldSettings: 'pivot',
  pivotFilterClear: 'pivot',
  pivotFilterHideSelected: 'pivot',
  pivotFilterKeepSelected: 'pivot',
  pivotFilterLabel: 'pivot',
  pivotFilterTop10: 'pivot',
  pivotFilterValue: 'pivot',
  pivotGroup: 'pivot',
  pivotListFormulas: 'sheet',
  pivotMove: 'pivot',
  pivotMoveFieldFirst: 'pivot',
  pivotMoveFieldLast: 'pivot',
  pivotMoveFieldLeft: 'pivot',
  pivotMoveFieldRight: 'pivot',
  pivotMoveFieldToColumns: 'pivot',
  pivotMoveFieldToRows: 'pivot',
  pivotMoveItemDown: 'pivot',
  pivotMoveItemFirst: 'pivot',
  pivotMoveItemLast: 'pivot',
  pivotMoveItemUp: 'pivot',
  pivotName: 'pivot',
  pivotNumberFormat: 'pivot',
  pivotOptions: 'pivot',
  pivotRefresh: 'pivot',
  pivotRemoveField: 'pivot',
  pivotRemoveGrandTotal: 'pivot',
  pivotRemoveValues: 'pivot',
  pivotRowHeaders: 'pivot',
  pivotRowStripes: 'pivot',
  pivotShowAsMore: 'pivot',
  pivotShowAsNormal: 'pivot',
  pivotShowAsPercentOfColumn: 'pivot',
  pivotShowAsPercentOfParentColumn: 'pivot',
  pivotShowAsPercentOfParentRow: 'pivot',
  pivotShowAsPercentOfRow: 'pivot',
  pivotShowAsPercentOfTotal: 'pivot',
  pivotShowDetails: 'sheet',
  pivotShowFieldList: 'pivot',
  pivotSolveOrder: 'pivot',
  pivotSortAsc: 'pivot',
  pivotSortDesc: 'pivot',
  pivotSortMore: 'pivot',
  pivotStyleGallery: 'pivot',
  pivotSubtotal: 'pivot',
  pivotSummarizeAverage: 'pivot',
  pivotSummarizeCount: 'pivot',
  pivotSummarizeDistinctCount: 'pivot',
  pivotSummarizeMax: 'pivot',
  pivotSummarizeMin: 'pivot',
  pivotSummarizeMore: 'pivot',
  pivotSummarizeProduct: 'pivot',
  pivotSummarizeSum: 'pivot',
  pivotUngroup: 'pivot',
  pmtFormula: 'formulaEdit',
  precedents: null,
  printPageLayout: 'print',
  printTitles: 'pageSetup',
  protectWorkbookReview: 'protection',
  protectionReview: 'protection',
  removeDupes: 'valueEdit',
  roundFormula: 'formulaEdit',
  rowGroup: 'resizeRows',
  rowHide: 'resizeRows',
  rowUngroup: 'resizeRows',
  rowUnhide: 'resizeRows',
  scalePercent: 'pageSetup',
  sheetBackground: 'pageSetup',
  // Sheet-tab commands: governed by workbook-structure protection.
  insertSheet: 'sheet',
  deleteSheet: 'sheet',
  renameSheet: 'sheet',
  moveSheet: 'sheet',
  copySheet: 'sheet',
  hideSheet: 'sheet',
  unhideSheet: 'sheet',
  sheetTabColor: 'sheet',
  sheetViewDelete: 'sheet',
  sheetViewSave: 'sheet',
  sortData: 'sort',
  sortFilterHome: 'sort',
  sum: 'formulaEdit',
  symbolInsert: 'valueEdit',
  tableFilterButton: 'filter',
  textToColumns: 'valueEdit',
  todayFormula: 'formulaEdit',
  windowVisibility: 'format',
  xlookupFormula: 'formulaEdit',
};

/** Commands whose useful effect is moving the active selection. These follow
 * the host's separate selection permission instead of the mutation policy. */
const SELECTING_NAVIGATION_COMMANDS: ReadonlySet<BuiltInCommandId> = new Set<BuiltInCommandId>([
  'selectAll',
  'errorChecking',
  'mac.formulas.errorCheck.run',
  'nextCommentReview',
  'previousCommentReview',
  'mac.review.previousComment',
  'mac.review.nextComment',
  'mac.review.previousNote',
  'mac.review.nextNote',
  'pivotSelectData',
  'pivotSelectLabels',
  'pivotSelectLabelsAndData',
  'pivotSelectTable',
]);

/** Operations whose writes touch no cell: decided by the operation alone. */
const WORKBOOK_EFFECT_OPERATIONS: ReadonlySet<InteractionOperation> = new Set([
  'resizeRows',
  'resizeColumns',
  'namedRange',
  'sheet',
  'object',
  'protection',
  'print',
  'export',
]);

/** Commands that turn AutoFilter on or off rather than use an existing filter. */
const FILTER_DEFINITION_COMMANDS: ReadonlySet<BuiltInCommandId> = new Set<BuiltInCommandId>([
  'filter',
  'mac.data.filter',
  'tableFilterButton',
]);

const FUNCTION_COMMAND_PREFIX = 'mac.function.';

const shiftedBand = (operation: InteractionOperation, selection: Range): Range | null => {
  if (operation === 'insertRows' || operation === 'deleteRows')
    return rowShiftBand(selection.sheet, Math.min(selection.r0, selection.r1));
  if (operation === 'insertColumns' || operation === 'deleteColumns')
    return columnShiftBand(selection.sheet, Math.min(selection.c0, selection.c1));
  return null;
};

const isBuiltInCommandRequest = (id: string): id is BuiltInCommandRequest =>
  isBuiltInCommandId(id) || id.startsWith(FUNCTION_COMMAND_PREFIX);

const invalidCommand = (id: string): PermissionDecision => ({
  allowed: false,
  code: 'invalid',
  reason: `${id} is not a built-in command`,
});

/** A function leaf edits a formula when its name is a known function or was recorded from the
 * live catalog; any other name is not a command. */
const functionLeafOperation = (
  store: SpreadsheetStore,
  name: string,
): InteractionOperation | undefined =>
  name in FUNCTION_SIGNATURES || getRecentFunctions(store, new Set([name])).includes(name)
    ? 'formulaEdit'
    : undefined;

/** The intent a built-in command declares: its operation and the cells it writes. `null` for a
 * command that never mutates, `undefined` for an id that names no built-in command. */
export const builtInCommandIntent = (
  store: SpreadsheetStore,
  id: string,
  origin: InteractionOrigin,
): OperationIntent | null | undefined => {
  if (!isBuiltInCommandRequest(id)) return undefined;
  const operation = isBuiltInCommandId(id)
    ? BUILT_IN_COMMAND_OPERATION[id]
    : functionLeafOperation(store, id.slice(FUNCTION_COMMAND_PREFIX.length));
  if (operation === undefined || operation === null) return operation;
  const workbookTheme = /^mac\.page\.(theme|color|font)\.(office|gray)$/.test(id);
  const state = store.getState();
  const selection = state.selection.range;
  const band = shiftedBand(operation, selection);
  let effects: readonly OperationEffect[];
  if (WORKBOOK_EFFECT_OPERATIONS.has(operation) || workbookTheme || id === 'sheetViewSelect')
    effects = [{ kind: 'workbook' }];
  // Turning AutoFilter on or off is a sheet-level filter change; using a filter is decided on its range.
  else if (isBuiltInCommandId(id) && FILTER_DEFINITION_COMMANDS.has(id))
    effects = [{ kind: 'sheet', sheet: selection.sheet }];
  // A sheet-level write, or one whose Protect Sheet option alone makes it available (locked
  // cells refuse it at run time): decided for the active sheet as a whole.
  else if (operation === 'pageSetup' || unlockedCellsOption(operation))
    effects = [{ kind: 'sheet', sheet: selection.sheet }];
  else if (band) effects = [{ kind: 'range', range: band }];
  else if (operation === 'format')
    effects = [selection, ...(state.selection.extraRanges ?? [])].map(
      (range): OperationEffect => ({ kind: 'range', range }),
    );
  else effects = [{ kind: 'range', range: selection }];
  return { operation, origin, commandId: id, effects };
};

/** Protection's decision for a built-in command on the active sheet, with or without a policy:
 *  the always-refused commands, an option-gated operation whose option is off, and a sheet
 *  operation under workbook-structure protection. Commands it does not name are left to their
 *  operation's own protection check. */
export const builtInProtectionDecision = (
  store: SpreadsheetStore,
  id: string,
): PermissionDecision => {
  if (!isBuiltInCommandId(id)) return { allowed: true };
  const state = store.getState();
  const sheet = state.data.sheetIndex;
  const operation = BUILT_IN_COMMAND_OPERATION[id];
  if (operation !== null && !workbookStructurePermits(state, operation))
    return { allowed: false, code: 'protected', reason: 'workbook structure is protected' };
  const option = operation ? unlockedCellsOption(operation) : undefined;
  // Turning AutoFilter on or off is refused on a protected sheet; using a filter needs Use AutoFilter.
  const filterRefused =
    operation === 'filter' &&
    id !== 'sheetViewSelect' &&
    !sheetPermits(state, sheet, FILTER_DEFINITION_COMMANDS.has(id) ? undefined : 'autoFilter');
  const refused =
    refusedOnProtectedSheet(state, sheet, id) ||
    filterRefused ||
    (option !== undefined && !sheetPermits(state, sheet, option));
  return refused
    ? { allowed: false, code: 'protected', reason: 'sheet is protected' }
    : { allowed: true };
};

const decideBuiltIn = (
  store: SpreadsheetStore,
  id: string,
  origin: InteractionOrigin,
): PermissionDecision => {
  const protection = builtInProtectionDecision(store, id);
  if (!protection.allowed) return protection;
  const controller = interactionControllerFor(store);
  if (!controller?.restricted) return { allowed: true };
  // Ids arrive from DOM datasets and public command APIs; validate at the boundary.
  if (!isBuiltInCommandRequest(id)) return invalidCommand(id);
  if (id === 'copy') return controller.canCopy();
  if (isBuiltInCommandId(id) && SELECTING_NAVIGATION_COMMANDS.has(id))
    return controller.canSelect();
  const intent = builtInCommandIntent(store, id, origin);
  if (intent === undefined) return invalidCommand(id);
  if (intent === null) return { allowed: true };
  return controller.canExecute(intent);
};

/** Decides whether a built-in UI command may run under the host's interaction policy. Without a
 * restricting policy every command is allowed; under one, a `null` command is allowed, a
 * selection-moving command follows the selection permission, and every other command is decided
 * by its operation and effects. An id outside {@link BuiltInCommandRequest} is denied as
 * `'invalid'`. */
export function canExecuteBuiltIn(
  store: SpreadsheetStore,
  id: BuiltInCommandRequest,
  origin: InteractionOrigin,
): PermissionDecision {
  return decideBuiltIn(store, id, origin);
}

/** {@link canExecuteBuiltIn} for an id that arrives as a plain string (a DOM dataset value or a
 * public command API argument). A string that is not a built-in command is denied as `'invalid'`
 * under a restricting policy and allowed without one, exactly as the typed form treats it. */
export function canExecuteBuiltInRequest(
  store: SpreadsheetStore,
  id: string,
  origin: InteractionOrigin,
): PermissionDecision {
  return decideBuiltIn(store, id, origin);
}
