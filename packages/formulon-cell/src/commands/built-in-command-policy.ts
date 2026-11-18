import { addrKey, MAX_COL, MAX_ROW } from '../engine/address.js';
import { formatHistoryBridgeFor } from '../engine/format-history-bridge.js';
import { findPivotTableAtCell } from '../engine/passthrough-sync.js';
import { sheetHasCells } from '../engine/sheet-kind.js';
import type { Range } from '../engine/types.js';
import {
  isWholeColumnRange,
  isWholeRowRange,
  normalizeRange,
  selectionRangeList,
} from '../store/selection-geometry.js';
import type { SpreadsheetStore } from '../store/store.js';
import { type BuiltInCommandId, isBuiltInCommandId } from './built-in-command-ids.js';
import { isMacChrome } from './chrome-platform.js';
import { insertCopiedBandIntents, insertCopiedBlockIntents } from './clipboard/paste-intents.js';
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
import { needsCells, valuesOnlyClipboard } from './interaction-policy.js';
import { outlineDetailAxis } from './outline.js';
import { deny } from './permission-reasons.js';
import { pivotReportRange } from './pivot-actions.js';
import {
  isCellWritable,
  isPartialShift,
  noteChangePermitted,
  protectionOption,
  rangeCellsWritable,
  refusedOnProtectedSheet,
  sheetPermits,
  tableChangePermitted,
  unlockedCellsOption,
  workbookStructurePermits,
} from './protection.js';
import { FUNCTION_SIGNATURES } from './refs.js';
import { selectionOverlapsTable } from './table-actions.js';
import { asThreadIntent } from './thread-intent.js';

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
  pasteFormatsOnly: 'paste',
  pasteFormulas: 'paste',
  pasteFormulasNumFmt: 'paste',
  cut: 'paste',
  clear: 'clear',
  insertCells: 'insertRows',
  // Decided on the intents its write declares (see insertCopiedDecision); 'paste' alone without
  // a copy of this workbook to size the insertion.
  insertCopiedCells: 'paste',
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
  newComment: 'comment',
  // Pick From Drop-down List edits the active cell.
  pickFromList: 'valueEdit',
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
  tableQuickStyles: 'table',
  tableSummarizePivot: 'pivot',
  tableInsertSlicer: 'object',
  pivotInsertSlicer: 'object',
  pivotTableInsert: 'pivot',
  tableName: 'table',
  // Refresh All runs data connections, not a table change: enabled on a protected sheet too.
  tableRefresh: null,
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
  // The Sparkline tab: settings and membership of sparkline groups are drawing-object changes;
  // Clear decides as clearing, by each cell's lock. The Axis menu and the Mac Axis button edit
  // nothing themselves and stay enabled: on a protected sheet the menu's controls are inert and
  // the button does nothing.
  sparklineAxis: null,
  sparklineAxisDialog: null,
  sparklineClear: 'clear',
  sparklineColor: 'object',
  sparklineEditData: 'object',
  sparklineGroup: 'object',
  sparklineMarkerColor: 'object',
  sparklineShowFirst: 'object',
  sparklineShowHigh: 'object',
  sparklineShowLast: 'object',
  sparklineShowLow: 'object',
  sparklineShowMarkers: 'object',
  sparklineShowNegative: 'object',
  sparklineStyleGallery: 'object',
  sparklineTypeColumn: 'object',
  sparklineTypeLine: 'object',
  sparklineTypeWinLoss: 'object',
  sparklineUngroup: 'object',
  shapeOutline: 'object',
  shapeRotate: 'object',
  shapeChangeShape: 'object',
  scaleWidth: 'pageSetup',
  scaleHeight: 'pageSetup',
  sheetViewSelect: 'filter',

  // Excel 365 Mac ribbon actions. Keep these aliases explicit so a restricted
  // embed cannot gain a new mutation merely by adopting the Mac profile.
  'mac.insert.pivotTable': 'pivot',
  'mac.insert.recommendedPivotTable': 'pivot',
  'mac.insert.table': 'table',
  'mac.insert.photo': 'object',
  'mac.insert.photo.file': 'object',
  'mac.insert.shapeRectangle': 'object',
  'mac.insert.shapeArrow': 'object',
  'mac.insert.shapeLine': 'object',
  'mac.insert.shapeRoundedRectangle': 'object',
  'mac.insert.shapeOval': 'object',
  'mac.insert.shapeTriangle': 'object',
  'mac.insert.shapeDiamond': 'object',
  'mac.insert.screenshot': 'object',
  'mac.insert.screenshot.clipping': 'object',
  'mac.insert.textBox': 'object',
  'mac.insert.textBox.horizontal': 'object',
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
  'mac.insert.sparkline.line': 'object',
  'mac.insert.sparkline.column': 'object',
  'mac.insert.sparkline.winLoss': 'object',
  'mac.insert.slicer': 'object',
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
  'mac.draw.eraser.stroke': null,
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
  'mac.formulas.defineName.define': 'namedRange',
  'mac.formulas.useInFormula': 'formulaEdit',
  'mac.formulas.useInFormula.paste': 'formulaEdit',
  'mac.formulas.createFromSelection': 'namedRange',
  'mac.formulas.removeArrows.all': null,
  'mac.formulas.removeArrows.precedents': null,
  'mac.formulas.removeArrows.dependents': null,
  'mac.formulas.errorCheck.trace': null,
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
  'mac.formulas.calc.auto': null,
  'mac.formulas.calc.autoNoTable': null,
  'mac.formulas.calc.manual': null,
  'mac.formulas.calc.formatStale': null,
  'mac.formulas.calc.compatVersion.1': null,
  'mac.formulas.calc.compatVersion.2': null,
  'mac.formulas.calc.compatVersion.3': null,
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
  'mac.review.convertNotes': 'comment',
  'mac.review.protectSheet': 'protection',
  'mac.review.protectWorkbook': 'protection',
  'mac.view.sheetViewKeep': 'sheet',
  'mac.view.sheetViewExit': 'sheet',
  'mac.view.sheetViewNew': 'sheet',
  'mac.view.sheetViewOptions': 'sheet',
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
  cellStyles: 'format',
  conditional: 'format',
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
  'mac.automate.showScripts.samples': null,
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
  'mac.review.accessibility.check': null,
  'mac.review.accessibility.formatTable': 'table',
  'mac.review.accessibility.unmerge': 'merge',
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
  'mac.view.cellFocus': null,
  'mac.view.cellFocus.autoHighlight': null,
  'mac.view.cellFocus.focusCell': null,
  'mac.view.formulaBar': null,
  'mac.view.freeze': null,
  'mac.view.freeze.firstColumn': null,
  'mac.view.freeze.firstRow': null,
  'mac.view.freeze.panes': null,
  'mac.view.gridlines': null,
  'mac.view.headings': null,
  'mac.view.pageBreakPreview': null,
  'mac.view.pageLayout': null,
  'mac.view.split': null,
  'mac.view.standard': null,
  'mac.view.zoom': null,
  'mac.view.zoom100': null,
  'mac.view.zoomSelection': null,
  nextCommentReview: null,
  openHyperlink: null,
  outlineSymbols: null,
  pageLayoutGridlinesView: null,
  pageLayoutHeadingsView: null,
  pageTheme: null,
  pivotBlankRows: 'pivot',
  pivotClear: 'pivot',
  pivotFieldsItemsSets: 'pivot',
  pivotGrandTotals: 'pivot',
  pivotReportLayout: 'pivot',
  pivotSelect: null,
  pivotSelectData: null,
  pivotSelectEnable: null,
  pivotSelectLabels: null,
  pivotSelectLabelsAndData: null,
  pivotSelectTable: null,
  pivotSubtotals: 'pivot',
  notesReview: null,
  previousCommentReview: null,
  recalcNow: null,
  recalcSheet: null,
  recordActions: null,
  script: null,
  selectAll: null,
  selectionPanePageLayout: null,
  showCommentsReview: null,
  showFormulasFormula: null,
  showHideNote: null,
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
  // The Home Clear control's face: decided as clearing, by the selection, not by Format cells.
  clearFormat: 'clear',
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
  pasteLink: 'paste',
  pasteNoBorders: 'paste',
  pdf: 'export',
  pivotActiveField: 'pivot',
  pivotButtons: 'pivot',
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
  pivotGroupSelection: 'pivot',
  pivotHeaders: 'pivot',
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
  pivotQuickStyles: 'pivot',
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
  'namedRange',
  'sheet',
  'protection',
  'print',
  'export',
]);

/** Row and column layout and drawing-object operations: decided for the sheet, by its Format
 *  rows/columns or Edit objects option. */
const SHEET_LAYOUT_OPERATIONS: ReadonlySet<InteractionOperation> = new Set([
  'resizeRows',
  'resizeColumns',
  'object',
]);

/** Show/Hide Detail: expanding or collapsing an outline is a row or a column layout change,
 *  by the axis of the group around the selection. */
const OUTLINE_DETAIL_COMMANDS: ReadonlySet<BuiltInCommandId> = new Set<BuiltInCommandId>([
  'outlineShowDetail',
  'outlineHideDetail',
  'mac.data.showDetail',
  'mac.data.hideDetail',
]);

/** The operation of built-in `id` in the current state: its classification, with Show/Hide
 *  Detail taking the axis of the outline it acts on. */
const builtInOperation = (
  store: SpreadsheetStore,
  id: BuiltInCommandId,
): InteractionOperation | null => {
  // Insert / Delete Cells over whole columns inserts or deletes the columns.
  if (id === 'insertCells' || id === 'deleteCells') {
    const columns = isWholeColumnRange(normalized(store.getState().selection.range));
    if (id === 'insertCells') return columns ? 'insertColumns' : 'insertRows';
    return columns ? 'deleteColumns' : 'deleteRows';
  }
  if (!OUTLINE_DETAIL_COMMANDS.has(id)) return BUILT_IN_COMMAND_OPERATION[id];
  const { selection } = store.getState();
  return outlineDetailAxis(store, selection.range) === 'col' ? 'resizeColumns' : 'resizeRows';
};

/** Clear Contents and the Home Clear control, which a locked selected cell greys as a whole. */
const CLEAR_CONTROL_COMMANDS: ReadonlySet<string> = new Set(['clear', 'clearFormat']);

/** Commands that change threaded comments, which sheet protection never decides. */
const THREAD_COMMANDS: ReadonlySet<string> = new Set([
  'newComment',
  'mac.insert.newComment',
  'mac.review.newComment',
  'mac.review.deleteComment',
]);

/** Commands that start a thread at the active cell, which a cell holding a note cannot take. */
const THREAD_ADD_COMMANDS: ReadonlySet<string> = new Set([
  'newComment',
  'mac.insert.newComment',
  'mac.review.newComment',
]);

const activeHasNote = (state: ReturnType<SpreadsheetStore['getState']>): boolean =>
  state.format.formats.get(addrKey(state.selection.active))?.comment !== undefined;

/** Whether protection refuses the note command `id`: a note goes by Edit objects. Review > Delete
 *  also deletes the active cell's threads, so it is refused only over a note. */
const noteCommandRefused = (
  state: ReturnType<SpreadsheetStore['getState']>,
  id: string,
  operation: InteractionOperation | null,
): boolean =>
  operation === 'comment' &&
  !THREAD_COMMANDS.has(id) &&
  (id !== 'deleteCommentReview' || activeHasNote(state)) &&
  !noteChangePermitted(state, state.selection.active);

/** Commands that turn AutoFilter on or off rather than use an existing filter. */
const FILTER_DEFINITION_COMMANDS: ReadonlySet<BuiltInCommandId> = new Set<BuiltInCommandId>([
  'filter',
  'mac.data.filter',
  'tableFilterButton',
]);

/** The table the active cell is in, or null when there is none. */
const activeTableRange = (store: SpreadsheetStore): Range | null => {
  const { selection, tables } = store.getState();
  const a = selection.active;
  return (
    tables.tables.find(
      (t) =>
        t.range.sheet === a.sheet &&
        a.row >= t.range.r0 &&
        a.row <= t.range.r1 &&
        a.col >= t.range.c0 &&
        a.col <= t.range.c1,
    )?.range ?? null
  );
};

/** Commands that create a table, or a Data Table, over the selection: refused where it overlaps an
 *  existing table. */
const TABLE_CREATION_COMMANDS: ReadonlySet<BuiltInCommandId> = new Set<BuiltInCommandId>([
  'formatTableInsert',
  'mac.insert.table',
  'mac.automate.makeSubtable',
  'mac.data.dataTable',
]);

/** The pivot report the active cell is in, or null when there is none. */
const activeReportRange = (store: SpreadsheetStore): Range | null => {
  const wb = formatHistoryBridgeFor(store)?.workbook;
  if (!wb?.capabilities.pivotTables) return null;
  const pivot = findPivotTableAtCell(wb, store.getState().selection.active);
  if (!pivot) return null;
  return pivotReportRange(wb, pivot.sheetIndex, pivot.pivotIndex);
};

/** Prefix of the Mac ribbon's per-function commands (`mac.function.SUM`). */
export const FUNCTION_COMMAND_PREFIX = 'mac.function.';

/** The Mac ribbon command that inserts function `name`. */
export const macFunctionCommandId = (name: string): `mac.function.${string}` =>
  `${FUNCTION_COMMAND_PREFIX}${name}`;

/** The function a Mac per-function command inserts; null for any other command. */
export const macFunctionNameOf = (id: string): string | null =>
  id.startsWith(FUNCTION_COMMAND_PREFIX) ? id.slice(FUNCTION_COMMAND_PREFIX.length) : null;

/** Commands that shift cells rather than whole rows or columns: Insert / Delete Cells over a
 *  selection that is not whole rows or columns, and a table's own rows and columns. */
const CELL_SHIFT_COMMANDS: ReadonlySet<string> = new Set([
  'insertCells',
  'deleteCells',
  'insertTableRows',
  'insertTableColumns',
]);

/** Insert / Delete Cells: over part of the rows they open a dialog of a shift, Entire row and
 *  Entire column. */
const CELL_SHIFT_DIALOG_COMMANDS: ReadonlySet<string> = new Set(['insertCells', 'deleteCells']);

const normalized = (range: Range): Range => normalizeRange(range);

/** The band an insert or delete of `id` moves: whole rows (columns) from the selection to the sheet
 *  end, or for a cell shift only the selection's columns (rows). */
const shiftedBand = (
  id: string,
  operation: InteractionOperation,
  selection: Range,
): Range | null => {
  const s = normalized(selection);
  const rows = operation === 'insertRows' || operation === 'deleteRows';
  const columns = operation === 'insertColumns' || operation === 'deleteColumns';
  if (!rows && !columns) return null;
  const whole = isWholeRowRange(s) || isWholeColumnRange(s);
  if (CELL_SHIFT_COMMANDS.has(id) && !(whole && !id.startsWith('insertTable'))) {
    return rows
      ? { sheet: s.sheet, r0: s.r0, c0: s.c0, r1: MAX_ROW, c1: s.c1 }
      : { sheet: s.sheet, r0: s.r0, c0: s.c0, r1: s.r1, c1: MAX_COL };
  }
  return rows ? rowShiftBand(s.sheet, s.r0) : columnShiftBand(s.sheet, s.c0);
};

const isBuiltInCommandRequest = (id: string): id is BuiltInCommandRequest =>
  isBuiltInCommandId(id) || macFunctionNameOf(id) !== null;

const invalidCommand = (id: string): PermissionDecision =>
  deny('invalid', 'notBuiltInCommand', { reason: `${id} is not a built-in command` });

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
    ? builtInOperation(store, id)
    : functionLeafOperation(store, macFunctionNameOf(id) ?? '');
  if (operation === undefined || operation === null) return operation;
  const workbookTheme = /^mac\.page\.(theme|color|font)\.(office|gray)$/.test(id);
  const state = store.getState();
  const selection = state.selection.range;
  const band = shiftedBand(id, operation, selection);
  const report =
    operation === 'pivot'
      ? activeReportRange(store)
      : operation === 'table' && !(isBuiltInCommandId(id) && TABLE_CREATION_COMMANDS.has(id))
        ? activeTableRange(store)
        : null;
  let effects: readonly OperationEffect[];
  if (WORKBOOK_EFFECT_OPERATIONS.has(operation) || workbookTheme || id === 'sheetViewSelect')
    effects = [{ kind: 'workbook' }];
  // Turning AutoFilter on or off is a sheet-level filter change; using a filter is decided on its range.
  else if (isBuiltInCommandId(id) && FILTER_DEFINITION_COMMANDS.has(id))
    effects = [{ kind: 'sheet', sheet: selection.sheet }];
  // A sheet-level write, or one whose Protect Sheet option alone makes it available (locked
  // cells refuse it at run time): decided for the active sheet as a whole.
  else if (
    operation === 'pageSetup' ||
    SHEET_LAYOUT_OPERATIONS.has(operation) ||
    unlockedCellsOption(operation)
  )
    effects = [{ kind: 'sheet', sheet: selection.sheet }];
  // A command on an existing report or table is decided on the one the active cell is in.
  else if (report) effects = [{ kind: 'range', range: report }];
  else if (id === 'pickFromList') effects = [{ kind: 'cells', cells: [state.selection.active] }];
  else if (band) effects = [{ kind: 'range', range: band }];
  else if (operation === 'format')
    effects = [selection, ...(state.selection.extraRanges ?? [])].map(
      (range): OperationEffect => ({ kind: 'range', range }),
    );
  else effects = [{ kind: 'range', range: selection }];
  const intent: OperationIntent = { operation, origin, commandId: id, effects };
  return THREAD_COMMANDS.has(id) ? asThreadIntent(intent) : intent;
};

/** Protection's decision for a built-in command on the active sheet, with or without a policy:
 *  the always-refused commands, an option-gated operation (Use PivotTable for pivot edits)
 *  whose option is off, a table change other than banding, and a sheet operation under
 *  workbook-structure protection. Commands it
 *  does not name are left to their operation's own protection check. */
export const builtInProtectionDecision = (
  store: SpreadsheetStore,
  id: string,
): PermissionDecision => {
  if (!isBuiltInCommandId(id)) return { allowed: true };
  const state = store.getState();
  const sheet = state.data.sheetIndex;
  const operation = builtInOperation(store, id);
  if (operation !== null && !workbookStructurePermits(state, operation))
    return deny('protected', 'workbookStructureProtected');
  if (operation !== null && needsCells(operation) && !sheetHasCells(state, sheet))
    return deny('operationDenied', 'sheetHasNoCells');
  // A cell shift is refused on a protected sheet whatever its options, so Insert / Delete Cells
  // over part of the rows never opens its dialog there; over whole rows or columns it acts at once.
  const shift =
    operation && CELL_SHIFT_COMMANDS.has(id) && !sheetPermits(state, sheet)
      ? shiftedBand(id, operation, state.selection.range)
      : null;
  const partial = shift !== null && operation !== null && isPartialShift(operation, shift);
  const option = operation
    ? (unlockedCellsOption(operation) ?? protectionOption(operation))
    : undefined;
  // Turning AutoFilter on or off is refused on a protected sheet; using a filter needs Use AutoFilter.
  const filterRefused =
    operation === 'filter' &&
    id !== 'sheetViewSelect' &&
    !sheetPermits(state, sheet, FILTER_DEFINITION_COMMANDS.has(id) ? undefined : 'autoFilter');
  const refused =
    refusedOnProtectedSheet(state, sheet, id) ||
    filterRefused ||
    (operation === 'pivot' && !sheetPermits(state, sheet, 'pivotTables')) ||
    (operation === 'object' && !sheetPermits(state, sheet, 'objects')) ||
    (operation === 'table' && !tableChangePermitted(state, sheet)) ||
    partial ||
    noteCommandRefused(state, id, operation) ||
    // Pick From Drop-down List edits the active cell, so a locked one greys it.
    (id === 'pickFromList' && !isCellWritable(state, state.selection.active)) ||
    // Clearing is all or nothing, so a locked selected cell greys Clear Contents and the whole
    // Home Clear control.
    (CLEAR_CONTROL_COMMANDS.has(id) &&
      !selectionRangeList(state.selection).every((range) =>
        rangeCellsWritable(state, normalized(range)),
      )) ||
    (option !== undefined && !sheetPermits(state, sheet, option));
  return refused ? deny('protected', 'sheetProtected') : { allowed: true };
};

const decideBuiltIn = (
  store: SpreadsheetStore,
  id: string,
  origin: InteractionOrigin,
): PermissionDecision => {
  // A cell holds a note or a thread, never both.
  if (THREAD_ADD_COMMANDS.has(id) && activeHasNote(store.getState()))
    return deny('invalid', 'cellHoldsNote');
  // The Mac desktop has no Allow Edit Ranges; stored ranges are still enforced.
  if (id === 'protectionReview' && isMacChrome(store))
    return deny('invalid', 'allowEditRangesUnavailable');
  const protection = builtInProtectionDecision(store, id);
  if (!protection.allowed) return protection;
  if (
    isBuiltInCommandId(id) &&
    TABLE_CREATION_COMMANDS.has(id) &&
    selectionOverlapsTable(store.getState())
  )
    return deny('invalid', 'selectionOverlapsTable');
  const controller = interactionControllerFor(store);
  if (!controller?.restricted) return { allowed: true };
  if (CELL_SHIFT_DIALOG_COMMANDS.has(id) && !wholeLineSelection(store)) {
    const choices = cellShiftChoices(store, id === 'insertCells' ? 'insert' : 'delete', origin);
    return choices.shift || choices.rows || choices.cols
      ? { allowed: true }
      : deny('operationDenied', 'dialogNoChoiceAllowed');
  }
  // Ids arrive from DOM datasets and public command APIs; validate at the boundary.
  if (!isBuiltInCommandRequest(id)) return invalidCommand(id);
  if (id === 'copy') return controller.canCopy();
  if (id === 'cut' && valuesOnlyClipboard(controller.policy))
    return deny('operationDenied', 'valuesOnlyClipboardCannotCut');
  if (id === 'insertCopiedCells') {
    const decision = insertCopiedDecision(store, origin);
    if (decision) return decision;
  }
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

/**
 * Insert Copied Cells decided on the intents its write declares for the copy on this workbook's
 * clipboard: a whole-row (column) copy inserts lines at the selection; any other block is allowed
 * while either shift direction of its dialog is. Null without such a copy.
 */
const insertCopiedDecision = (
  store: SpreadsheetStore,
  origin: InteractionOrigin,
): PermissionDecision | null => {
  const controller = interactionControllerFor(store);
  const state = store.getState();
  const copied = state.ui.copyRange;
  if (!controller || !copied) return null;
  const source = normalized(copied);
  const selection = normalized(state.selection.range);
  const at = { sheet: selection.sheet, row: selection.r0, col: selection.c0 };
  const cut = state.ui.copyMode === 'cut';
  const commandId = 'insertCopiedCells';
  const decide = (intents: readonly OperationIntent[]): PermissionDecision => {
    for (const intent of intents) {
      const decision = controller.canExecute(intent);
      if (!decision.allowed) return decision;
    }
    return { allowed: true };
  };
  const band = insertCopiedBandIntents(state, origin, commandId, source, cut, at);
  if (band) return decide(band);
  const size = { rows: source.r1 - source.r0 + 1, cols: source.c1 - source.c0 + 1 };
  const down = decide(
    insertCopiedBlockIntents(state, origin, commandId, size, cut ? source : null, at, 'down'),
  );
  if (down.allowed) return down;
  const right = decide(
    insertCopiedBlockIntents(state, origin, commandId, size, cut ? source : null, at, 'right'),
  );
  return right.allowed ? right : down;
};

const wholeLineSelection = (store: SpreadsheetStore): boolean => {
  const s = normalized(store.getState().selection.range);
  return isWholeRowRange(s) || isWholeColumnRange(s);
};

/** Which choices of the Insert (Delete) Cells dialog may run from `origin`: the shift of the
 *  selection's cells (refused on a protected sheet), Entire row and Entire column (by their
 *  options). The dialog opens only when one of them is allowed. */
export function cellShiftChoices(
  store: SpreadsheetStore,
  kind: 'insert' | 'delete',
  origin: InteractionOrigin,
): { shift: boolean; rows: boolean; cols: boolean } {
  const state = store.getState();
  const selection = normalized(state.selection.range);
  const operation = kind === 'insert' ? 'insertRows' : 'deleteRows';
  const band = shiftedBand(kind === 'insert' ? 'insertCells' : 'deleteCells', operation, selection);
  const controller = interactionControllerFor(store);
  const shift =
    sheetPermits(state, selection.sheet) &&
    (!controller?.restricted ||
      (band !== null &&
        controller.canExecute({
          operation,
          origin,
          commandId: kind === 'insert' ? 'insertCells' : 'deleteCells',
          effects: [{ kind: 'range', range: band }],
        }).allowed));
  return {
    shift,
    rows: decideBuiltIn(store, kind === 'insert' ? 'rowInsertAbove' : 'rowDelete', origin).allowed,
    cols: decideBuiltIn(store, kind === 'insert' ? 'colInsertLeft' : 'colDelete', origin).allowed,
  };
}
