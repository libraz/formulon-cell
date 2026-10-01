import type { SpreadsheetStore } from '../store/store.js';
import { interactionControllerFor } from './interaction-controller.js';
import type {
  InteractionOperation,
  InteractionOrigin,
  OperationIntent,
  PermissionDecision,
} from './interaction-policy.js';

/** Stable built-in command classification. Keep this explicit: an unknown
 * command must not become executable merely because a host allowed a broad
 * operation category. */
export const BUILT_IN_COMMAND_OPERATION: Readonly<Record<string, InteractionOperation>> = {
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
  alignLeft: 'format',
  alignCenter: 'format',
  alignRight: 'format',
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
  editPhonetic: 'format',
  protect: 'protection',
  protectReview: 'protection',
  pageSetup: 'pageSetup',
  print: 'print',
  export: 'export',
  printArea: 'pageSetup',
  pageBreaks: 'pageSetup',
  dataValidation: 'validation',
  namedRanges: 'namedRange',
  formatTableInsert: 'table',
  formatTableHome: 'table',
  pivotTableInsert: 'table',
  chartInsert: 'object',
  pictureInsert: 'object',
  shapesInsert: 'object',
  screenshotInsert: 'object',
  workbookObjectsView: 'object',
  sheetInsert: 'sheet',
  sheetDelete: 'sheet',
  sheetRename: 'sheet',
  undo: 'valueEdit',
  redo: 'valueEdit',
};

const SAFE_NAVIGATION_COMMANDS: ReadonlySet<string> = new Set([
  'copy',
  'selectAll',
  'openHyperlink',
  'toggleWatch',
  'findHome',
  'gotoSpecial',
  'gotoSpecialHome',
  'viewNormal',
  'viewPageLayout',
  'viewPageBreakPreview',
  'zoom75',
  'zoom100',
  'zoom125',
  'zoomDialog',
  'showFormulas',
  'r1c1',
]);

const unsupported = (id: string): PermissionDecision => ({
  allowed: false,
  code: 'unsupported',
  reason: `built-in command ${id} is not classified for restricted execution`,
});

export function canExecuteBuiltIn(
  store: SpreadsheetStore,
  id: string,
  origin: InteractionOrigin,
): PermissionDecision {
  const controller = interactionControllerFor(store);
  if (!controller?.restricted) return { allowed: true };
  if (id === 'copy') return controller.canCopy();
  if (id === 'selectAll') return controller.canSelect();
  if (SAFE_NAVIGATION_COMMANDS.has(id)) return { allowed: true };
  const operation = BUILT_IN_COMMAND_OPERATION[id];
  if (!operation) return unsupported(id);
  const state = store.getState();
  const intent: OperationIntent = {
    operation,
    origin,
    commandId: id,
    effects:
      operation === 'print' || operation === 'export'
        ? [{ kind: 'workbook' }]
        : [{ kind: 'range', range: state.selection.range }],
  };
  return controller.canExecute(intent);
}
