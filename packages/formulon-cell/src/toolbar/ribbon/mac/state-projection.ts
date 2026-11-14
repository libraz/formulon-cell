import { canExecuteBuiltIn } from '../../../commands/built-in-command-policy.js';
import { selectedSessionChart } from '../../../commands/chart-design.js';
import { commentAt } from '../../../commands/comment.js';
import { workbookHasNotes } from '../../../commands/convert-notes.js';
import { cellValueIsFormulaError } from '../../../commands/error-indicators.js';
import { activeFilterRange } from '../../../commands/filter-ranges.js';
import { hasOutlineDetail } from '../../../commands/outline.js';
import { isWorkbookStructureProtected, noteChangePermitted } from '../../../commands/protection.js';
import { arrangeStackPosition } from '../../../commands/session-illustration.js';
import { TEMPORARY_SHEET_VIEW_ID } from '../../../commands/sheet-views.js';
import { activeTableOverlay } from '../../../commands/table-design.js';
import { addrKey, parseAddrKey } from '../../../engine/address.js';
import { findPivotTableAtCell } from '../../../engine/passthrough-sync.js';
import { getMacInk } from '../../../interact/mac-ink.js';
import type { SpreadsheetInstance } from '../../../mount/types.js';
import { rangeValidationsIntersect } from '../../../store/range-validations.js';
import { projectDisabledState } from '../../menu-a11y.js';
import { indexRibbonButtons, projectActiveState } from '../../ribbon-active-state.js';
import { projectPressedState } from '../button.js';
import { EMBEDDED_UNSUPPORTED, macRibbonLabelForCommand } from './labels.js';
import { MAC_CONTEXT_LABELS } from './labels-data.js';
import { toolbarLangForLocale } from './locale.js';

const MAC_ACTIVE_STATE_KEYS: Readonly<Record<string, keyof ReturnType<typeof projectActiveState>>> =
  {
    'mac.page.showGridlines': 'gridlinesVisible',
    'mac.page.printGridlines': 'printGridlines',
    'mac.page.showHeadings': 'headingsVisible',
    'mac.page.printHeadings': 'printHeadings',
    'mac.view.standard': 'workbookView',
    'mac.view.pageBreakPreview': 'workbookView',
    'mac.view.pageLayout': 'workbookView',
    'mac.view.gridlines': 'gridlinesVisible',
    'mac.view.headings': 'headingsVisible',
    'mac.formulas.showFormulas': 'formulasVisible',
    'mac.review.showNotes': 'allNotesVisible',
    'mac.data.filter': 'filterOn',
    'mac.review.protectSheet': 'protected',
    'mac.view.freeze': 'frozen',
    'mac.view.split': 'splitPanes',
  };

const MAC_POLICY_PROJECTED_COMMANDS = [
  'mac.page.theme',
  'mac.page.color',
  'mac.page.font',
  'mac.page.theme.office',
  'mac.page.theme.gray',
  'mac.page.color.office',
  'mac.page.color.gray',
  'mac.page.font.office',
  'mac.page.font.gray',
  'mac.insert.newComment',
  'mac.insert.shapeArrow',
  'mac.review.newComment',
  'mac.review.deleteComment',
  'mac.review.previousComment',
  'mac.review.nextComment',
  'mac.review.showComments',
  'mac.review.previousNote',
  'mac.review.nextNote',
  'mac.review.showHideNote',
  'mac.review.convertNotes',
  'mac.data.sortAsc',
  'mac.data.sortDesc',
  'mac.data.sortCustom',
  'mac.data.filter',
  'mac.data.clear',
  'mac.data.reapply',
  'mac.data.advancedFilter',
  'mac.formulas.removeArrows.all',
  'mac.formulas.removeArrows.precedents',
  'mac.formulas.removeArrows.dependents',
  'mac.formulas.errorCheck.run',
  'mac.formulas.errorCheck.trace',
  'mac.data.dataTable',
  'mac.view.split',
  'mac.data.validation.settings',
  'mac.data.validation.circleInvalid',
  'mac.data.validation.clearCircles',
  'mac.data.validation.clearRules',
] as const;

/** Commands whose availability follows the document alone, not the interaction policy. */
const MAC_CONTENT_STATE_COMMANDS = [
  'mac.data.showDetail',
  'mac.data.hideDetail',
  'mac.data.workbookLinks',
  'mac.insert.slicer',
  'mac.formulas.useInFormula',
  'mac.view.sheetViewKeep',
  'mac.view.sheetViewExit',
] as const;

const contentDisabledReason = (
  instance: SpreadsheetInstance,
  id: (typeof MAC_CONTENT_STATE_COMMANDS)[number],
): string | null => {
  const lang = toolbarLangForLocale(instance.i18n.locale);
  if (id === 'mac.data.workbookLinks') {
    return instance.workbook.getExternalLinks().length === 0
      ? MAC_CONTEXT_LABELS.externalLinksRequired[lang]
      : null;
  }
  if (id === 'mac.view.sheetViewKeep' || id === 'mac.view.sheetViewExit') {
    const shown = instance.store.getState().sheetViews.activeViewId;
    if (id === 'mac.view.sheetViewKeep') {
      return shown === TEMPORARY_SHEET_VIEW_ID
        ? null
        : MAC_CONTEXT_LABELS.temporarySheetViewRequired[lang];
    }
    return shown === null ? MAC_CONTEXT_LABELS.sheetViewRequired[lang] : null;
  }
  if (id === 'mac.insert.slicer') {
    const state = instance.store.getState();
    const onTable = activeTableOverlay(state) != null;
    const onPivot = findPivotTableAtCell(instance.workbook, state.selection.active);
    return onTable || onPivot ? null : MAC_CONTEXT_LABELS.slicerRequired[lang];
  }
  if (id === 'mac.formulas.useInFormula') {
    return instance.workbook.definedNames().next().done === true
      ? MAC_CONTEXT_LABELS.namesRequired[lang]
      : null;
  }
  return hasOutlineDetail(instance.store, instance.store.getState().selection.range)
    ? null
    : MAC_CONTEXT_LABELS.outlineRequired[lang];
};

const hasValuesInSelection = (instance: SpreadsheetInstance): boolean => {
  const state = instance.store.getState();
  const { sheet, r0, r1, c0, c1 } = state.selection.range;
  for (const [key, cell] of state.data.cells) {
    if (cell.value.kind === 'blank') continue;
    const addr = parseAddrKey(key);
    if (
      addr &&
      addr.sheet === sheet &&
      addr.row >= r0 &&
      addr.row <= r1 &&
      addr.col >= c0 &&
      addr.col <= c1
    ) {
      return true;
    }
  }
  return false;
};

const hasValidationInSelection = (instance: SpreadsheetInstance): boolean => {
  const { format, selection } = instance.store.getState();
  const { sheet, r0, r1, c0, c1 } = selection.range;
  const area = (r1 - r0 + 1) * (c1 - c0 + 1);
  if (rangeValidationsIntersect(format, selection.range)) return true;
  if (area <= format.formats.size) {
    for (let row = r0; row <= r1; row += 1) {
      for (let col = c0; col <= c1; col += 1) {
        if (format.formats.get(addrKey({ sheet, row, col }))?.validation) return true;
      }
    }
    return false;
  }
  for (const [key, cellFormat] of format.formats) {
    if (!cellFormat.validation) continue;
    const addr = parseAddrKey(key);
    if (
      addr &&
      addr.sheet === sheet &&
      addr.row >= r0 &&
      addr.row <= r1 &&
      addr.col >= c0 &&
      addr.col <= c1
    ) {
      return true;
    }
  }
  return false;
};

const contextualDisabledReason = (instance: SpreadsheetInstance, id: string): string | null => {
  const state = instance.store.getState();
  const strings = instance.i18n.strings.ribbonMenu;
  const lang = toolbarLangForLocale(instance.i18n.locale);
  if (/^mac\.page\.(theme|color|font)(\.|$)/.test(id)) {
    if (!instance.workbook?.capabilities?.workbookTheme)
      return MAC_CONTEXT_LABELS.themeEngineUnavailable[lang];
    if (state.protection.protectedSheets.size > 0) return MAC_CONTEXT_LABELS.themeProtected[lang];
    if (instance.workbook.getTheme()?.source === 2) return MAC_CONTEXT_LABELS.themeUnreadable[lang];
  }
  // The file stores no line end, so an engine that saves shapes cannot keep an arrow.
  if (id === 'mac.insert.shapeArrow' && instance.workbook?.capabilities?.shapes) {
    return strings.shapeArrowNotSaved;
  }
  if (
    id === 'mac.insert.newComment' ||
    (id.startsWith('mac.review.') &&
      ['newComment', 'deleteComment', 'previousComment', 'nextComment', 'showComments'].includes(
        id.slice('mac.review.'.length),
      ))
  ) {
    if (!instance.workbook?.capabilities?.threadedComments)
      return MAC_CONTEXT_LABELS.threadedCommentsUnavailable[lang];
    const roots =
      state.threadedComments.bySheet
        .get(state.data.sheetIndex)
        ?.filter((entry) => !entry.parentId) ?? [];
    if (
      id === 'mac.review.deleteComment' &&
      !roots.some(
        (entry) =>
          entry.cell.row === state.selection.active.row &&
          entry.cell.col === state.selection.active.col,
      )
    )
      return strings.commentDeleteRequiresActive;
    // Show Comments stays available on an empty sheet; it opens an empty list.
    if (['mac.review.previousComment', 'mac.review.nextComment'].includes(id) && !roots.length)
      return strings.commentNone;
  }
  // Note navigation needs a note anywhere in the book; Show/Hide needs one on the active cell.
  if (
    id === 'mac.review.previousNote' ||
    id === 'mac.review.nextNote' ||
    id === 'mac.review.showHideNote'
  ) {
    const formats = [...state.format.formats.values()];
    if (!formats.some((format) => typeof format.comment === 'string' && format.comment.length > 0))
      return strings.commentNone;
    if (id === 'mac.review.showHideNote' && commentAt(state, state.selection.active) === null)
      return strings.commentDeleteRequiresActive;
  }
  // Convert to Comments needs a thread-capable engine and a note on some sheet. Provisional.
  if (id === 'mac.review.convertNotes') {
    if (!instance.workbook?.capabilities?.threadedComments)
      return MAC_CONTEXT_LABELS.threadedCommentsUnavailable[lang];
    if (instance.workbook && !workbookHasNotes(instance.workbook)) return strings.commentNone;
  }
  if (id === 'mac.view.split' && !instance.workbook?.capabilities?.sheetSplit) {
    return strings.splitNotSaved;
  }
  if (id === 'mac.data.dataTable' && !instance.workbook?.capabilities?.dataTables) {
    return EMBEDDED_UNSUPPORTED[lang];
  }
  if (id === 'mac.data.clear' && !activeFilterRange(state)) {
    return MAC_CONTEXT_LABELS.filterRangeRequired[lang];
  }
  if (id === 'mac.data.reapply' && state.ui.filterCriteria.length === 0) {
    return MAC_CONTEXT_LABELS.filterCriteriaRequired[lang];
  }

  if (id.startsWith('mac.formulas.removeArrows.')) {
    const traces = state.traces.items;
    const hasPrecedents = traces.some((trace) => trace.kind === 'precedent');
    const hasDependents = traces.some((trace) => trace.kind === 'dependent');
    if (id === 'mac.formulas.removeArrows.all' && !hasPrecedents && !hasDependents) {
      return strings.removeArrowsRequiresAny;
    }
    if (id === 'mac.formulas.removeArrows.precedents' && !hasPrecedents) {
      return strings.removePrecedentArrowsRequiresAny;
    }
    if (id === 'mac.formulas.removeArrows.dependents' && !hasDependents) {
      return strings.removeDependentArrowsRequiresAny;
    }
  }

  if (id === 'mac.formulas.errorCheck.trace') {
    const active = state.selection.active;
    const cell = state.data.cells.get(addrKey(active));
    if (!cell?.formula || !cellValueIsFormulaError(cell.value)) {
      return strings.traceErrorRequiresFormulaError;
    }
  }

  if (id === 'mac.data.validation.circleInvalid' || id === 'mac.data.validation.clearRules') {
    if (!hasValidationInSelection(instance)) return strings.validationRequiresRules;
  }
  if (
    id === 'mac.data.validation.clearCircles' &&
    state.errorIndicators.validationCircles.size === 0
  ) {
    return strings.validationClearCirclesRequiresAny;
  }
  return null;
};

/** Projects Mac alias buttons onto the same store-derived state as the
 *  generic ribbon. Hosts call this after an instance mutation; it is kept
 *  separate from the render pass so active controls do not require a full
 *  menu rebuild. */
export const projectMacRibbonState = (
  host: HTMLElement,
  instance: SpreadsheetInstance | null,
): void => {
  if (!instance) return;
  const lang = toolbarLangForLocale(instance.i18n.locale);
  const active = projectActiveState(instance);
  const ink = getMacInk(instance);
  const inkTool = ink?.getTool();
  const buttons = indexRibbonButtons(host);
  const drawingStates: Readonly<Record<string, boolean>> = {
    'mac.draw.toggle': ink?.isActive() ?? false,
    'mac.draw.eraser': inkTool === 'eraser',
    'mac.draw.penBlack': inkTool === 'pen-black',
    'mac.draw.penRed': inkTool === 'pen-red',
    'mac.draw.pencil': inkTool === 'pencil',
    'mac.draw.highlighter': inkTool === 'highlighter',
    'mac.draw.trackpad': ink?.isActive() === true && ink.getTrackpadMode(),
  };
  // Toggles that follow a pane or bar being shown, plus the fixed ones.
  const toggleStates: Readonly<Record<string, boolean>> = {
    'mac.formulas.watch': instance.store.getState().ui.watchPanelOpen,
    objectSelectionPane: instance.store.getState().ui.workbookObjectsOpen,
    'mac.view.formulaBar': instance.store.getState().ui.formulaBarVisible,
    // The page-layout rulers are always on in this host; Zoom to Selection is a
    // toggle the desktop never leaves pressed.
    'mac.view.ruler': true,
    'mac.view.zoomSelection': false,
    // Lasso Select is a toggle the host cannot turn on.
    'mac.draw.lasso': false,
    // Show Changes and Always Open Read-Only have no state here, and Show Comments opens a
    // list rather than holding a view on.
    'mac.review.showChanges': false,
    'mac.review.showComments': false,
    'mac.review.alwaysReadOnly': false,
    'mac.review.protectWorkbook': isWorkbookStructureProtected(instance.store.getState()),
  };
  for (const [command, pressed] of Object.entries({ ...drawingStates, ...toggleStates })) {
    for (const button of buttons.get(command) ?? []) {
      projectPressedState(button, pressed);
    }
  }

  for (const [command, key] of Object.entries(MAC_ACTIVE_STATE_KEYS)) {
    let pressed = Boolean(active[key]);
    if (
      command === 'mac.view.standard' ||
      command === 'mac.view.pageLayout' ||
      command === 'mac.view.pageBreakPreview'
    ) {
      const view = command.endsWith('standard')
        ? 'normal'
        : command.endsWith('pageLayout')
          ? 'pageLayout'
          : 'pageBreakPreview';
      pressed = active.workbookView === view;
    }
    for (const button of buttons.get(command) ?? []) {
      projectPressedState(button, pressed);
    }
  }

  // The Freeze menu's first item doubles as Unfreeze while panes are frozen.
  for (const button of buttons.get('mac.view.freeze.panes') ?? []) {
    button.dataset.freezeLabel ??= button.textContent ?? '';
    const label = active.frozen
      ? macRibbonLabelForCommand('mac.view.freeze.off', lang)
      : button.dataset.freezeLabel;
    button.textContent = label;
    button.title = label;
    button.setAttribute('aria-label', label);
  }

  // The Notes menu's first item edits the active cell's note, and is absent when that note is
  // locked against editing.
  const noteState = instance.store.getState();
  const activeNote = commentAt(noteState, noteState.selection.active) !== null;
  for (const button of buttons.get('mac.review.newNote') ?? []) {
    if (!button.closest('.fc-tb__menu')) continue;
    button.dataset.newNoteLabel ??= button.textContent ?? '';
    const label = activeNote
      ? macRibbonLabelForCommand('mac.review.editNote', lang)
      : button.dataset.newNoteLabel;
    button.textContent = label;
    button.title = label;
    button.setAttribute('aria-label', label);
    button.hidden = activeNote && !noteChangePermitted(noteState, noteState.selection.active);
  }

  // Recommended Charts offers its action only once the selection holds data.
  const hasData = hasValuesInSelection(instance);
  for (const [id, shown] of [
    ['mac.insert.recommendedChart.none', !hasData],
    ['mac.insert.recommendedChart.apply', hasData],
  ] as const) {
    for (const button of buttons.get(id) ?? []) button.hidden = !shown;
  }

  // A chart moves only among the sheet's charts: nothing in front or behind it
  // leaves the matching command off, as the desktop does for a lone chart.
  const chartState = instance.store.getState();
  const chart = selectedSessionChart(chartState);
  if (chart) {
    const charts = chartState.charts.charts.filter((item) => item.sheet === chart.sheet);
    const { atFront, atBack } = arrangeStackPosition(charts, chart.id);
    const strings = instance.i18n.strings.ribbonMenu;
    for (const [id, blocked, reason] of [
      ['objectBringForward', atFront, strings.arrangeAtFront],
      ['objectSendBackward', atBack, strings.arrangeAtBack],
    ] as const) {
      for (const button of buttons.get(id) ?? []) {
        projectDisabledState(button, blocked, reason, {
          datasetKey: 'ribbonDisabledReason',
          titlePrefix:
            button.getAttribute('aria-label')?.trim() || macRibbonLabelForCommand(id, lang),
        });
      }
    }
  }

  for (const id of MAC_CONTENT_STATE_COMMANDS) {
    const reason = contentDisabledReason(instance, id);
    for (const button of buttons.get(id) ?? []) {
      projectDisabledState(button, reason !== null, reason, {
        datasetKey: 'ribbonDisabledReason',
        titlePrefix:
          button.getAttribute('aria-label')?.trim() || macRibbonLabelForCommand(id, lang),
      });
    }
  }

  for (const id of MAC_POLICY_PROJECTED_COMMANDS) {
    const decision = canExecuteBuiltIn(instance.store, id, 'ribbon');
    const contextualReason = contextualDisabledReason(instance, id);
    const disabled = !decision.allowed || contextualReason !== null;
    const reason = decision.allowed
      ? contextualReason
      : instance.i18n.strings.backstage.commandUnavailable;
    for (const button of buttons.get(id) ?? []) {
      const titlePrefix =
        button.getAttribute('aria-label')?.trim() || macRibbonLabelForCommand(id, lang);
      projectDisabledState(button, disabled, reason, {
        datasetKey: 'ribbonDisabledReason',
        titlePrefix,
      });
    }
  }
};
