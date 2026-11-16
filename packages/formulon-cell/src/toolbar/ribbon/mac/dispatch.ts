import type { BuiltInCommandId } from '../../../commands/built-in-command-ids.js';
import {
  canExecuteBuiltIn,
  canExecuteBuiltInRequest,
  macFunctionNameOf,
} from '../../../commands/built-in-command-policy.js';
import { inputLocaleForWorkbook } from '../../../commands/coerce-input.js';
import {
  listThreadedComments,
  recordThreadedChange,
  removeThreadedComment,
} from '../../../commands/comment.js';
import { threadWrite } from '../../../commands/comment-actions.js';
import { selectNextFormulaError } from '../../../commands/error-indicators.js';
import { interactionControllerFor } from '../../../commands/interaction-controller.js';
import { setMarginPreset, setPageOrientation, setPaperSize } from '../../../commands/page-setup.js';
import { pageSetupWrite } from '../../../commands/page-setup-actions.js';
import { FUNCTION_SIGNATURES } from '../../../commands/refs.js';
import { executeNoteVisibilityAction } from '../../../commands/ribbon-comment.js';
import { MAX_COL, MAX_ROW } from '../../../engine/address.js';
import {
  openCommentDialogFrom,
  openDataValidationDialogFrom,
  openDefineNameDialogFrom,
  openPageSetupDialogFrom,
} from '../../../interact/dialog-entry.js';
import { focusSheet } from '../../../interact/dialog-focus.js';
import { ensureMacInk } from '../../../interact/mac-ink.js';
import { isNavigationAddrAllowed } from '../../../interact/navigation-policy.js';
import { createChartVariantFromSelection } from '../../../mount/dynamic-dropdown-defaults/insert.js';
import {
  convertNotesFrom,
  selectAdjacentComment,
} from '../../../mount/dynamic-dropdown-defaults/review.js';
import { showCreateTableDialog } from '../../../mount/dynamic-dropdown-defaults/styles.js';
import { createDefaultDynamicDropdownsCtx } from '../../../mount/dynamic-dropdowns-defaults.js';
import type { SpreadsheetInstance } from '../../../mount/types.js';
import { applyWorkbookTheme } from '../../../mount/workbook-theme-action.js';
import { mutators } from '../../../store/store.js';
import { errorCheckingEnabled } from '../../../store/types.js';
import { reportDialogLabels, showReport } from '../../dialogs/report.js';
import type { ApplyRibbonCommandDeps } from '../apply-ribbon-command.js';
import { createMacRibbonActions, MAC_RIBBON_ACTION_IDS } from './actions.js';
import { MAC_CONTEXT_LABELS } from './labels-data.js';
import { toolbarLangForLocale } from './locale.js';
import { type LiveFunctionNames, liveNameSet } from './menus.js';
import { reportMacRibbonError } from './report-error.js';
import { isThreadedCommentCommand } from './state-projection.js';

const aliases: Readonly<Record<string, string>> = {
  'mac.home.addins': 'addIn',
  'mac.insert.link': 'hyperlinkInsert',
  'mac.page.background': 'sheetBackground',
  'mac.page.printTitles': 'printTitles',
  'mac.page.pageSetup': 'pageSetupAdvanced',
  'mac.page.showGridlines': 'pageLayoutGridlinesView',
  'mac.page.printGridlines': 'pageLayoutGridlinesPrint',
  'mac.page.showHeadings': 'pageLayoutHeadingsView',
  'mac.page.printHeadings': 'pageLayoutHeadingsPrint',
  'mac.formulas.namesManager': 'namedRanges',
  'mac.formulas.precedents': 'precedents',
  'mac.formulas.dependents': 'dependents',
  'mac.formulas.showFormulas': 'showFormulasFormula',
  'mac.formulas.watch': 'watch',
  'mac.formulas.recalc': 'recalcNow',
  'mac.formulas.sheetRecalc': 'recalcNow',
  'mac.data.workbookLinks': 'linksData',
  'mac.data.group': 'outlineGroup',
  'mac.data.ungroup': 'outlineUngroup',
  'mac.data.showDetail': 'outlineShowDetail',
  'mac.data.hideDetail': 'outlineHideDetail',
  'mac.formulas.calcOptions': 'calcOptions',
  'mac.review.spelling': 'spellingReview',
  'mac.review.accessibility': 'accessibility',
  'mac.review.accessibility.check': 'accessibility',
  'mac.review.accessibility.formatTable': 'formatTableHome',
  'mac.review.protectSheet': 'protectReview',
  'mac.review.protectWorkbook': 'protectWorkbookReview',
  'mac.view.standard': 'viewNormal',
  'mac.view.split': 'split',
  'mac.view.pageBreakPreview': 'viewPageBreakPreview',
  'mac.view.pageLayout': 'viewPageLayout',
  'mac.view.gridlines': 'viewGridlines',
  'mac.view.headings': 'viewHeadings',
  'mac.view.formulaBar': 'viewFormulaBar',
  'mac.view.zoomSelection': 'zoomSelection',
  'mac.view.zoom': 'zoomDialog',
  'mac.view.zoom100': 'zoom100',
  'mac.view.sheetViewKeep': 'sheetViewKeep',
  'mac.view.sheetViewExit': 'sheetViewExit',
  'mac.view.sheetViewNew': 'sheetViewNew',
  'mac.view.sheetViewOptions': 'sheetViewOptions',
};

const chartKinds = {
  'mac.insert.chartColumn': 'column',
  'mac.insert.chartBar': 'bar',
  'mac.insert.chartLine': 'line',
  'mac.insert.chartArea': 'area',
  'mac.insert.chartPie': 'pie',
  'mac.insert.chartDoughnut': 'doughnut',
  'mac.insert.chartScatter': 'scatter',
} as const;

/** Gallery entries that insert a chart of a kind with a stacked grouping. */
const chartGroupings = {
  'mac.insert.chartColumnStacked': ['column', 'stacked'],
  'mac.insert.chartColumn100': ['column', 'percent-stacked'],
  'mac.insert.chartBarStacked': ['bar', 'stacked'],
  'mac.insert.chartBar100': ['bar', 'percent-stacked'],
  'mac.insert.chartLineStacked': ['line', 'stacked'],
  'mac.insert.chartLine100': ['line', 'percent-stacked'],
  'mac.insert.chartAreaStacked': ['area', 'stacked'],
  'mac.insert.chartArea100': ['area', 'percent-stacked'],
} as const;

const shapeKinds = {
  'mac.insert.shapeLine': 'line',
  'mac.insert.shapeArrow': 'arrow',
  'mac.insert.shapeRectangle': 'rectangle',
  'mac.insert.shapeRoundedRectangle': 'rounded-rectangle',
  'mac.insert.shapeOval': 'oval',
  'mac.insert.shapeTriangle': 'triangle',
  'mac.insert.shapeDiamond': 'diamond',
} as const;

const autoSumFunctions = {
  'mac.autosum.SUM': 'SUM',
  'mac.autosum.AVERAGE': 'AVERAGE',
  'mac.autosum.COUNT': 'COUNT',
  'mac.autosum.MAX': 'MAX',
  'mac.autosum.MIN': 'MIN',
} as const;

const formulaAuditActions = {
  'mac.formulas.removeArrows.all': 'clear-all',
  'mac.formulas.removeArrows.precedents': 'clear-precedents',
  'mac.formulas.removeArrows.dependents': 'clear-dependents',
  'mac.formulas.errorCheck.run': 'error-checking',
  'mac.formulas.errorCheck.trace': 'trace-error',
} as const;

const mergeActions = {
  'mac.review.accessibility.unmerge': 'unmergeCells',
} as const;

const validationActions = {
  'mac.data.validation.settings': 'settings',
  'mac.data.validation.circleInvalid': 'circle-invalid',
  'mac.data.validation.clearCircles': 'clear-circles',
  'mac.data.validation.clearRules': 'clear-rules',
} as const;

const papers = {
  'mac.page.size.a4': 'A4',
  'mac.page.size.a3': 'A3',
  'mac.page.size.letter': 'letter',
  'mac.page.size.legal': 'legal',
} as const;

const printAreaActions = {
  'mac.page.printArea.set': 'set',
  'mac.page.printArea.add': 'add',
  'mac.page.printArea.clear': 'clear',
} as const;

const sortActions = {
  'mac.data.sortAsc': 'asc',
  'mac.data.sortDesc': 'desc',
  'mac.data.sortCustom': 'custom',
  'mac.data.filter': 'filter',
  'mac.data.clear': 'filter-clear',
  'mac.data.reapply': 'filter-reapply',
  'mac.data.advancedFilter': 'filter-advanced',
  'mac.data.removeDuplicates': 'dedupe',
} as const;

const drawTools = {
  'mac.draw.eraser': 'eraser',
  'mac.draw.eraser.stroke': 'eraser',
  'mac.draw.penBlack': 'pen-black',
  'mac.draw.penRed': 'pen-red',
  'mac.draw.pencil': 'pencil',
  'mac.draw.highlighter': 'highlighter',
} as const;

// Menu availability and execution of a root's legacy primary action have
// separate permissions. Public applyCommand callers can invoke roots too.
const rootActionPermissions: Readonly<Record<string, BuiltInCommandId>> = {
  'mac.formulas.autoSum': 'mac.autosum.SUM',
  'mac.formulas.removeArrows': 'mac.formulas.removeArrows.all',
  'mac.formulas.errorCheck': 'mac.formulas.errorCheck.run',
  'mac.data.validation': 'mac.data.validation.settings',
  'mac.page.pageBreaks': 'mac.page.break.insert',
};

/** Ids the dispatcher handles in code rather than through a table above. */
const operations = [
  'mac.insert.recommendedPivotTable',
  'mac.insert.table',
  'mac.insert.photo',
  'mac.insert.photo.file',
  'mac.insert.screenshot',
  'mac.insert.screenshot.clipping',
  'mac.insert.textBox',
  'mac.insert.textBox.horizontal',
  'mac.insert.headerFooter',
  'mac.insert.symbol',
  'mac.formulas.autoSum',
  'mac.formulas.defineName',
  'mac.formulas.defineName.define',
  'mac.formulas.useInFormula',
  'mac.formulas.useInFormula.paste',
  'mac.formulas.createFromSelection',
  'mac.formulas.errorCheck',
  'mac.formulas.removeArrows',
  'mac.formulas.calc.auto',
  'mac.formulas.calc.manual',
  'mac.formulas.calc.autoNoTable',
  'mac.formulas.calc.formatStale',
  'mac.formulas.calc.compatVersion.1',
  'mac.formulas.calc.compatVersion.2',
  'mac.formulas.calc.compatVersion.3',
  'mac.page.break.insert',
  'mac.page.break.remove',
  'mac.page.break.reset',
  'mac.data.flashFill',
  'mac.data.textToColumns',
  'mac.data.validation',
  'mac.page.theme.office',
  'mac.page.theme.gray',
  'mac.page.color.office',
  'mac.page.color.gray',
  'mac.page.font.office',
  'mac.page.font.gray',
  'mac.page.theme',
  'mac.page.color',
  'mac.page.font',
  'mac.page.margins.normal',
  'mac.page.margins.wide',
  'mac.page.margins.narrow',
  'mac.page.margins.custom',
  'mac.page.orientation.portrait',
  'mac.page.orientation.landscape',
  'mac.page.pageBreaks',
  'mac.view.cellFocus.focusCell',
  'mac.view.cellFocus.autoHighlight',
  'mac.view.freeze',
  'mac.view.freeze.panes',
  'mac.view.freeze.firstRow',
  'mac.view.freeze.firstColumn',
  'mac.insert.newComment',
  'mac.review.newComment',
  'mac.review.deleteComment',
  'mac.review.previousComment',
  'mac.review.nextComment',
  'mac.review.showComments',
  'mac.review.previousNote',
  'mac.review.nextNote',
  'mac.review.newNote',
  'mac.review.showHideNote',
  'mac.review.showNotes',
  'mac.review.convertNotes',
  'mac.draw.toggle',
  'mac.draw.trackpad',
] as const;

const SUPPORTED_COMMAND_IDS: ReadonlySet<string> = new Set([
  ...Object.keys(aliases),
  'mac.insert.recommendedChart',
  'mac.insert.recommendedChart.apply',
  ...Object.keys(chartKinds),
  ...Object.keys(chartGroupings),
  ...Object.keys(shapeKinds),
  ...Object.keys(autoSumFunctions),
  ...Object.keys(formulaAuditActions),
  ...Object.keys(validationActions),
  ...Object.keys(mergeActions),
  ...Object.keys(papers),
  ...Object.keys(printAreaActions),
  ...Object.keys(sortActions),
  ...Object.keys(drawTools),
  ...operations,
  ...MAC_RIBBON_ACTION_IDS,
]);

export const isMacRibbonCommandSupported = (id: string, liveNames?: LiveFunctionNames): boolean => {
  const name = macFunctionNameOf(id);
  if (name !== null) {
    const liveSet = liveNameSet(liveNames);
    return liveSet === undefined ? name in FUNCTION_SIGNATURES : liveSet.has(name);
  }
  return SUPPORTED_COMMAND_IDS.has(id);
};

const invoke = (instance: SpreadsheetInstance, operation: () => unknown): void => {
  try {
    const result = operation();
    if (result instanceof Promise)
      void result.catch((error: unknown) => reportMacRibbonError(instance, error));
  } catch (error) {
    reportMacRibbonError(instance, error);
  }
};

const runMacErrorCheck = (instance: SpreadsheetInstance): Promise<void> | void => {
  const controller = interactionControllerFor(instance.store);
  if (controller && !controller.canSelect().allowed) return;

  const state = instance.store.getState();
  const range = {
    sheet: state.data.sheetIndex,
    r0: 0,
    c0: 0,
    r1: MAX_ROW,
    c1: MAX_COL,
  };
  const next = selectNextFormulaError(
    instance.store,
    range,
    (addr) => isNavigationAddrAllowed(instance.store, addr),
    {
      macNa: true,
      numberAsText: true,
      formulaRules: true,
      errorChecking: errorCheckingEnabled(state.ui.errorChecking),
      locale: inputLocaleForWorkbook(instance.workbook),
      workbook: instance.workbook,
    },
  );
  if (next) {
    focusSheet(instance.host);
    return;
  }

  const strings = instance.i18n.strings;
  return showReport({
    title: strings.ribbonMenu.errorChecking,
    items: [],
    ...reportDialogLabels(strings),
  }).then(() => focusSheet(instance.host));
};

export function dispatchMacRibbonCommand(
  id: string,
  deps: ApplyRibbonCommandDeps,
  dispatchGeneric: (id: string) => boolean,
): boolean {
  const instance = deps.inst;
  if (!instance) return false;
  // Only a per-function command is checked against the live catalog.
  const reportedNames =
    macFunctionNameOf(id) === null
      ? undefined
      : typeof instance.workbook?.functionNames === 'function'
        ? instance.workbook.functionNames()
        : null;
  if (!isMacRibbonCommandSupported(id, reportedNames)) return false;
  const requiredAction = rootActionPermissions[id];
  if (requiredAction && !canExecuteBuiltIn(instance.store, requiredAction, 'ribbon').allowed)
    return true;
  if (
    isThreadedCommentCommand(id) &&
    (!instance.workbook?.capabilities?.threadedComments ||
      !canExecuteBuiltInRequest(instance.store, id, 'ribbon').allowed)
  )
    return true;
  const alias = aliases[id];
  if (alias) return dispatchGeneric(alias);
  const run = (): unknown => {
    const themeAction = /^mac\.page\.(theme|color|font)\.(office|gray)$/.exec(id);
    if (themeAction) {
      if (!canExecuteBuiltInRequest(instance.store, id, 'ribbon').allowed) return;
      return applyWorkbookTheme(
        instance,
        themeAction[2] as 'office' | 'gray',
        themeAction[1] === 'color' ? 'colors' : themeAction[1] === 'font' ? 'fonts' : 'all',
        id,
      );
    }
    if (['mac.page.theme', 'mac.page.color', 'mac.page.font'].includes(id)) return;
    if (id === 'mac.formulas.errorCheck.run') return runMacErrorCheck(instance);
    let dropdowns: ReturnType<typeof createDefaultDynamicDropdownsCtx> | undefined;
    const ctx = () => (dropdowns ??= createDefaultDynamicDropdownsCtx(instance));
    const state = instance.store.getState();
    const sheet = state.data.sheetIndex;
    if (id === 'mac.insert.recommendedChart' || id === 'mac.insert.recommendedChart.apply') {
      return ctx().createRecommendedChartFromSelection();
    }
    const chart = chartKinds[id as keyof typeof chartKinds];
    if (chart) return createChartVariantFromSelection(instance, chart, undefined, id);
    const grouped = chartGroupings[id as keyof typeof chartGroupings];
    if (grouped)
      return createChartVariantFromSelection(instance, grouped[0], { grouping: grouped[1] }, id);
    const shape = shapeKinds[id as keyof typeof shapeKinds];
    if (shape) return ctx().insertShapeFromRibbon(shape);
    const autoSum = autoSumFunctions[id as keyof typeof autoSumFunctions];
    if (autoSum) return ctx().applyAutoSumFormula(autoSum);
    const auditAction = formulaAuditActions[id as keyof typeof formulaAuditActions];
    if (auditAction) return ctx().applyFormulaAuditAction(auditAction);
    const validationAction = validationActions[id as keyof typeof validationActions];
    if (validationAction) return ctx().applyDataValidationAction(validationAction);
    const mergeAction = mergeActions[id as keyof typeof mergeActions];
    if (mergeAction) return ctx().applyMergeAction(mergeAction);
    const functionName = macFunctionNameOf(id);
    if (functionName !== null) return instance.openFunctionArguments(functionName);
    // Writes run under the Mac command id that asked for them.
    const writeCtx = {
      store: instance.store,
      history: instance.history,
      origin: 'ribbon',
      commandId: id,
    } as const;
    const pageSetup = <T>(write: () => T): T | undefined =>
      pageSetupWrite(writeCtx, sheet, undefined, write);
    if (id.startsWith('mac.page.margins.')) {
      const preset = id.slice('mac.page.margins.'.length);
      if (preset === 'custom')
        return openPageSetupDialogFrom(instance.store, 'margins', {
          origin: 'ribbon',
          commandId: 'mac.page.margins.custom',
        });
      if (preset === 'normal' || preset === 'wide' || preset === 'narrow')
        return pageSetup(() => setMarginPreset(instance.store, sheet, preset, instance.history));
    }
    if (id.startsWith('mac.page.orientation.')) {
      const value = id.slice('mac.page.orientation.'.length);
      if (value === 'portrait' || value === 'landscape')
        return pageSetup(() => setPageOrientation(instance.store, sheet, value, instance.history));
    }
    const paper = papers[id as keyof typeof papers];
    if (paper) return pageSetup(() => setPaperSize(instance.store, sheet, paper, instance.history));
    const printAction = printAreaActions[id as keyof typeof printAreaActions];
    if (printAction) return ctx().applyPrintAreaAction(printAction);
    const sortAction = sortActions[id as keyof typeof sortActions];
    if (sortAction) return ctx().applySortMenuAction(sortAction);
    switch (id) {
      case 'mac.insert.recommendedPivotTable':
        return ctx().applyPivotTableAction('recommended');
      case 'mac.insert.table':
        return showCreateTableDialog(instance, {}, { origin: 'ribbon', commandId: id });
      case 'mac.insert.photo':
      case 'mac.insert.photo.file':
        return ctx().insertPictureFromRibbon('device');
      case 'mac.insert.screenshot':
      case 'mac.insert.screenshot.clipping':
        return ctx().insertScreenshotFromRibbon();
      case 'mac.insert.textBox':
      case 'mac.insert.textBox.horizontal':
        return ctx().insertTextBoxFromRibbon();
      case 'mac.insert.headerFooter':
        return openPageSetupDialogFrom(instance.store, 'headerFooter', {
          origin: 'ribbon',
          commandId: 'mac.insert.headerFooter',
        });
      case 'mac.insert.symbol':
        return ctx().applySymbolAction('more');
      case 'mac.formulas.autoSum':
        return ctx().applyAutoSumFormula('SUM');
      case 'mac.formulas.defineName':
      case 'mac.formulas.defineName.define':
        return openDefineNameDialogFrom(instance.store, {
          origin: 'ribbon',
          commandId: 'mac.formulas.defineName',
        });
      case 'mac.formulas.useInFormula':
      case 'mac.formulas.useInFormula.paste':
        return ctx().applyDefinedNameAction('use-formula');
      case 'mac.formulas.createFromSelection':
        return ctx().applyDefinedNameAction('create-from-selection');
      case 'mac.formulas.removeArrows':
        return ctx().applyFormulaAuditAction('clear-all');
      case 'mac.formulas.errorCheck':
        return runMacErrorCheck(instance);
      case 'mac.formulas.calc.auto':
        return ctx().applyCalcOptionAction('auto');
      case 'mac.formulas.calc.manual':
        return ctx().applyCalcOptionAction('manual');
      case 'mac.formulas.calc.autoNoTable':
        return ctx().applyCalcOptionAction('auto-no-table');
      case 'mac.formulas.calc.formatStale':
        return ctx().applyCalcOptionAction('format-stale');
      case 'mac.formulas.calc.compatVersion.1':
      case 'mac.formulas.calc.compatVersion.2':
      case 'mac.formulas.calc.compatVersion.3':
        return ctx().applyCalcOptionAction(`compat-${id.slice(-1)}`);
      case 'mac.data.flashFill':
        return ctx().applyFillDirection('flash');
      case 'mac.data.textToColumns':
        return ctx().splitTextToColumnsCustom();
      case 'mac.data.validation':
        return openDataValidationDialogFrom(instance.store, {
          origin: 'ribbon',
          commandId: 'mac.data.validation',
        });
      case 'mac.page.pageBreaks':
      case 'mac.page.break.insert':
        return ctx().applyPageBreakAction('insert');
      case 'mac.page.break.remove':
        return ctx().applyPageBreakAction('remove');
      case 'mac.page.break.reset':
        return ctx().applyPageBreakAction('reset-all');
      case 'mac.view.cellFocus.focusCell':
        mutators.setFocusCell(instance.store, !state.ui.focusCell);
        return true;
      case 'mac.view.cellFocus.autoHighlight':
        mutators.setAutoHighlight(instance.store, !state.ui.autoHighlight);
        return true;
      case 'mac.view.freeze':
      case 'mac.view.freeze.panes':
        return ctx().applyFreezeAction(
          state.layout.splitPanes !== true &&
            (state.layout.freezeRows > 0 || state.layout.freezeCols > 0)
            ? 'off'
            : 'selection',
        );
      case 'mac.view.freeze.firstRow':
        return ctx().applyFreezeAction('row');
      case 'mac.view.freeze.firstColumn':
        return ctx().applyFreezeAction('col');
      case 'mac.insert.newComment':
      case 'mac.review.newComment':
        return openCommentDialogFrom(instance.store, 'thread', { origin: 'ribbon', commandId: id });
      case 'mac.review.newNote':
        return openCommentDialogFrom(instance.store, 'note', { origin: 'ribbon', commandId: id });
      case 'mac.review.deleteComment': {
        const active = state.selection.active;
        const roots = listThreadedComments(instance.store, sheet, instance.workbook).filter(
          (entry) =>
            !entry.parentId && entry.cell.row === active.row && entry.cell.col === active.col,
        );
        return threadWrite(writeCtx, roots.length > 0 ? [active] : [], undefined, () =>
          recordThreadedChange(
            instance.history,
            instance.store,
            instance.workbook,
            sheet,
            () => {
              for (const root of roots)
                removeThreadedComment(instance.store, sheet, root.id, instance.workbook, writeCtx);
            },
            writeCtx,
          ),
        );
      }
      case 'mac.review.previousComment':
      case 'mac.review.nextComment':
        if (deps.hooks?.review?.selectComment) {
          return deps.hooks.review.selectComment(
            id === 'mac.review.nextComment' ? 1 : -1,
            'thread',
          );
        }
        return selectAdjacentComment(instance, id === 'mac.review.nextComment' ? 1 : -1, 'thread');
      case 'mac.review.previousNote':
      case 'mac.review.nextNote':
        if (deps.hooks?.review?.selectComment) {
          return deps.hooks.review.selectComment(id === 'mac.review.nextNote' ? 1 : -1, 'note');
        }
        return selectAdjacentComment(instance, id === 'mac.review.nextNote' ? 1 : -1, 'note');
      case 'mac.review.showComments':
        mutators.setCommentsPaneOpen(
          instance.store,
          !instance.store.getState().ui.commentsPaneOpen,
        );
        return true;
      case 'mac.review.convertNotes':
        return convertNotesFrom(instance, id);
      case 'mac.review.showHideNote':
      case 'mac.review.showNotes':
        return executeNoteVisibilityAction({
          store: instance.store,
          workbook: instance.workbook,
          action: id === 'mac.review.showHideNote' ? 'toggle-active' : 'toggle-all',
        });
    }
    if (id.startsWith('mac.draw.')) {
      const ink = ensureMacInk(instance);
      if (!ink)
        throw new Error(
          MAC_CONTEXT_LABELS.inkUnavailable[toolbarLangForLocale(instance.i18n.locale)],
        );
      if (id === 'mac.draw.toggle') return ink.toggle();
      if (id === 'mac.draw.trackpad')
        return ink.setTrackpadMode(!ink.isActive() || !ink.getTrackpadMode());
      const tool = drawTools[id as keyof typeof drawTools];
      if (tool) return ink.setTool(tool);
    }
    const action = createMacRibbonActions(instance)[id];
    if (action) return action();
    const unavailable = MAC_CONTEXT_LABELS.unavailableCommand[
      toolbarLangForLocale(instance.i18n.locale)
    ].replace('{id}', id);
    throw new Error(unavailable);
  };
  invoke(instance, run);
  deps.runtime.projectFormatToolbar();
  return true;
}
