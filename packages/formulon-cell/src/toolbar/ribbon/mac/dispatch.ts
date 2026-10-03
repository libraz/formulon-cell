import { canExecuteBuiltIn } from '../../../commands/built-in-command-policy.js';
import { listComments } from '../../../commands/comment.js';
import { selectNextFormulaError } from '../../../commands/error-indicators.js';
import { interactionControllerFor } from '../../../commands/interaction-controller.js';
import { setMarginPreset, setPageOrientation, setPaperSize } from '../../../commands/page-setup.js';
import { FUNCTION_SIGNATURES } from '../../../commands/refs.js';
import { colLetter, MAX_COL, MAX_ROW } from '../../../engine/address.js';
import { ensureMacInk } from '../../../interact/mac-ink.js';
import { isNavigationAddrAllowed } from '../../../interact/navigation-policy.js';
import { createDefaultDynamicDropdownsCtx } from '../../../mount/dynamic-dropdowns-defaults.js';
import type { SpreadsheetInstance } from '../../../mount/types.js';
import { reportDialogLabels, showReport } from '../../dialogs/report.js';
import type { ApplyRibbonCommandDeps } from '../apply-ribbon-command.js';
import { createMacRibbonActions, MAC_RIBBON_ACTION_IDS } from './actions.js';
import { toolbarLangForLocale } from './locale.js';
import { reportMacRibbonError } from './report-error.js';

type LiveFunctionNames = ReadonlySet<string> | readonly string[] | null | undefined;

const liveNameSet = (names: LiveFunctionNames): ReadonlySet<string> | undefined => {
  if (names === null || names === undefined) return undefined;
  return names instanceof Set ? names : new Set(names);
};

const aliases: Readonly<Record<string, string>> = {
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
  'mac.review.nextNote': 'nextCommentReview',
  'mac.review.previousNote': 'previousCommentReview',
  'mac.review.spelling': 'spellingReview',
  'mac.review.accessibility': 'accessibility',
  'mac.review.protectSheet': 'protectReview',
  'mac.review.protectWorkbook': 'protectWorkbookReview',
  'mac.view.standard': 'viewNormal',
  'mac.view.pageBreakPreview': 'viewPageBreakPreview',
  'mac.view.pageLayout': 'viewPageLayout',
  'mac.view.gridlines': 'viewGridlines',
  'mac.view.headings': 'viewHeadings',
  'mac.view.formulaBar': 'viewFormulaBar',
  'mac.view.zeros': 'viewZeros',
  'mac.view.zoomSelection': 'zoomSelection',
  'mac.view.zoom': 'zoomDialog',
  'mac.view.zoom100': 'zoom100',
  'mac.view.sheetViewSave': 'sheetViewSave',
  'mac.view.sheetViewDelete': 'sheetViewDelete',
};

const chartKinds = {
  'mac.insert.chartColumn': 'column',
  'mac.insert.chartBar': 'bar',
  'mac.insert.chartLine': 'line',
  'mac.insert.chartArea': 'area',
  'mac.insert.chartPie': 'pie',
  'mac.insert.chartScatter': 'scatter',
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
  'mac.formulas.errorCheck.ignore': 'ignore-error',
} as const;

const validationActions = {
  'mac.data.validation.settings': 'settings',
  'mac.data.validation.circleInvalid': 'circle-invalid',
  'mac.data.validation.clearCircles': 'clear-circles',
  'mac.data.validation.clearRules': 'clear-rules',
} as const;

const operations = [
  'mac.insert.recommendedPivotTable',
  'mac.insert.table',
  'mac.insert.photo',
  'mac.insert.screenshot',
  'mac.insert.headerFooter',
  'mac.insert.symbol',
  'mac.formulas.autoSum',
  'mac.formulas.defineName',
  'mac.formulas.useInFormula',
  'mac.formulas.createFromSelection',
  'mac.formulas.createNames.topRow',
  'mac.formulas.createNames.bottomRow',
  'mac.formulas.createNames.leftColumn',
  'mac.formulas.createNames.rightColumn',
  'mac.formulas.errorCheck',
  'mac.formulas.removeArrows',
  'mac.formulas.calc.auto',
  'mac.formulas.calc.manual',
  'mac.formulas.calc.autoNoTable',
  'mac.formulas.calc.iterative',
  'mac.page.break.insert',
  'mac.page.break.remove',
  'mac.page.break.reset',
  'mac.data.sortAsc',
  'mac.data.sortDesc',
  'mac.data.sortCustom',
  'mac.data.filter',
  'mac.data.clear',
  'mac.data.reapply',
  'mac.data.advancedFilter',
  'mac.data.flashFill',
  'mac.data.textToColumns',
  'mac.data.removeDuplicates',
  'mac.data.validation',
  'mac.page.margins.normal',
  'mac.page.margins.wide',
  'mac.page.margins.narrow',
  'mac.page.margins.custom',
  'mac.page.orientation.portrait',
  'mac.page.orientation.landscape',
  'mac.page.size.a4',
  'mac.page.size.a3',
  'mac.page.size.letter',
  'mac.page.size.legal',
  'mac.page.printArea.set',
  'mac.page.printArea.add',
  'mac.page.printArea.clear',
  'mac.page.pageBreaks',
  'mac.view.freeze',
  'mac.view.freeze.off',
  'mac.view.freeze.firstRow',
  'mac.view.freeze.firstColumn',
  'mac.review.newNote',
  'mac.review.showNotes',
  'mac.draw.toggle',
  'mac.draw.eraser',
  'mac.draw.penBlack',
  'mac.draw.penRed',
  'mac.draw.pencil',
  'mac.draw.highlighter',
  'mac.draw.trackpad',
] as const;

export const isMacRibbonCommandSupported = (id: string, liveNames?: LiveFunctionNames): boolean => {
  if (id.startsWith('mac.function.')) {
    const name = id.slice('mac.function.'.length);
    const liveSet = liveNameSet(liveNames);
    return liveSet === undefined ? name in FUNCTION_SIGNATURES : liveSet.has(name);
  }
  return (
    id in aliases ||
    id in chartKinds ||
    id in shapeKinds ||
    id in autoSumFunctions ||
    id in formulaAuditActions ||
    id in validationActions ||
    operations.some((operation) => operation === id) ||
    (MAC_RIBBON_ACTION_IDS as readonly string[]).includes(id)
  );
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
  const next = selectNextFormulaError(instance.store, range, (addr) =>
    isNavigationAddrAllowed(instance.store, addr),
  );
  if (next) {
    instance.host.focus();
    return;
  }

  const strings = instance.i18n.strings;
  return showReport({
    title: strings.ribbonMenu.errorChecking,
    items: [],
    ...reportDialogLabels(strings),
  }).then(() => instance.host.focus());
};

export function dispatchMacRibbonCommand(
  id: string,
  deps: ApplyRibbonCommandDeps,
  dispatchGeneric: (id: string) => boolean,
): boolean {
  const instance = deps.inst;
  if (!instance) return false;
  const reportedNames =
    typeof instance.workbook?.functionNames === 'function'
      ? instance.workbook.functionNames()
      : null;
  if (!isMacRibbonCommandSupported(id, reportedNames)) return false;
  // Menu availability and execution of a root's legacy primary action have
  // separate permissions. Public applyCommand callers can invoke roots too.
  const rootActionPermissions: Readonly<Record<string, string>> = {
    'mac.formulas.autoSum': 'mac.autosum.SUM',
    'mac.formulas.removeArrows': 'mac.formulas.removeArrows.all',
    'mac.formulas.errorCheck': 'mac.formulas.errorCheck.run',
    'mac.data.validation': 'mac.data.validation.settings',
    'mac.page.pageBreaks': 'mac.page.break.insert',
  };
  const requiredAction = rootActionPermissions[id];
  if (requiredAction && !canExecuteBuiltIn(instance.store, requiredAction, 'ribbon').allowed)
    return true;
  const alias = aliases[id];
  if (alias) return dispatchGeneric(alias);
  const run = (): unknown => {
    if (id === 'mac.formulas.errorCheck.run') return runMacErrorCheck(instance);
    const ctx = createDefaultDynamicDropdownsCtx(instance);
    const state = instance.store.getState();
    const sheet = state.data.sheetIndex;
    const chart = chartKinds[id as keyof typeof chartKinds];
    if (chart) return ctx.createChartFromSelection(chart);
    const shape = shapeKinds[id as keyof typeof shapeKinds];
    if (shape) return ctx.insertShapeFromRibbon(shape);
    const autoSum = autoSumFunctions[id as keyof typeof autoSumFunctions];
    if (autoSum) return ctx.applyAutoSumFormula(autoSum);
    const auditAction = formulaAuditActions[id as keyof typeof formulaAuditActions];
    if (auditAction) return ctx.applyFormulaAuditAction(auditAction);
    const validationAction = validationActions[id as keyof typeof validationActions];
    if (validationAction) return ctx.applyDataValidationAction(validationAction);
    if (id.startsWith('mac.function.'))
      return instance.openFunctionArguments(id.slice('mac.function.'.length));
    if (id.startsWith('mac.page.margins.')) {
      const preset = id.slice('mac.page.margins.'.length);
      if (preset === 'custom') return instance.openPageSetup('margins');
      if (preset === 'normal' || preset === 'wide' || preset === 'narrow')
        return setMarginPreset(instance.store, sheet, preset, instance.history);
    }
    if (id.startsWith('mac.page.orientation.')) {
      const value = id.slice('mac.page.orientation.'.length);
      if (value === 'portrait' || value === 'landscape')
        return setPageOrientation(instance.store, sheet, value, instance.history);
    }
    const papers = {
      'mac.page.size.a4': 'A4',
      'mac.page.size.a3': 'A3',
      'mac.page.size.letter': 'letter',
      'mac.page.size.legal': 'legal',
    } as const;
    const paper = papers[id as keyof typeof papers];
    if (paper) return setPaperSize(instance.store, sheet, paper, instance.history);
    const print = {
      'mac.page.printArea.set': 'set',
      'mac.page.printArea.add': 'add',
      'mac.page.printArea.clear': 'clear',
    } as const;
    const printAction = print[id as keyof typeof print];
    if (printAction) return ctx.applyPrintAreaAction(printAction);
    const sort = {
      'mac.data.sortAsc': 'asc',
      'mac.data.sortDesc': 'desc',
      'mac.data.sortCustom': 'custom',
      'mac.data.filter': 'filter',
      'mac.data.clear': 'filter-clear',
      'mac.data.reapply': 'filter-reapply',
      'mac.data.advancedFilter': 'filter-advanced',
      'mac.data.removeDuplicates': 'dedupe',
    } as const;
    const sortAction = sort[id as keyof typeof sort];
    if (sortAction) return ctx.applySortMenuAction(sortAction);
    const names = {
      'mac.formulas.createNames.topRow': 'create-top-row',
      'mac.formulas.createNames.bottomRow': 'create-bottom-row',
      'mac.formulas.createNames.leftColumn': 'create-left-column',
      'mac.formulas.createNames.rightColumn': 'create-right-column',
    } as const;
    const nameAction = names[id as keyof typeof names];
    if (nameAction) return ctx.applyDefinedNameAction(nameAction);
    switch (id) {
      case 'mac.insert.recommendedPivotTable':
        return ctx.applyPivotTableAction('recommended');
      case 'mac.insert.table':
        return dispatchGeneric('formatTableInsert');
      case 'mac.insert.photo':
        return ctx.insertPictureFromRibbon('device');
      case 'mac.insert.screenshot':
        return ctx.insertScreenshotFromRibbon();
      case 'mac.insert.headerFooter':
        return instance.openPageSetup('headerFooter');
      case 'mac.insert.symbol':
        return ctx.applySymbolAction('more');
      case 'mac.formulas.autoSum':
        return ctx.applyAutoSumFormula('SUM');
      case 'mac.formulas.defineName':
        return instance.openDefineNameDialog();
      case 'mac.formulas.useInFormula':
        return ctx.applyDefinedNameAction('use-formula');
      case 'mac.formulas.createFromSelection':
        return ctx.applyDefinedNameAction('create-top-row');
      case 'mac.formulas.removeArrows':
        return ctx.applyFormulaAuditAction('clear-all');
      case 'mac.formulas.errorCheck':
        return runMacErrorCheck(instance);
      case 'mac.formulas.calc.auto':
        return ctx.applyCalcOptionAction('auto');
      case 'mac.formulas.calc.manual':
        return ctx.applyCalcOptionAction('manual');
      case 'mac.formulas.calc.autoNoTable':
        return ctx.applyCalcOptionAction('auto-no-table');
      case 'mac.formulas.calc.iterative':
        return ctx.applyCalcOptionAction('iterative');
      case 'mac.data.flashFill':
        return ctx.applyFillDirection('flash');
      case 'mac.data.textToColumns':
        return ctx.splitTextToColumnsCustom();
      case 'mac.data.validation':
        return instance.openDataValidationDialog();
      case 'mac.page.pageBreaks':
      case 'mac.page.break.insert':
        return ctx.applyPageBreakAction('insert');
      case 'mac.page.break.remove':
        return ctx.applyPageBreakAction('remove');
      case 'mac.page.break.reset':
        return ctx.applyPageBreakAction('reset-all');
      case 'mac.view.freeze':
        return ctx.applyFreezeAction('selection');
      case 'mac.view.freeze.off':
        return ctx.applyFreezeAction('off');
      case 'mac.view.freeze.firstRow':
        return ctx.applyFreezeAction('row');
      case 'mac.view.freeze.firstColumn':
        return ctx.applyFreezeAction('col');
      case 'mac.review.newNote':
        return instance.openCommentDialog();
      case 'mac.review.showNotes':
        return showReport({
          title: toolbarLangForLocale(instance.i18n.locale) === 'ja' ? 'メモ' : 'Notes',
          items: listComments(state)
            .filter((note) => isNavigationAddrAllowed(instance.store, note.addr))
            .map((note) => ({
              severity: 'info',
              label: `${colLetter(note.addr.col)}${note.addr.row + 1}`,
              detail: note.text,
            })),
          ...reportDialogLabels(instance.i18n.strings),
        });
    }
    if (id.startsWith('mac.draw.')) {
      const ink = ensureMacInk(instance);
      if (!ink) throw new Error('Ink is unavailable for this spreadsheet.');
      if (id === 'mac.draw.toggle') return ink.toggle();
      if (id === 'mac.draw.trackpad')
        return ink.setTrackpadMode(!ink.isActive() || !ink.getTrackpadMode());
      const tools = {
        'mac.draw.eraser': 'eraser',
        'mac.draw.penBlack': 'pen-black',
        'mac.draw.penRed': 'pen-red',
        'mac.draw.pencil': 'pencil',
        'mac.draw.highlighter': 'highlighter',
      } as const;
      const tool = tools[id as keyof typeof tools];
      if (tool) return ink.setTool(tool);
    }
    const action = createMacRibbonActions(instance)[id];
    if (action) return action();
    throw new Error(`Unavailable ribbon command: ${id}`);
  };
  invoke(instance, run);
  deps.runtime.projectFormatToolbar();
  return true;
}
