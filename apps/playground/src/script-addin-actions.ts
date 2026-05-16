// Script / add-in / PDF / sheet-view ribbon actions extracted from main.ts.
// The factory owns the small helpers (review cell lookup, range label, script
// label) and exposes only the externally-called surface. `showRibbonReport` is
// surfaced on the API because other extracted modules (xlsx-io etc.) consume
// it via a getter from main.ts.

import {
  analyzeAccessibilityCells,
  analyzeSpellingCells,
  applyTextScriptToRange,
  buildTranslationReviewItems,
  colLetter,
  deleteSheetView,
  parseScriptCommand,
  type Range,
  type ReviewCell,
  type RibbonReportItem,
  recordSheetViewsChange,
  reviewCellsFromState,
  type ScriptCommand,
  type SpreadsheetInstance,
  saveSheetView,
  selectNextFormulaError,
} from '@libraz/formulon-cell';
import { showMessage, showPrompt, showReport } from './dialogs.js';

export interface ScriptAddinActionsRibbonText {
  accessibility: string;
  spelling: string;
  translate: string;
  recordActions: string;
  addIn: string;
  pdf: string;
}

export interface ScriptAddinActionsRibbonMenuText {
  scriptCommandUppercase: string;
  scriptCommandLowercase: string;
  scriptCommandTrim: string;
  scriptCommandClear: string;
  scriptDialogTitle: string;
  scriptDialogCommand: string;
  scriptCommandPrompt: string;
  scriptDialogRun: string;
  scriptCommandInvalid: string;
  automationRunStatus: string;
  automationBuiltInScriptsLabel: string;
  automationBuiltInScriptsDetail: string;
  automationRecentRunsLabel: string;
  automationRunDetail: string;
  automationNoRuns: string;
  automationScriptsTitle: string;
  recordActionsStatus: string;
  recordActionsEmpty: string;
  addInBuiltInLabel: string;
  addInBuiltInDetail: string;
  addInExternalLabel: string;
  addInExternalDetail: string;
  addInGet: string;
  addInStoreLabel: string;
  addInStoreDetail: string;
  addInManage: string;
  addInManagedStatus: string;
  addInMy: string;
  pdfShare: string;
  pdfShareReady: string;
}

export interface ScriptAddinActionsRibbonReportText {
  noIssues: string;
  info: string;
  warning: string;
}

export interface ScriptAddinActionsViewToolbarText {
  saveView: string;
  views: string;
}

export interface AutomationRun {
  label: string;
  range: string;
  changed: number;
}

export interface ScriptAddinActionsCtx {
  getInst: () => SpreadsheetInstance | null;
  ribbonLang: 'ja' | 'en';
  ribbonText: ScriptAddinActionsRibbonText;
  ribbonMenuText: ScriptAddinActionsRibbonMenuText;
  ribbonReportText: ScriptAddinActionsRibbonReportText;
  viewToolbarText: ScriptAddinActionsViewToolbarText;
  automationRuns: AutomationRun[];
  statusMetric: HTMLElement | null;
  refreshWorkbookCells: () => void;
  projectFormatToolbar: () => void;
  focusSheet: () => void;
}

export interface ScriptAddinActionsApi {
  showRibbonReport: (title: string, items: readonly RibbonReportItem[]) => void;
  runAccessibilityCheck: () => void;
  runSpellingReview: () => void;
  openTranslateReview: () => void;
  runPlaygroundScriptCommand: (op: ScriptCommand) => void;
  runPlaygroundScript: () => Promise<void>;
  applyScriptAction: (action: string) => Promise<void>;
  openAllScripts: () => void;
  recordSelectedActions: () => void;
  openAddInManager: () => void;
  applyAddInAction: (action: string) => void;
  applyPdfAction: (action: string) => void;
  runFormulaErrorChecking: () => void;
  saveCurrentSheetViewFromRibbon: () => Promise<void>;
  deleteActiveSheetViewFromRibbon: () => void;
}

export const createScriptAddinActions = (ctx: ScriptAddinActionsCtx): ScriptAddinActionsApi => {
  const {
    getInst,
    ribbonLang,
    ribbonText,
    ribbonMenuText,
    ribbonReportText,
    viewToolbarText,
    automationRuns,
    statusMetric,
    refreshWorkbookCells,
    projectFormatToolbar,
    focusSheet,
  } = ctx;

  const reviewCellsForSheet = (
    sheet: number,
    range?: { sheet: number; r0: number; c0: number; r1: number; c1: number },
  ): ReviewCell[] => {
    const inst = getInst();
    if (!inst) return [];
    return reviewCellsFromState(inst.store.getState(), sheet, range);
  };

  const showRibbonReport = (title: string, items: readonly RibbonReportItem[]): void => {
    void showReport({
      title,
      items,
      emptyLabel: ribbonReportText.noIssues,
      closeLabel: ribbonLang === 'ja' ? '閉じる' : 'Close',
      infoLabel: ribbonReportText.info,
      warningLabel: ribbonReportText.warning,
    });
  };

  const selectionRangeLabel = (range: Range): string => {
    const start = `${colLetter(range.c0)}${range.r0 + 1}`;
    const end = `${colLetter(range.c1)}${range.r1 + 1}`;
    return start === end ? start : `${start}:${end}`;
  };

  const scriptCommandLabel = (command: ScriptCommand): string => {
    switch (command) {
      case 'uppercase':
        return ribbonMenuText.scriptCommandUppercase;
      case 'lowercase':
        return ribbonMenuText.scriptCommandLowercase;
      case 'trim':
        return ribbonMenuText.scriptCommandTrim;
      case 'clear':
        return ribbonMenuText.scriptCommandClear;
    }
  };

  const runAccessibilityCheck = (): void => {
    const inst = getInst();
    if (!inst) return;
    const sheet = inst.store.getState().data.sheetIndex;
    const items = analyzeAccessibilityCells(reviewCellsForSheet(sheet), ribbonLang);
    if (statusMetric)
      statusMetric.textContent = `${ribbonText.accessibility} · ${items.filter((i) => i.severity === 'warning').length} ${ribbonReportText.warning}`;
    showRibbonReport(ribbonText.accessibility, items);
  };

  const runSpellingReview = (): void => {
    const inst = getInst();
    if (!inst) return;
    const sheet = inst.store.getState().data.sheetIndex;
    const items = analyzeSpellingCells(reviewCellsForSheet(sheet), ribbonLang);
    if (statusMetric)
      statusMetric.textContent = `${ribbonText.spelling} · ${items.filter((i) => i.severity === 'warning').length} ${ribbonReportText.warning}`;
    showRibbonReport(ribbonText.spelling, items);
  };

  const openTranslateReview = (): void => {
    const inst = getInst();
    if (!inst) return;
    const state = inst.store.getState();
    const items = buildTranslationReviewItems(
      reviewCellsForSheet(state.data.sheetIndex, state.selection.range),
      ribbonLang,
    );
    showRibbonReport(ribbonText.translate, items);
  };

  const runPlaygroundScriptCommand = (op: ScriptCommand): void => {
    const inst = getInst();
    if (!inst) return;
    const range = inst.store.getState().selection.range;
    inst.history.begin();
    let changed = 0;
    try {
      changed = applyTextScriptToRange(inst.store.getState(), inst.workbook, range, op);
    } finally {
      inst.history.end();
    }
    refreshWorkbookCells();
    automationRuns.unshift({
      label: scriptCommandLabel(op),
      range: selectionRangeLabel(range),
      changed,
    });
    automationRuns.splice(8);
    if (statusMetric)
      statusMetric.textContent = ribbonMenuText.automationRunStatus.replace(
        '{count}',
        String(changed),
      );
    focusSheet();
  };

  const runPlaygroundScript = async (): Promise<void> => {
    if (!getInst()) return;
    const command = await showPrompt({
      title: ribbonMenuText.scriptDialogTitle,
      label: ribbonMenuText.scriptDialogCommand,
      placeholder: ribbonMenuText.scriptCommandPrompt,
      okLabel: ribbonMenuText.scriptDialogRun,
      cancelLabel: ribbonLang === 'ja' ? 'キャンセル' : 'Cancel',
      validate: (value) => (parseScriptCommand(value) ? null : ribbonMenuText.scriptCommandInvalid),
    });
    if (!command || !getInst()) return;
    const op = parseScriptCommand(command);
    if (!op) return;
    runPlaygroundScriptCommand(op);
  };

  const applyScriptAction = async (action: string): Promise<void> => {
    if (action === 'custom') {
      await runPlaygroundScript();
      return;
    }
    const op = parseScriptCommand(action);
    if (!op) return;
    runPlaygroundScriptCommand(op);
  };

  const openAllScripts = (): void => {
    const t = ribbonMenuText;
    const items: RibbonReportItem[] = [
      {
        severity: 'info',
        label: t.automationBuiltInScriptsLabel,
        detail: t.automationBuiltInScriptsDetail,
      },
    ];
    if (automationRuns.length) {
      items.push(
        ...automationRuns.map((run) => ({
          severity: 'info' as const,
          label: `${t.automationRecentRunsLabel}: ${run.label}`,
          detail: t.automationRunDetail
            .replace('{command}', run.label)
            .replace('{range}', run.range)
            .replace('{count}', String(run.changed)),
        })),
      );
    } else {
      items.push({
        severity: 'info',
        label: t.automationRecentRunsLabel,
        detail: t.automationNoRuns,
      });
    }
    showRibbonReport(t.automationScriptsTitle, items);
  };

  const recordSelectedActions = (): void => {
    const inst = getInst();
    if (!inst) return;
    const range = selectionRangeLabel(inst.store.getState().selection.range);
    automationRuns.unshift({
      label: ribbonText.recordActions,
      range,
      changed: 0,
    });
    automationRuns.splice(8);
    if (statusMetric) statusMetric.textContent = `${ribbonText.recordActions} · ${range}`;
    showRibbonReport(ribbonText.recordActions, [
      {
        severity: 'info',
        label: ribbonMenuText.recordActionsStatus,
        detail: ribbonMenuText.recordActionsEmpty,
      },
    ]);
    focusSheet();
  };

  const openAddInManager = (): void => {
    showRibbonReport(ribbonText.addIn, [
      {
        severity: 'info',
        label: ribbonMenuText.addInBuiltInLabel,
        detail: ribbonMenuText.addInBuiltInDetail,
      },
      {
        severity: 'info',
        label: ribbonMenuText.addInExternalLabel,
        detail: ribbonMenuText.addInExternalDetail,
      },
    ]);
  };

  const applyAddInAction = (action: string): void => {
    const t = ribbonMenuText;
    if (action === 'get') {
      showRibbonReport(t.addInGet, [
        { severity: 'info', label: t.addInStoreLabel, detail: t.addInStoreDetail },
        { severity: 'info', label: t.addInBuiltInLabel, detail: t.addInBuiltInDetail },
      ]);
      return;
    }
    if (action === 'manage') {
      showRibbonReport(t.addInManage, [
        { severity: 'info', label: t.addInManagedStatus, detail: t.addInExternalDetail },
      ]);
      return;
    }
    if (action === 'my') {
      showRibbonReport(t.addInMy, [
        { severity: 'info', label: t.addInBuiltInLabel, detail: t.addInBuiltInDetail },
        { severity: 'info', label: t.addInExternalLabel, detail: t.addInExternalDetail },
      ]);
      return;
    }
    openAddInManager();
  };

  const applyPdfAction = (action: string): void => {
    const i = getInst();
    if (!i) return;
    if (action === 'preferences') {
      i.openPageSetup();
      return;
    }
    i.print('pdf');
    if (action === 'share') {
      showRibbonReport(ribbonText.pdf, [
        { severity: 'info', label: ribbonMenuText.pdfShare, detail: ribbonMenuText.pdfShareReady },
      ]);
    }
  };

  const runFormulaErrorChecking = (): void => {
    const i = getInst();
    if (!i) return;
    const found = selectNextFormulaError(i.store);
    if (found) {
      projectFormatToolbar();
      focusSheet();
      return;
    }
    void showMessage({
      title: ribbonLang === 'ja' ? 'エラー チェック' : 'Error Checking',
      message:
        ribbonLang === 'ja'
          ? '選択範囲に数式エラーは見つかりませんでした。'
          : 'No formula errors were found in the selected range.',
    });
  };

  const saveCurrentSheetViewFromRibbon = async (): Promise<void> => {
    const i = getInst();
    if (!i) return;
    const count = i.store.getState().sheetViews.views.length + 1;
    const defaultName = `${viewToolbarText.views} ${count}`;
    const name = await showPrompt({
      title: viewToolbarText.saveView,
      label: viewToolbarText.views,
      initial: defaultName,
      okLabel: 'OK',
      cancelLabel: ribbonLang === 'ja' ? 'キャンセル' : 'Cancel',
      validate: (value) =>
        value.trim() ? null : ribbonLang === 'ja' ? '名前を入力してください。' : 'Enter a name.',
    });
    const trimmed = name?.trim();
    if (!trimmed) {
      focusSheet();
      return;
    }
    const id = `view-${Date.now().toString(36)}-${count}`;
    recordSheetViewsChange(i.history, i.store, () => {
      saveSheetView(i.store, id, trimmed);
      i.store.setState((s) => ({ ...s, sheetViews: { ...s.sheetViews, activeViewId: id } }));
    });
    projectFormatToolbar();
    focusSheet();
  };

  const deleteActiveSheetViewFromRibbon = (): void => {
    const i = getInst();
    if (!i) return;
    const id = i.store.getState().sheetViews.activeViewId;
    if (!id) {
      focusSheet();
      return;
    }
    deleteSheetView(i.store, id, i.history);
    projectFormatToolbar();
    focusSheet();
  };

  return {
    showRibbonReport,
    runAccessibilityCheck,
    runSpellingReview,
    openTranslateReview,
    runPlaygroundScriptCommand,
    runPlaygroundScript,
    applyScriptAction,
    openAllScripts,
    recordSelectedActions,
    openAddInManager,
    applyAddInAction,
    applyPdfAction,
    runFormulaErrorChecking,
    saveCurrentSheetViewFromRibbon,
    deleteActiveSheetViewFromRibbon,
  };
};
