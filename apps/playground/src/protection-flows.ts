// Sheet / workbook protection flows and allow-edit-range actions.
// Extracted from main.ts to keep ribbon wiring slim. The factory pattern lets
// the host pass in live references (instance, status bar, dialog helpers,
// localized text) without coupling this module to global state.

import {
  addAllowedEditRange,
  clearAllowedEditRanges,
  type dictionaries,
  isWorkbookStructureProtected,
  protectedSheetPassword,
  type Range,
  recordFormatChange,
  recordProtectionChange,
  type SpreadsheetInstance,
  setCellLocked,
  setWorkbookStructureProtected,
  type toolbarMenuText,
  workbookStructurePassword,
} from '@libraz/formulon-cell';

import { showMessage, showPrompt } from './dialogs.js';

export interface ProtectionFlowsCtx {
  getInst: () => SpreadsheetInstance | null;
  ribbonLang: 'ja' | 'en';
  ribbonMenuText: ReturnType<typeof toolbarMenuText>;
  shellText: { optional: string };
  protectionText: (typeof dictionaries)['ja']['protection'];
  statusMetric: HTMLElement | null;
  parseA1Range: (raw: string, sheet: number) => Range | null;
  rangeRef: (range: Range) => string;
  renderSheetTabs: () => void;
  projectFormatToolbar: () => void;
  focusSheet: () => void;
}

export interface ProtectionFlowsApi {
  runSheetProtectionFlow: () => Promise<void>;
  runWorkbookProtectionFlow: (protect: boolean) => Promise<void>;
  applyProtectAction: (action: string) => Promise<void>;
}

export const createProtectionFlows = (ctx: ProtectionFlowsCtx): ProtectionFlowsApi => {
  const {
    getInst,
    ribbonLang,
    ribbonMenuText,
    shellText,
    protectionText,
    statusMetric,
    parseA1Range,
    rangeRef,
    renderSheetTabs,
    projectFormatToolbar,
    focusSheet,
  } = ctx;

  const runSheetProtectionFlow = async (): Promise<void> => {
    const i = getInst();
    if (!i) return;
    const ja = ribbonLang === 'ja';
    const sheet = i.store.getState().data.sheetIndex;
    if (i.isSheetProtected()) {
      const saved = protectedSheetPassword(i.store.getState(), sheet);
      if (saved) {
        const entered = await showPrompt({
          title: ja ? 'シート保護の解除' : 'Unprotect Sheet',
          label: ja ? 'パスワード' : 'Password',
          initial: '',
        });
        if (entered === null) return;
        if (entered !== saved) {
          void showMessage({
            title: ja ? 'シート保護の解除' : 'Unprotect Sheet',
            message: ja ? 'パスワードが正しくありません。' : 'The password is incorrect.',
          });
          return;
        }
      }
      recordProtectionChange(i.history, i.store, i.workbook, () => {
        i.setSheetProtected(false);
      });
      focusSheet();
      return;
    }
    const password = await showPrompt({
      title: ja ? 'シートの保護' : 'Protect Sheet',
      label: ja ? 'パスワード (省略可)' : 'Password (optional)',
      initial: '',
    });
    if (password === null) return;
    recordProtectionChange(i.history, i.store, i.workbook, () => {
      i.setSheetProtected(true, password || undefined);
    });
    focusSheet();
  };

  const runWorkbookProtectionFlow = async (protect: boolean): Promise<void> => {
    const i = getInst();
    if (!i) return;
    if (protect) {
      if (isWorkbookStructureProtected(i.store.getState())) return;
      const password = await showPrompt({
        title: ribbonMenuText.protectWorkbookCommand.replace(/\.\.\.$/, ''),
        label: `${protectionText.password} (${shellText.optional})`,
        initial: '',
      });
      if (password === null) return;
      recordProtectionChange(i.history, i.store, i.workbook, () => {
        setWorkbookStructureProtected(i.store, true, password ? { password } : undefined);
      });
      if (statusMetric) statusMetric.textContent = ribbonMenuText.workbookProtectedStatus;
      renderSheetTabs();
      focusSheet();
      return;
    }

    if (!isWorkbookStructureProtected(i.store.getState())) return;
    const saved = workbookStructurePassword(i.store.getState());
    if (saved) {
      const entered = await showPrompt({
        title: ribbonMenuText.unprotectWorkbookCommand.replace(/\.\.\.$/, ''),
        label: protectionText.password,
        initial: '',
      });
      if (entered === null) return;
      if (entered !== saved) {
        void showMessage({
          title: ribbonMenuText.unprotectWorkbookCommand.replace(/\.\.\.$/, ''),
          message: ribbonMenuText.workbookIncorrectPassword,
        });
        return;
      }
    }
    recordProtectionChange(i.history, i.store, i.workbook, () => {
      setWorkbookStructureProtected(i.store, false);
    });
    if (statusMetric) statusMetric.textContent = ribbonMenuText.workbookUnprotectedStatus;
    renderSheetTabs();
    focusSheet();
  };

  const applyProtectAction = async (action: string): Promise<void> => {
    const i = getInst();
    if (!i) return;
    if (action === 'protect-sheet') {
      if (i.isSheetProtected()) return;
      await runSheetProtectionFlow();
      return;
    }
    if (action === 'unprotect-sheet') {
      if (!i.isSheetProtected()) return;
      await runSheetProtectionFlow();
      return;
    }
    if (action === 'lock-cell' || action === 'unlock-cell') {
      const locked = action === 'lock-cell';
      recordFormatChange(i.history, i.store, () => {
        setCellLocked(i.store, i.store.getState().selection.range, locked);
      });
      if (statusMetric) {
        statusMetric.textContent = locked
          ? ribbonMenuText.cellsLockedStatus
          : ribbonMenuText.cellsUnlockedStatus;
      }
      projectFormatToolbar();
      focusSheet();
      return;
    }
    if (action === 'protect-workbook') {
      await runWorkbookProtectionFlow(true);
      return;
    }
    if (action === 'unprotect-workbook') {
      await runWorkbookProtectionFlow(false);
      return;
    }
    if (action === 'allow-edit-ranges') {
      const state = i.store.getState();
      const raw = await showPrompt({
        title: ribbonMenuText.allowEditRangesDialogTitle,
        label: ribbonMenuText.allowEditRangesDialogRange,
        initial: rangeRef(state.selection.range),
        okLabel: ribbonLang === 'ja' ? 'OK' : 'OK',
        cancelLabel: ribbonLang === 'ja' ? 'キャンセル' : 'Cancel',
        validate: (value) =>
          parseA1Range(value, state.selection.active.sheet)
            ? null
            : ribbonMenuText.allowEditRangesDialogInvalid,
      });
      if (raw === null) return;
      const range = parseA1Range(raw, state.selection.active.sheet);
      if (!range) return;
      recordProtectionChange(i.history, i.store, i.workbook, () => {
        addAllowedEditRange(i.store, range, { title: rangeRef(range) });
      });
      if (statusMetric) {
        statusMetric.textContent = ribbonMenuText.allowedEditRangeAddedStatus.replace(
          '{range}',
          rangeRef(range),
        );
      }
      focusSheet();
      return;
    }
    if (action === 'clear-allowed-edit-ranges') {
      recordProtectionChange(i.history, i.store, i.workbook, () => {
        clearAllowedEditRanges(i.store, i.store.getState().data.sheetIndex);
      });
      if (statusMetric) statusMetric.textContent = ribbonMenuText.allowedEditRangesClearedStatus;
      focusSheet();
    }
  };

  return { runSheetProtectionFlow, runWorkbookProtectionFlow, applyProtectAction };
};
