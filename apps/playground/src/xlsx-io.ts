// File menu wiring and xlsx import/export, extracted from main.ts. The factory
// owns docName and the file-menu DOM lookups; event wiring stays in main.ts and
// calls into the returned API.

import {
  buildSpreadsheetCompatibilityReport,
  dictionaries,
  type RibbonReportItem,
  type SpreadsheetInstance,
  WorkbookHandle,
} from '@libraz/formulon-cell';

export interface XlsxIoShellText {
  saved: string;
  loading: string;
  openFailed: string;
  saveFailed: string;
  saveAs: string;
  fileName: string;
  save: string;
  enterFileName: string;
}

export interface XlsxIoPromptOptions {
  title: string;
  label: string;
  initial?: string;
  okLabel?: string;
  validate?: (value: string) => string | null;
}

export interface XlsxIoMessageOptions {
  title: string;
  message: string;
}

export interface XlsxIoCtx {
  getInst: () => SpreadsheetInstance | null;
  setInst: (instance: SpreadsheetInstance | null) => void;
  ribbonLang: 'ja' | 'en';
  markDirty: () => void;
  refreshWorkbookCells: () => void;
  shellText: XlsxIoShellText;
  docState: HTMLElement | null;
  /** Returns the sheet-tabs renderer lazily — its declaration sits much later in main.ts. */
  getRenderSheetTabs: () => () => void;
  showPrompt: (options: XlsxIoPromptOptions) => Promise<string | null>;
  showMessage: (options: XlsxIoMessageOptions) => Promise<void> | void;
  /** Returns the report renderer lazily so main.ts can declare it later. */
  getShowRibbonReport: () => (title: string, items: readonly RibbonReportItem[]) => void;
}

export interface XlsxIoApi {
  openFileMenu: () => void;
  closeFileMenu: () => void;
  triggerOpen: () => void;
  triggerSave: (filename?: string) => void;
  triggerSaveAs: () => Promise<void>;
  loadXlsxFile: (file: File) => Promise<void>;
  inspectWorkbookFromBackstage: () => void;
  setDocName: (name: string) => void;
  getDocName: () => string;
}

export const createXlsxIo = (ctx: XlsxIoCtx): XlsxIoApi => {
  const {
    getInst,
    ribbonLang,
    shellText,
    docState,
    getRenderSheetTabs,
    showPrompt,
    showMessage,
    getShowRibbonReport,
  } = ctx;
  // setInst, markDirty, refreshWorkbookCells are reserved for future hooks.
  void ctx.setInst;
  void ctx.markDirty;
  void ctx.refreshWorkbookCells;

  const fileMenuBtn = document.getElementById('menu-file');
  const fileMenuDrop = document.getElementById('menu-file-dropdown');
  const fileInput = document.getElementById('file-input') as HTMLInputElement | null;

  let docName = 'Book1';

  const setDocName = (name: string): void => {
    docName = name;
    const el = document.getElementById('doc-name');
    if (el) el.textContent = name;
  };

  const openFileMenu = (): void => {
    if (!fileMenuDrop) return;
    fileMenuDrop.hidden = false;
    fileMenuBtn?.setAttribute('aria-expanded', 'true');
  };
  const closeFileMenu = (): void => {
    if (!fileMenuDrop) return;
    fileMenuDrop.hidden = true;
    fileMenuBtn?.setAttribute('aria-expanded', 'false');
  };

  const triggerOpen = (): void => fileInput?.click();

  const downloadBytes = (bytes: Uint8Array, filename: string): void => {
    const blob = new Blob([bytes as BlobPart], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 1_000);
  };

  const triggerSave = (filename = `${docName.replace(/\.xlsx$/i, '')}.xlsx`): void => {
    const inst = getInst();
    if (!inst) return;
    try {
      const bytes = inst.workbook.save();
      downloadBytes(bytes, filename);
      if (docState) docState.textContent = shellText.saved;
    } catch (err) {
      // eslint-disable-next-line no-console
      console.error('save failed', err);
      if (docState) docState.textContent = shellText.saveFailed;
    }
  };

  const triggerSaveAs = async (): Promise<void> => {
    const name = await showPrompt({
      title: shellText.saveAs,
      label: shellText.fileName,
      initial: docName,
      okLabel: shellText.save,
      validate: (value) => (value.trim() ? null : shellText.enterFileName),
    });
    if (!name) return;
    const trimmed = name.trim();
    setDocName(trimmed);
    triggerSave(trimmed.endsWith('.xlsx') ? trimmed : `${trimmed}.xlsx`);
  };

  const inspectWorkbookFromBackstage = (): void => {
    const i = getInst();
    if (!i) return;
    getShowRibbonReport()(
      dictionaries[ribbonLang].backstage.inspect,
      buildSpreadsheetCompatibilityReport(i.workbook, dictionaries[ribbonLang].workbookObjects),
    );
  };

  const loadXlsxFile = async (file: File): Promise<void> => {
    const inst = getInst();
    if (!inst) return;
    if (docState) docState.textContent = shellText.loading;
    try {
      const buf = await file.arrayBuffer();
      const next = await WorkbookHandle.loadBytes(new Uint8Array(buf));
      await inst.setWorkbook(next);
      setDocName(file.name);
      if (docState) docState.textContent = shellText.saved;
      getRenderSheetTabs()();
    } catch (err) {
      // eslint-disable-next-line no-console
      console.error('open failed', err);
      if (docState) docState.textContent = shellText.openFailed;
      void showMessage({
        title: shellText.openFailed,
        message: err instanceof Error ? err.message : String(err),
      });
    }
  };

  return {
    openFileMenu,
    closeFileMenu,
    triggerOpen,
    triggerSave,
    triggerSaveAs,
    loadXlsxFile,
    inspectWorkbookFromBackstage,
    setDocName,
    getDocName: () => docName,
  };
};
