import { fillRange } from '../commands/fill.js';
import { toggleBold, toggleItalic, toggleStrike, toggleUnderline } from '../commands/format.js';
import { type History, recordFormatChange } from '../commands/history.js';
import { flushFormatToEngine } from '../engine/cell-format-sync.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { SpreadsheetStore } from '../store/store.js';
import { mutators } from '../store/store.js';

interface HostShortcutInput {
  findReplace: () => { open(): void } | null;
  formatDialog: () => { open(): void } | null;
  formatPainter: () => { activate(sticky?: boolean): void } | null;
  goToDialog: () => { open(): void } | null;
  history: History;
  hostTag: HTMLInputElement;
  hyperlinkDialog: () => { open(): void } | null;
  invalidate: () => void;
  pasteSpecialDialog: () => { open(): void } | null;
  quickAnalysis: () => { open(): void } | null;
  store: SpreadsheetStore;
  wb: () => WorkbookHandle;
}

export function createHostShortcutHandler(input: HostShortcutInput): (e: KeyboardEvent) => void {
  return (e: KeyboardEvent): void => {
    const currentWb = input.wb();
    const meta = e.ctrlKey || e.metaKey;
    if (e.key === 'F9') {
      e.preventDefault();
      currentWb.recalc();
      mutators.replaceCells(input.store, currentWb.cells(input.store.getState().data.sheetIndex));
      input.invalidate();
      return;
    }
    if (!meta) return;
    const k = e.key.toLowerCase();
    if (e.shiftKey && k === 'c') {
      const painter = input.formatPainter();
      if (!painter) return;
      e.preventDefault();
      painter.activate(false);
      return;
    }
    if (e.shiftKey && k === 'v') {
      const dialog = input.pasteSpecialDialog();
      if (!dialog) return;
      e.preventDefault();
      dialog.open();
      return;
    }
    if (e.altKey && k === 'v') {
      const dialog = input.pasteSpecialDialog();
      if (!dialog) return;
      e.preventDefault();
      dialog.open();
      return;
    }
    if (e.ctrlKey && !e.metaKey && k === 'q') {
      const quick = input.quickAnalysis();
      if (!quick) return;
      e.preventDefault();
      quick.open();
      return;
    }
    if (k === 'f') {
      const findReplace = input.findReplace();
      if (!findReplace) return;
      e.preventDefault();
      findReplace.open();
    } else if (k === 'k') {
      const dialog = input.hyperlinkDialog();
      if (!dialog) return;
      e.preventDefault();
      dialog.open();
    } else if (k === 'a') {
      e.preventDefault();
      mutators.selectAll(input.store);
    } else if (e.key === '1') {
      const dialog = input.formatDialog();
      if (!dialog) return;
      e.preventDefault();
      dialog.open();
    } else if (e.key === '`') {
      e.preventDefault();
      mutators.setShowFormulas(input.store, !input.store.getState().ui.showFormulas);
    } else if (e.altKey && k === 'r') {
      e.preventDefault();
      mutators.setR1C1(input.store, !input.store.getState().ui.r1c1);
    } else if (e.key === ';') {
      e.preventDefault();
      const now = new Date();
      const utcMs = Date.UTC(now.getFullYear(), now.getMonth(), now.getDate());
      const serial = utcMs / 86_400_000 + 25569;
      currentWb.setNumber(input.store.getState().selection.active, Math.floor(serial));
      mutators.replaceCells(input.store, currentWb.cells(input.store.getState().data.sheetIndex));
    } else if (e.shiftKey && e.key === ':') {
      e.preventDefault();
      const now = new Date();
      const frac =
        (now.getUTCHours() * 3600 + now.getUTCMinutes() * 60 + now.getUTCSeconds()) / 86400;
      currentWb.setNumber(input.store.getState().selection.active, frac);
      mutators.replaceCells(input.store, currentWb.cells(input.store.getState().data.sheetIndex));
    } else if (k === 'd') {
      e.preventDefault();
      const r = input.store.getState().selection.range;
      if (r.r1 > r.r0) {
        fillRange(
          input.store.getState(),
          currentWb,
          { sheet: r.sheet, r0: r.r0, c0: r.c0, r1: r.r0, c1: r.c1 },
          r,
        );
        mutators.replaceCells(input.store, currentWb.cells(input.store.getState().data.sheetIndex));
      }
    } else if (k === 'r') {
      e.preventDefault();
      const r = input.store.getState().selection.range;
      if (r.c1 > r.c0) {
        fillRange(
          input.store.getState(),
          currentWb,
          { sheet: r.sheet, r0: r.r0, c0: r.c0, r1: r.r1, c1: r.c0 },
          r,
        );
        mutators.replaceCells(input.store, currentWb.cells(input.store.getState().data.sheetIndex));
      }
    } else if (k === 'b') {
      e.preventDefault();
      recordFormatChange(input.history, input.store, () => {
        toggleBold(input.store.getState(), input.store);
      });
      flushFormatToEngine(currentWb, input.store, input.store.getState().data.sheetIndex);
    } else if (k === 'i') {
      e.preventDefault();
      recordFormatChange(input.history, input.store, () => {
        toggleItalic(input.store.getState(), input.store);
      });
      flushFormatToEngine(currentWb, input.store, input.store.getState().data.sheetIndex);
    } else if (k === 'u') {
      e.preventDefault();
      recordFormatChange(input.history, input.store, () => {
        toggleUnderline(input.store.getState(), input.store);
      });
      flushFormatToEngine(currentWb, input.store, input.store.getState().data.sheetIndex);
    } else if (e.key === '5') {
      e.preventDefault();
      recordFormatChange(input.history, input.store, () => {
        toggleStrike(input.store.getState(), input.store);
      });
      flushFormatToEngine(currentWb, input.store, input.store.getState().data.sheetIndex);
    }
  };
}
