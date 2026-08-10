import { deleteCells, insertCells } from '../commands/cell-shift.js';
import { executeRibbonFillAction, fillRange } from '../commands/fill.js';
import { clearFilter, recordFilterChange, setAutoFilter } from '../commands/filter.js';
import {
  applyFormatPatch,
  setNumFmt,
  toggleBold,
  toggleItalic,
  toggleStrike,
  toggleUnderline,
} from '../commands/format.js';
import { formatAsTable } from '../commands/format-as-table.js';
import { type History, recordFormatChange, recordTablesChange } from '../commands/history.js';
import {
  hideCols,
  hideRows,
  showColsAroundSelection,
  showRowsAroundSelection,
} from '../commands/structure.js';
import { flushFormatToEngine } from '../engine/cell-format-sync.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { Strings } from '../i18n/strings.js';
import { openCellShiftDialog } from '../interact/cell-shift-dialog.js';
import { formatWithPending } from '../store/pending-format.js';
import type { SpreadsheetStore } from '../store/store.js';
import { mutators } from '../store/store.js';
import { type NumberFormatAction, numberFormatForAction } from '../toolbar/number-format.js';
import { matchesRibbonShortcut } from '../toolbar/ribbon-model.js';

const DIRECT_NUMBER_FORMAT_BY_CODE: Readonly<Record<string, NumberFormatAction>> = {
  Backquote: 'general',
  Digit1: 'fixed',
  Digit2: 'time',
  Digit3: 'shortDate',
  Digit4: 'currency',
  Digit5: 'percent',
  Digit6: 'scientific',
};

const DIRECT_NUMBER_FORMAT_BY_KEY: Readonly<Record<string, NumberFormatAction>> = {
  '~': 'general',
  '!': 'fixed',
  '@': 'time',
  '#': 'shortDate',
  $: 'currency',
  '%': 'percent',
  '^': 'scientific',
};

const directNumberFormatAction = (e: KeyboardEvent): NumberFormatAction | null =>
  (e.shiftKey && (DIRECT_NUMBER_FORMAT_BY_CODE[e.code] ?? DIRECT_NUMBER_FORMAT_BY_KEY[e.key])) ||
  null;

type RepeatableFormatFlag = 'bold' | 'italic' | 'strike' | 'underline';
type FormatToggle = (
  state: ReturnType<SpreadsheetStore['getState']>,
  store: SpreadsheetStore,
) => void;

interface HostShortcutInput {
  addSheet: () => void;
  findReplace: () => { open(tab?: 'find' | 'replace'): void } | null;
  formatDialog: () => { open(): void } | null;
  formatPainter: () => { activate(sticky?: boolean): void } | null;
  goToDialog: () => { open(): void } | null;
  history: History;
  hostTag: HTMLInputElement;
  hyperlinkDialog: () => { open(): void } | null;
  invalidate: () => void;
  namedRangeDialog: () => { open(): void } | null;
  pasteSpecialDialog: () => { open(): void } | null;
  quickAnalysis: () => { open(): void } | null;
  locale: string;
  store: SpreadsheetStore;
  strings: () => Strings;
  wb: () => WorkbookHandle;
}

export function createHostShortcutHandler(input: HostShortcutInput): (e: KeyboardEvent) => void {
  return (e: KeyboardEvent): void => {
    const currentWb = input.wb();
    const meta = e.ctrlKey || e.metaKey;
    const applyDirectNumberFormat = (action: NumberFormatAction): void => {
      const fmt = numberFormatForAction(action, input.locale);
      if (!fmt) return;
      recordFormatChange(
        input.history,
        input.store,
        () => {
          setNumFmt(
            input.store.getState(),
            input.store,
            action === 'fixed' ? { kind: 'fixed', decimals: 2, thousands: true } : fmt,
          );
        },
        { repeat: () => applyDirectNumberFormat(action) },
      );
      input.history.setRepeat(() => applyDirectNumberFormat(action));
      flushFormatToEngine(currentWb, input.store, input.store.getState().data.sheetIndex);
      input.invalidate();
    };
    const applyFormatToggle = (
      key: RepeatableFormatFlag,
      toggle: FormatToggle,
      value?: boolean,
    ): void => {
      let applied = value;
      recordFormatChange(
        input.history,
        input.store,
        () => {
          const state = input.store.getState();
          if (value === undefined) toggle(state, input.store);
          else applyFormatPatch(state, input.store, state.selection.range, { [key]: value });
          applied =
            formatWithPending(input.store.getState(), input.store.getState().selection.active)?.[
              key
            ] === true;
        },
        {
          repeat: () => {
            if (applied !== undefined) applyFormatToggle(key, toggle, applied);
          },
        },
      );
      input.history.setRepeat(() => {
        if (applied !== undefined) applyFormatToggle(key, toggle, applied);
      });
      flushFormatToEngine(currentWb, input.store, input.store.getState().data.sheetIndex);
      input.invalidate();
    };
    if (e.shiftKey && !e.ctrlKey && !e.metaKey && !e.altKey && e.key === 'F11') {
      e.preventDefault();
      input.addSheet();
      return;
    }
    if (matchesRibbonShortcut(e, 'recalcNow')) {
      e.preventDefault();
      currentWb.recalc();
      mutators.replaceCells(input.store, currentWb.cells(input.store.getState().data.sheetIndex));
      input.invalidate();
      return;
    }
    if (!e.altKey && !e.ctrlKey && !e.metaKey && !e.shiftKey && e.key === 'F4') {
      if (input.history.repeatLast()) e.preventDefault();
      return;
    }
    if (matchesRibbonShortcut(e, 'namedRanges')) {
      const dialog = input.namedRangeDialog();
      if (!dialog) return;
      e.preventDefault();
      dialog.open();
      return;
    }
    if (!meta) return;
    const k = e.key.toLowerCase();
    const insertCellsShortcut =
      e.shiftKey && (e.key === '+' || e.code === 'Equal' || e.code === 'NumpadAdd');
    const deleteCellsShortcut =
      !e.shiftKey && (e.key === '-' || e.code === 'Minus' || e.code === 'NumpadSubtract');
    if (insertCellsShortcut || deleteCellsShortcut) {
      e.preventDefault();
      const kind = insertCellsShortcut ? 'insert' : 'delete';
      openCellShiftDialog({
        strings: input.strings(),
        kind,
        onSubmit: (direction) => {
          const range = input.store.getState().selection.range;
          const changed =
            kind === 'insert'
              ? direction === 'down' || direction === 'right'
                ? insertCells(input.store, currentWb, input.history, range, direction)
                : false
              : direction === 'up' || direction === 'left'
                ? deleteCells(input.store, currentWb, input.history, range, direction)
                : false;
          if (!changed) return;
          mutators.replaceCells(
            input.store,
            currentWb.cells(input.store.getState().data.sheetIndex),
          );
          input.invalidate();
        },
      });
      return;
    }
    const numberFormatAction = directNumberFormatAction(e);
    if (numberFormatAction) {
      e.preventDefault();
      applyDirectNumberFormat(numberFormatAction);
      return;
    }
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
    if (e.shiftKey && k === 'l') {
      e.preventDefault();
      recordFilterChange(input.history, input.store, () => {
        const state = input.store.getState();
        if (state.ui.filterRange) clearFilter(state, input.store, state.ui.filterRange);
        else setAutoFilter(input.store, state.selection.range);
      });
      input.invalidate();
      return;
    }
    if (k === 't' || k === 'l') {
      e.preventDefault();
      recordTablesChange(input.history, input.store, () => {
        formatAsTable(input.store, input.store.getState().selection.range, { workbook: currentWb });
      });
      input.invalidate();
      return;
    }
    if (e.key === '9') {
      e.preventDefault();
      const range = input.store.getState().selection.range;
      if (e.shiftKey)
        showRowsAroundSelection(input.store, input.history, range.r0, range.r1, currentWb);
      else hideRows(input.store, input.history, range.r0, range.r1, currentWb);
      input.invalidate();
      return;
    }
    if (e.key === '0') {
      e.preventDefault();
      const range = input.store.getState().selection.range;
      if (e.shiftKey)
        showColsAroundSelection(input.store, input.history, range.c0, range.c1, currentWb);
      else hideCols(input.store, input.history, range.c0, range.c1, currentWb);
      input.invalidate();
      return;
    }
    if (matchesRibbonShortcut(e, 'findHome', 'findReview')) {
      const findReplace = input.findReplace();
      if (!findReplace) return;
      e.preventDefault();
      findReplace.open();
    } else if (k === 'h') {
      const findReplace = input.findReplace();
      if (!findReplace) return;
      e.preventDefault();
      findReplace.open('replace');
    } else if (matchesRibbonShortcut(e, 'hyperlinkInsert')) {
      const dialog = input.hyperlinkDialog();
      if (!dialog) return;
      e.preventDefault();
      dialog.open();
    } else if (matchesRibbonShortcut(e, 'formatCells', 'formatCellsHome')) {
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
    } else if (e.key === ':' || (e.shiftKey && e.key === ';')) {
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
          { formatting: 'with', store: input.store },
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
          { formatting: 'with', store: input.store },
        );
        mutators.replaceCells(input.store, currentWb.cells(input.store.getState().data.sheetIndex));
      }
    } else if (k === 'e') {
      e.preventDefault();
      executeRibbonFillAction({
        store: input.store,
        workbook: currentWb,
        history: input.history,
        action: 'flash',
      });
      input.invalidate();
    } else if (k === 'b') {
      e.preventDefault();
      applyFormatToggle('bold', toggleBold);
    } else if (k === 'i') {
      e.preventDefault();
      applyFormatToggle('italic', toggleItalic);
    } else if (k === 'u') {
      e.preventDefault();
      applyFormatToggle('underline', toggleUnderline);
    } else if (e.key === '5') {
      e.preventDefault();
      applyFormatToggle('strike', toggleStrike);
    }
  };
}
