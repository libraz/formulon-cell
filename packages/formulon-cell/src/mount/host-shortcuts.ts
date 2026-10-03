import { deleteCells, insertCells } from '../commands/cell-shift.js';
import { insertCopiedBand } from '../commands/clipboard/insert-copied-cells.js';
import type { ClipboardSnapshot } from '../commands/clipboard/snapshot.js';
import { executeRibbonFillAction } from '../commands/fill.js';
import { clearFilter, recordFilterChange, setAutoFilter } from '../commands/filter.js';
import {
  applySelectionFormatPatch,
  setNumFmt,
  toggleBold,
  toggleItalic,
  toggleStrike,
  toggleUnderline,
  withSelectionFormatOrigin,
} from '../commands/format.js';
import { formatAsTable } from '../commands/format-as-table.js';
import { type History, recordFormatChange, recordTablesChange } from '../commands/history.js';
import { interactionControllerFor } from '../commands/interaction-controller.js';
import {
  deleteCols,
  deleteRows,
  hideCols,
  hideRows,
  insertCols,
  insertRows,
  showColsAroundSelection,
  showRowsAroundSelection,
} from '../commands/structure.js';
import { MAX_COL, MAX_ROW } from '../engine/address.js';
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

const isWholeRowSelection = (range: { c0: number; c1: number }): boolean =>
  range.c0 === 0 && range.c1 >= MAX_COL;

const isWholeColumnSelection = (range: { r0: number; r1: number }): boolean =>
  range.r0 === 0 && range.r1 >= MAX_ROW;

const hasActiveCopy = (state: ReturnType<SpreadsheetStore['getState']>): boolean =>
  Boolean(state.ui.copyRange || state.ui.copyRanges?.length);

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
  getClipboardSnapshot?: () => ClipboardSnapshot | null;
  findReplace: () => { open(tab?: 'find' | 'replace'): void } | null;
  formatDialog: () => { open(): void } | null;
  formatPainter: () => { activate(sticky?: boolean): void } | null;
  goToDialog: () => { open(): void } | null;
  history: History;
  host: HTMLElement;
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
    const restricted = interactionControllerFor(input.store)?.policy !== undefined;
    const meta = e.ctrlKey || e.metaKey;
    const macPlatform =
      (input.host.closest<HTMLElement>('.fc-host') ?? input.host).dataset.fcPlatform === 'mac';
    const k = e.key.toLowerCase();
    // Control-U starts cell editing on Excel for Mac. The sheet keyboard
    // router owns the edit transition; keep this host-level format handler
    // from treating the same key as the Windows underline shortcut.
    if (macPlatform && e.ctrlKey && !e.metaKey && !e.shiftKey && !e.altKey && k === 'u') {
      e.preventDefault();
      return;
    }
    // Cmd+Control+V is Paste Special on Excel for Mac. Handling it here as
    // well as in the grid router covers feature re-attachment order; the
    // first listener to see the event stops the other route from firing.
    const target = e.target;
    const textEditingTarget =
      target instanceof HTMLInputElement ||
      target instanceof HTMLTextAreaElement ||
      (target instanceof HTMLElement && target.isContentEditable);
    if (
      macPlatform &&
      e.metaKey &&
      e.ctrlKey &&
      !e.shiftKey &&
      !e.altKey &&
      k === 'v' &&
      !textEditingTarget &&
      input.store.getState().ui.editor.kind === 'idle'
    ) {
      const dialog = input.pasteSpecialDialog();
      if (!dialog) return;
      e.preventDefault();
      e.stopImmediatePropagation();
      if (!restricted) dialog.open();
      return;
    }
    const applyDirectNumberFormat = (action: NumberFormatAction): void => {
      const fmt = numberFormatForAction(action, input.locale);
      if (!fmt) return;
      recordFormatChange(
        input.history,
        input.store,
        () => {
          withSelectionFormatOrigin(
            input.store,
            'keyboard',
            () =>
              setNumFmt(
                input.store.getState(),
                input.store,
                action === 'fixed' ? { kind: 'fixed', decimals: 2, thousands: true } : fmt,
              ),
            'numberFormat',
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
          if (value === undefined)
            withSelectionFormatOrigin(
              input.store,
              'keyboard',
              () => toggle(state, input.store),
              key,
            );
          else
            applySelectionFormatPatch(
              state,
              input.store,
              { [key]: value },
              {
                origin: 'keyboard',
                commandId: key,
              },
            );
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
      if (restricted) {
        e.preventDefault();
        return;
      }
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
      if (restricted) {
        e.preventDefault();
        return;
      }
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
    const insertCellsShortcut =
      (e.shiftKey && (e.key === '+' || e.code === 'Equal')) || e.code === 'NumpadAdd';
    const deleteCellsShortcut =
      !e.shiftKey && (e.key === '-' || e.code === 'Minus' || e.code === 'NumpadSubtract');
    if (insertCellsShortcut || deleteCellsShortcut) {
      e.preventDefault();
      if (restricted) return;
      const kind = insertCellsShortcut ? 'insert' : 'delete';
      const state = input.store.getState();
      const selected = state.selection.range;
      const range = {
        sheet: selected.sheet,
        r0: Math.min(selected.r0, selected.r1),
        r1: Math.max(selected.r0, selected.r1),
        c0: Math.min(selected.c0, selected.c1),
        c1: Math.max(selected.c0, selected.c1),
      };
      if (kind === 'insert' && hasActiveCopy(state)) {
        const snapshot = input.getClipboardSnapshot?.();
        const sourceRange = snapshot?.logicalRange ?? snapshot?.range;
        const sourceIsWholeBand =
          sourceRange !== undefined &&
          (isWholeRowSelection(sourceRange) || isWholeColumnSelection(sourceRange));
        if (snapshot && sourceIsWholeBand) {
          const inserted = insertCopiedBand(input.store, currentWb, input.history, snapshot, range);
          if (inserted) {
            mutators.replaceCells(
              input.store,
              currentWb.cells(input.store.getState().data.sheetIndex),
            );
            input.invalidate();
            return;
          }
          return;
        }
      }
      if (
        (kind === 'delete' || !hasActiveCopy(state)) &&
        (isWholeRowSelection(range) || isWholeColumnSelection(range))
      ) {
        if (isWholeRowSelection(range)) {
          const count = range.r1 - range.r0 + 1;
          if (kind === 'insert') insertRows(input.store, currentWb, input.history, range.r0, count);
          else deleteRows(input.store, currentWb, input.history, range.r0, count);
        } else {
          const count = range.c1 - range.c0 + 1;
          if (kind === 'insert') insertCols(input.store, currentWb, input.history, range.c0, count);
          else deleteCols(input.store, currentWb, input.history, range.c0, count);
        }
        mutators.replaceCells(input.store, currentWb.cells(input.store.getState().data.sheetIndex));
        input.invalidate();
        return;
      }
      openCellShiftDialog({
        host: input.host,
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
      if (restricted) return;
      applyDirectNumberFormat(numberFormatAction);
      return;
    }
    if (e.shiftKey && k === 'c') {
      const painter = input.formatPainter();
      if (!painter) return;
      e.preventDefault();
      if (restricted) return;
      painter.activate(false);
      return;
    }
    if (e.shiftKey && k === 'v') {
      const dialog = input.pasteSpecialDialog();
      if (!dialog) return;
      e.preventDefault();
      if (restricted) return;
      dialog.open();
      return;
    }
    if (e.altKey && k === 'v') {
      const dialog = input.pasteSpecialDialog();
      if (!dialog) return;
      e.preventDefault();
      if (restricted) return;
      dialog.open();
      return;
    }
    if (e.ctrlKey && !e.metaKey && k === 'q') {
      const quick = input.quickAnalysis();
      if (!quick) return;
      e.preventDefault();
      if (restricted) return;
      quick.open();
      return;
    }
    if (e.shiftKey && k === 'l') {
      e.preventDefault();
      if (restricted) return;
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
      if (restricted) return;
      recordTablesChange(input.history, input.store, () => {
        formatAsTable(input.store, input.store.getState().selection.range, { workbook: currentWb });
      });
      input.invalidate();
      return;
    }
    if (e.key === '9') {
      e.preventDefault();
      if (restricted) return;
      const range = input.store.getState().selection.range;
      if (e.shiftKey)
        showRowsAroundSelection(input.store, input.history, range.r0, range.r1, currentWb);
      else hideRows(input.store, input.history, range.r0, range.r1, currentWb);
      input.invalidate();
      return;
    }
    if (e.key === '0') {
      e.preventDefault();
      if (restricted) return;
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
      if (restricted) return;
      dialog.open();
    } else if (matchesRibbonShortcut(e, 'formatCells', 'formatCellsHome')) {
      const dialog = input.formatDialog();
      if (!dialog) return;
      e.preventDefault();
      if (restricted) return;
      dialog.open();
    } else if (e.key === '`') {
      e.preventDefault();
      mutators.setShowFormulas(input.store, !input.store.getState().ui.showFormulas);
    } else if (e.altKey && k === 'r') {
      e.preventDefault();
      mutators.setR1C1(input.store, !input.store.getState().ui.r1c1);
    } else if (e.key === ';') {
      e.preventDefault();
      if (restricted) return;
      const now = new Date();
      const utcMs = Date.UTC(now.getFullYear(), now.getMonth(), now.getDate());
      const serial = utcMs / 86_400_000 + 25569;
      currentWb.setNumber(input.store.getState().selection.active, Math.floor(serial));
      mutators.replaceCells(input.store, currentWb.cells(input.store.getState().data.sheetIndex));
    } else if (e.key === ':' || (e.shiftKey && e.key === ';')) {
      e.preventDefault();
      if (restricted) return;
      const now = new Date();
      const frac =
        (now.getUTCHours() * 3600 + now.getUTCMinutes() * 60 + now.getUTCSeconds()) / 86400;
      currentWb.setNumber(input.store.getState().selection.active, frac);
      mutators.replaceCells(input.store, currentWb.cells(input.store.getState().data.sheetIndex));
    } else if (k === 'd') {
      e.preventDefault();
      if (restricted) return;
      const r = input.store.getState().selection.range;
      if (r.r1 > r.r0) {
        const changed = executeRibbonFillAction({
          store: input.store,
          workbook: currentWb,
          history: input.history,
          action: 'down',
        });
        if (changed) {
          mutators.replaceCells(
            input.store,
            currentWb.cells(input.store.getState().data.sheetIndex),
          );
          input.invalidate();
        }
      }
    } else if (k === 'r') {
      e.preventDefault();
      if (restricted) return;
      const r = input.store.getState().selection.range;
      if (r.c1 > r.c0) {
        const changed = executeRibbonFillAction({
          store: input.store,
          workbook: currentWb,
          history: input.history,
          action: 'right',
        });
        if (changed) {
          mutators.replaceCells(
            input.store,
            currentWb.cells(input.store.getState().data.sheetIndex),
          );
          input.invalidate();
        }
      }
    } else if (k === 'e') {
      e.preventDefault();
      if (restricted) return;
      executeRibbonFillAction({
        store: input.store,
        workbook: currentWb,
        history: input.history,
        action: 'flash',
      });
      input.invalidate();
    } else if (k === 'b') {
      e.preventDefault();
      if (restricted) return;
      applyFormatToggle('bold', toggleBold);
    } else if (k === 'i') {
      e.preventDefault();
      if (restricted) return;
      applyFormatToggle('italic', toggleItalic);
    } else if (k === 'u') {
      e.preventDefault();
      if (restricted) return;
      applyFormatToggle('underline', toggleUnderline);
    } else if (e.key === '5') {
      e.preventDefault();
      if (restricted) return;
      applyFormatToggle('strike', toggleStrike);
    }
  };
}
