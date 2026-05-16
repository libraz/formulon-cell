import type { History } from '../commands/history.js';
import { formatA1FormulaAsR1C1 } from '../commands/refs.js';
import { formatCellForEdit } from '../engine/edit-seed.js';
import type { ChangeEvent, WorkbookHandle } from '../engine/workbook-handle.js';
import type { SpreadsheetEmitter } from '../events.js';
import type { ExtensionHandle, resolveFlags } from '../extensions/index.js';
import type { Strings } from '../i18n/strings.js';
import { attachAutoFillOptions } from '../interact/auto-fill-options.js';
import { attachClipboard } from '../interact/clipboard.js';
import { attachContextMenu } from '../interact/context-menu.js';
import { InlineEditor } from '../interact/editor.js';
import { attachFindReplace } from '../interact/find-replace.js';
import { attachKeyboard } from '../interact/keyboard.js';
import { attachPasteOptions } from '../interact/paste-options.js';
import { attachPasteSpecial } from '../interact/paste-special.js';
import { attachPointer } from '../interact/pointer.js';
import { attachQuickAnalysis } from '../interact/quick-analysis.js';
import {
  attachValidationAlert,
  attachValidationList,
  attachValidationPrompt,
} from '../interact/validation.js';
import type { GridRenderer } from '../render/grid.js';
import type { SpreadsheetStore } from '../store/store.js';
import { mutators } from '../store/store.js';
import type { SheetTabsController } from './sheet-tabs-controller.js';

type FeatureFlags = ReturnType<typeof resolveFlags>;

export interface EngineBinding {
  editor: InlineEditor;
  pasteSpecialDialog: ReturnType<typeof attachPasteSpecial> | null;
  findReplace: ReturnType<typeof attachFindReplace> | null;
  validation: ReturnType<typeof attachValidationList> | null;
  validationPrompt: ReturnType<typeof attachValidationPrompt> | null;
  validationAlert: ReturnType<typeof attachValidationAlert> | null;
  quickAnalysis: ReturnType<typeof attachQuickAnalysis> | null;
  clipboardH: ReturnType<typeof attachClipboard> | null;
  contextMenu: ExtensionHandle | null;
  unbind: () => void;
}

interface AttachEngineBindingInput {
  emitter: SpreadsheetEmitter;
  flags: FeatureFlags;
  getCommentDialog: () => { open(): void } | null;
  getFormatDialog: () => { open(): void } | null;
  getFormatPainter: () => { isActive(): boolean } | null;
  getGoToDialog: () => { open(mode?: 'go-to' | 'special'): void } | null;
  getHyperlinkDialog: () => { open(): void } | null;
  getNamedRangeDialog: () => { open(): void } | null;
  getPivotTableDialog: () => { open(): void } | null;
  getSessionCharts: () => unknown | null;
  getSheetTabs: () => SheetTabsController | null;
  grid: HTMLElement;
  history: History;
  host: HTMLElement;
  renderer: GridRenderer;
  store: SpreadsheetStore;
  strings: Strings;
  tag: HTMLInputElement;
  updateChrome: () => void;
  wb: WorkbookHandle;
}

export const WB_REGISTRY_IDS = [
  'clipboard',
  'pasteSpecial',
  'quickAnalysis',
  'contextMenu',
  'findReplace',
  'validation',
] as const;

export function attachEngineBinding(input: AttachEngineBindingInput): EngineBinding {
  const {
    emitter,
    flags,
    getCommentDialog,
    getFormatDialog,
    getFormatPainter,
    getGoToDialog,
    getHyperlinkDialog,
    getNamedRangeDialog,
    getPivotTableDialog,
    getSessionCharts,
    getSheetTabs,
    grid,
    history,
    host,
    renderer,
    store,
    strings,
    tag,
    updateChrome,
    wb,
  } = input;

  const refreshCells = (): void => {
    mutators.replaceCells(store, wb.cells(store.getState().data.sheetIndex));
  };

  const validationAlert = flags.validation
    ? attachValidationAlert({
        host,
        labels: {
          ok: strings.formatDialog.ok,
          stop: strings.formatDialog.validationErrorStop,
          warning: strings.formatDialog.validationErrorWarning,
          information: strings.formatDialog.validationErrorInfo,
        },
      })
    : null;

  const editor = new InlineEditor({
    host,
    grid,
    store,
    wb,
    getLabels: () => ({
      autocomplete: strings.autocomplete,
      argHelper: strings.argHelper,
    }),
    onValidation: (outcome) => validationAlert?.show(outcome),
    onAfterCommit: refreshCells,
  });
  const detachPtr = attachPointer(grid, store, wb, refreshCells, history, () =>
    editor.isActive() && editor.isFormulaEdit()
      ? {
          isFormulaEdit: () => editor.isFormulaEdit(),
          insertRefAtCaret: (ref) => editor.insertRefAtCaret(ref),
        }
      : null,
  );
  const autoFillOptions = attachAutoFillOptions({
    host: grid,
    store,
    wb,
    strings,
    history,
    onAfterCommit: refreshCells,
  });
  const pasteOptions = attachPasteOptions({
    host,
    grid,
    store,
    wb,
    strings,
    history,
    onAfterCommit: refreshCells,
  });
  // Clipboard must be attached before the keyboard router so the router can
  // forward Mod+C/X/V to the clipboard handle (browsers won't dispatch
  // copy/paste events on our non-editable, user-select:none host).
  const clipboardH = flags.clipboard
    ? attachClipboard({
        host,
        history,
        store,
        wb,
        onAfterCommit: refreshCells,
        onPasteOptions: pasteOptions.show,
      })
    : null;
  const detachKey = flags.shortcuts
    ? attachKeyboard({
        host,
        store,
        wb,
        history,
        onBeginEdit: (seed) => editor.begin(seed),
        onClearActive: () => {
          refreshCells();
          updateChrome();
        },
        onAfterHistory: refreshCells,
        onGoTo: () => {
          const goToDialog = getGoToDialog();
          if (goToDialog) {
            goToDialog.open('go-to');
            return;
          }
          tag.focus();
          tag.select();
        },
        onSwitchSheet: (delta) => getSheetTabs()?.switchRelative(delta),
        onEditComment: () => getCommentDialog()?.open(),
        onClipboardShortcut: clipboardH ? (kind) => void clipboardH.runShortcut(kind) : undefined,
      })
    : (): void => {};
  const pasteSpecialDialog =
    flags.pasteSpecial && clipboardH
      ? attachPasteSpecial({
          host,
          store,
          wb,
          strings,
          history,
          getSnapshot: () => clipboardH.getSnapshot(),
          onAfterCommit: refreshCells,
        })
      : null;
  const detachContextMenu = flags.contextMenu
    ? attachContextMenu({
        host,
        grid,
        store,
        wb,
        strings,
        onAfterCommit: refreshCells,
        onClipboardShortcut: clipboardH ? (kind) => void clipboardH.runShortcut(kind) : undefined,
        onFormatDialog: () => getFormatDialog()?.open(),
        onPasteSpecial: () => pasteSpecialDialog?.open(),
        onInsertHyperlink: () => getHyperlinkDialog()?.open(),
        onEditComment: () => getCommentDialog()?.open(),
        onDefineName: () => getNamedRangeDialog()?.open(),
        getClipboardSnapshot: clipboardH ? () => clipboardH.getSnapshot() : undefined,
        onToggleWatch: flags.watchWindow
          ? (addr) => {
              const watches = store.getState().watch.watches;
              const isOn = watches.some(
                (w) => w.sheet === addr.sheet && w.row === addr.row && w.col === addr.col,
              );
              if (isOn) mutators.removeWatch(store, addr);
              else {
                mutators.addWatch(store, addr);
                mutators.setWatchPanelOpen(store, true);
              }
            }
          : undefined,
        isWatched: flags.watchWindow
          ? (addr) =>
              store
                .getState()
                .watch.watches.some(
                  (w) => w.sheet === addr.sheet && w.row === addr.row && w.col === addr.col,
                )
          : undefined,
      })
    : (): void => {};
  const contextMenu = flags.contextMenu
    ? ({
        dispose: () => {
          detachContextMenu();
        },
      } satisfies ExtensionHandle)
    : null;
  const findReplace = flags.findReplace
    ? attachFindReplace({
        host,
        store,
        wb,
        strings,
        history,
        onAfterCommit: refreshCells,
      })
    : null;
  const validation = flags.validation
    ? attachValidationList({
        grid,
        store,
        wb,
        onAfterCommit: refreshCells,
      })
    : null;
  const validationPrompt = flags.validation ? attachValidationPrompt({ grid, store }) : null;
  const quickAnalysis = flags.quickAnalysis
    ? attachQuickAnalysis({
        host,
        store,
        wb,
        strings,
        history,
        onAfterCommit: refreshCells,
        invalidate: () => renderer.invalidate(),
        onOpenPivotTable: flags.pivotTableDialog ? () => getPivotTableDialog()?.open() : undefined,
        canOpenPivotTable: () => !!getPivotTableDialog(),
        canCreateChart: () => flags.charts && !!getSessionCharts(),
      })
    : null;

  const onDblClick = (e: MouseEvent): void => {
    if (e.button !== 0) return;
    if (editor.isActive()) return;
    if (getFormatPainter()?.isActive()) return;
    const s = store.getState();
    const a = s.selection.active;
    const key = `${a.sheet}:${a.row}:${a.col}`;
    const fmt = s.format.formats.get(key);
    const seed = formatCellForEdit(s.data.cells.get(key), wb, a, {
      formulaOverride: wb.cellFormula(a),
      formulaHidden: fmt?.formulaHidden === true,
      sheetProtected: s.protection.protectedSheets.has(a.sheet),
      formatFormula: s.ui.r1c1 ? (formula) => formatA1FormulaAsR1C1(formula, a) : undefined,
    });
    editor.begin(seed);
    e.preventDefault();
  };
  grid.addEventListener('dblclick', onDblClick);

  const unsubWb = wb.subscribe((e: ChangeEvent) => {
    if (e.kind === 'value') {
      const formula = wb.cellFormula(e.addr);
      const cell = { value: e.next, formula };
      store.setState((s) => {
        const cells = new Map(s.data.cells);
        cells.set(`${e.addr.sheet}:${e.addr.row}:${e.addr.col}`, cell);
        return { ...s, data: { ...s.data, cells } };
      });
      emitter.emit('cellChange', { addr: e.addr, value: e.next, formula });
    } else if (e.kind === 'recalc') {
      emitter.emit('recalc', { dirty: e.dirty });
    } else if (
      e.kind === 'sheet-add' ||
      e.kind === 'sheet-rename' ||
      e.kind === 'sheet-remove' ||
      e.kind === 'sheet-move'
    ) {
      getSheetTabs()?.update();
    }
  });

  return {
    editor,
    pasteSpecialDialog,
    findReplace,
    validation,
    validationPrompt,
    validationAlert,
    quickAnalysis,
    clipboardH,
    contextMenu,
    unbind: () => {
      detachPtr();
      autoFillOptions.detach();
      pasteOptions.detach();
      detachKey();
      clipboardH?.detach();
      detachContextMenu();
      findReplace?.detach();
      pasteSpecialDialog?.detach();
      validation?.detach();
      validationPrompt?.detach();
      validationAlert?.detach();
      quickAnalysis?.detach();
      grid.removeEventListener('dblclick', onDblClick);
      unsubWb();
      if (editor.isActive()) editor.cancel();
    },
  };
}
