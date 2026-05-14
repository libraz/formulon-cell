import type { History } from '../commands/history.js';
import { formatCellForEdit } from '../engine/edit-seed.js';
import type { ChangeEvent, WorkbookHandle } from '../engine/workbook-handle.js';
import type { SpreadsheetEmitter } from '../events.js';
import type { ExtensionHandle, resolveFlags } from '../extensions/index.js';
import type { Strings } from '../i18n/strings.js';
import { attachClipboard } from '../interact/clipboard.js';
import { attachContextMenu } from '../interact/context-menu.js';
import { InlineEditor } from '../interact/editor.js';
import { attachFindReplace } from '../interact/find-replace.js';
import { attachKeyboard } from '../interact/keyboard.js';
import { attachPasteSpecial } from '../interact/paste-special.js';
import { attachPointer } from '../interact/pointer.js';
import { attachQuickAnalysis } from '../interact/quick-analysis.js';
import { attachValidationList } from '../interact/validation.js';
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
  getGoToDialog: () => { open(): void } | null;
  getHyperlinkDialog: () => { open(): void } | null;
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

  const editor = new InlineEditor({
    host,
    grid,
    store,
    wb,
    getLabels: () => ({
      autocomplete: strings.autocomplete,
      argHelper: strings.argHelper,
    }),
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
  // Clipboard must be attached before the keyboard router so the router can
  // forward Mod+C/X/V to the clipboard handle (browsers won't dispatch
  // copy/paste events on our non-editable, user-select:none host).
  const clipboardH = flags.clipboard
    ? attachClipboard({
        host,
        store,
        wb,
        onAfterCommit: refreshCells,
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
            goToDialog.open();
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
        store,
        wb,
        strings,
        history,
        onAfterCommit: refreshCells,
        onFormatDialog: () => getFormatDialog()?.open(),
        onPasteSpecial: () => pasteSpecialDialog?.open(),
        onInsertHyperlink: () => getHyperlinkDialog()?.open(),
        onEditComment: () => getCommentDialog()?.open(),
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
  const quickAnalysis = flags.quickAnalysis
    ? attachQuickAnalysis({
        host,
        store,
        wb,
        strings,
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
    const seed =
      wb.cellFormula(a) ??
      formatCellForEdit(s.data.cells.get(`${a.sheet}:${a.row}:${a.col}`), wb, a);
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
    quickAnalysis,
    clipboardH,
    contextMenu,
    unbind: () => {
      detachPtr();
      detachKey();
      clipboardH?.detach();
      detachContextMenu();
      findReplace?.detach();
      pasteSpecialDialog?.detach();
      validation?.detach();
      quickAnalysis?.detach();
      grid.removeEventListener('dblclick', onDblClick);
      unsubWb();
      if (editor.isActive()) editor.cancel();
    },
  };
}
