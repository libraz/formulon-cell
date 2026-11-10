// Built-in feature factories — Extension wrappers around the `attach*`
// modules. Each factory returns an `Extension` whose id matches the
// corresponding `FeatureId`, so consumers can replace a built-in by:
//
//   const inst = await Spreadsheet.mount(host, {
//     features: { findReplace: false }, // suppress the default
//     extensions: [myCustomFindReplace()], // and substitute your own
//   });
//
// The factory wrappers call the same `attach*` functions that `mount.ts`
// uses internally for its inline construction. The two paths therefore
// behave identically, and library consumers building chrome that wraps a
// dialog (e.g., adding a brand header) can pull the factory and compose
// freely.
//
// Cross-feature dependencies (context menu opening the format dialog
// etc.) resolve through `ctx.resolve(id)` at call time, so a user
// replacement registered under the same id participates correctly.

import type { ClipboardSnapshot } from '../commands/clipboard/snapshot.js';
import type { History } from '../commands/history.js';
import { autofitPivotTableColumns, refreshPivotReportsAfterEdit } from '../commands/pivot-table.js';
import {
  clearSessionIllustration,
  duplicateSessionIllustration,
  updateSessionIllustration,
} from '../commands/session-illustration.js';
import { setSheetZoom, setWorkbookView } from '../commands/view.js';
import { addrKey } from '../engine/address.js';
import { nextCalcMode } from '../engine/calc-mode.js';
import { effectiveCellFormat } from '../engine/cell-format-projection.js';
import type { Addr } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { attachBorderDraw } from '../interact/border-draw.js';
import { attachClipboard } from '../interact/clipboard.js';
import { attachCommentDialog } from '../interact/comment-dialog.js';
import { attachConditionalDialog } from '../interact/conditional-dialog.js';
import { attachContextMenu } from '../interact/context-menu.js';
import type { DialogEntry } from '../interact/dialog-entry.js';
import { attachFindReplace } from '../interact/find-replace.js';
import { attachFormatDialog, type FormatDialogHandle } from '../interact/format-dialog.js';
import { attachFormatPainter } from '../interact/format-painter.js';
import { attachGoToDialog } from '../interact/goto-dialog.js';
import { attachHover } from '../interact/hover.js';
import { attachHyperlinkDialog, type HyperlinkDialogHandle } from '../interact/hyperlink-dialog.js';
import { attachIterativeDialog } from '../interact/iterative-dialog.js';
import {
  attachNamedRangeDialog,
  type NamedRangeDialogHandle,
} from '../interact/named-range-dialog.js';
import { attachPageSetupDialog } from '../interact/page-setup-dialog.js';
import { attachPasteSpecial } from '../interact/paste-special.js';
import { showPivotSpillCallout } from '../interact/pivot-spill-callout.js';
import { attachPivotTableDialog } from '../interact/pivot-table-dialog.js';
import { attachQuickAnalysis } from '../interact/quick-analysis.js';
import { attachSessionCharts } from '../interact/session-charts.js';
import { attachSessionIllustrations } from '../interact/session-illustrations.js';
import { attachSlicer } from '../interact/slicer.js';
import { attachStatusBar } from '../interact/status-bar.js';
import { attachValidationList } from '../interact/validation.js';
import { attachViewToolbar } from '../interact/view-toolbar.js';
import { attachWatchPanel } from '../interact/watch-panel.js';
import { attachWheel } from '../interact/wheel.js';
import { attachWorkbookObjectsPanel } from '../interact/workbook-objects.js';
import { platformLacksFeature } from '../mount/feature-flags.js';
import { mutators } from '../store/store.js';
import { resolveTheme } from '../theme/resolve.js';
import { builtIn } from './built-in-mark.js';
import type { BuiltInExtensionContext, Extension, ExtensionHandle } from './types.js';
import { isMacHost } from './ui-options.js';

const refreshCells = (ctx: BuiltInExtensionContext): void => {
  const wb = ctx.workbook;
  mutators.replaceCells(ctx.store, wb.cells(ctx.store.getState().data.sheetIndex));
};

/** Format-painter handle. Built-in id `'formatPainter'`. Pair with
 *  `features: { formatPainter: false }` to replace. */
export const formatPainter = (): Extension =>
  builtIn({
    id: 'formatPainter',
    priority: 50,
    setup(ctx: BuiltInExtensionContext) {
      const handle = attachFormatPainter({
        host: ctx.host,
        store: ctx.store,
        history: ctx.history,
        wb: () => ctx.workbook,
        formatResolver: (addr) => {
          const wb = ctx.workbook;
          return !wb.isStub && wb.capabilities.effectiveStyle
            ? effectiveCellFormat(wb, ctx.store.getState(), addr)
            : ctx.store.getState().format.formats.get(addrKey(addr));
        },
      });
      return {
        ...handle,
        dispose: handle.detach,
      };
    },
  });

/** Border draw handle. Built-in id `'borderDraw'`. Drives Excel-parity
 *  draw / draw-grid / erase modes. Pair with `features: { borderDraw: false }`
 *  to replace. */
export const borderDraw = (): Extension =>
  builtIn({
    id: 'borderDraw',
    priority: 50,
    setup(ctx: BuiltInExtensionContext) {
      const handle = attachBorderDraw({
        host: ctx.host,
        store: ctx.store,
        history: ctx.history,
      });
      return {
        ...handle,
        dispose: handle.detach,
      };
    },
  });

/** Status bar (bottom of the spreadsheet) — id `'statusBar'`. */
export const statusBar = (): Extension =>
  builtIn({
    id: 'statusBar',
    priority: 50,
    setup(ctx: BuiltInExtensionContext) {
      let handle!: ReturnType<typeof attachStatusBar>;
      handle = attachStatusBar({
        statusbar: ctx.statusbar,
        store: ctx.store,
        strings: ctx.i18n.strings,
        getEngineLabel: () => {
          const wb = ctx.workbook;
          return wb.isStub ? 'stub' : `formulon ${wb.version}`;
        },
        getCalcMode: () => ctx.workbook.calcMode(),
        onCycleCalcMode: () => {
          const wb = ctx.workbook;
          const cur = wb.calcMode();
          if (cur === null) return;
          wb.setCalcMode(nextCalcMode(cur));
          handle.refresh();
        },
        onRecalc: () => {
          const wb = ctx.workbook;
          wb.recalc();
          refreshCells(ctx);
          ctx.invalidate();
        },
        onZoomChange: (zoom) => {
          setSheetZoom(ctx.store, zoom, ctx.workbook);
          ctx.invalidate();
        },
        onViewModeChange: (mode) => {
          setWorkbookView(ctx.store, mode, ctx.workbook);
          ctx.invalidate();
        },
      });
      return {
        refresh: () => handle.refresh(),
        // attachStatusBar's setStrings updates the closure var and re-renders
        // labels in place — no detach/reattach needed.
        setStrings: (next) => handle.setStrings(next),
        rebindWorkbook: () => handle.refresh(),
        dispose: handle.detach,
      };
    },
  });

const hostPlatform = (host: HTMLElement): 'mac' | 'default' =>
  isMacHost(host) ? 'mac' : 'default';

/** Quick Analysis popover (Ctrl+Q) — id `'quickAnalysis'`. */
export const quickAnalysis = (): Extension =>
  builtIn({
    id: 'quickAnalysis',
    priority: 50,
    setup(ctx: BuiltInExtensionContext) {
      // A platform whose desktop app has no Quick Analysis gets none from the extension either.
      if (platformLacksFeature('quickAnalysis', hostPlatform(ctx.host))) {
        return { open: () => {}, close: () => {}, dispose: () => {} };
      }
      const handle = attachQuickAnalysis({
        host: ctx.host,
        store: ctx.store,
        wb: ctx.workbook,
        strings: ctx.i18n.strings,
        history: ctx.history as History,
        onAfterCommit: () => refreshCells(ctx),
        invalidate: () => ctx.invalidate(),
        onOpenPivotTable: () =>
          ctx.resolve<ExtensionHandle & { open: () => void }>('pivotTableDialog')?.open(),
        canOpenPivotTable: () => !!ctx.resolve('pivotTableDialog'),
        canCreateChart: () => !!ctx.resolve('charts'),
      });
      const onKey = (e: KeyboardEvent): void => {
        if (e.ctrlKey && !e.metaKey && e.key.toLowerCase() === 'q') {
          e.preventDefault();
          handle.open();
        }
      };
      ctx.host.addEventListener('keydown', onKey);
      return {
        open: () => handle.open(),
        close: () => handle.close(),
        rebindWorkbook: (wb) => handle.bindWorkbook(wb),
        setStrings: (next) => handle.setStrings(next),
        dispose: () => {
          ctx.host.removeEventListener('keydown', onKey);
          handle.detach();
        },
      };
    },
  });

/** Session chart overlays — id `'charts'`. */
export const charts = (): Extension =>
  builtIn({
    id: 'charts',
    priority: 50,
    setup(ctx: BuiltInExtensionContext) {
      const handle = attachSessionCharts({
        host: ctx.host,
        layer: ctx.grid.parentElement ?? ctx.host,
        store: ctx.store,
        getWb: () => ctx.workbook,
        labels: ctx.i18n.strings.sessionCharts,
        history: ctx.history,
        locale: ctx.i18n.locale,
      });
      return {
        refresh: () => handle.refresh(),
        openSettings: (tab: Parameters<typeof handle.openSettings>[0]) => handle.openSettings(tab),
        setStrings: (next) => handle.setLabels(next.sessionCharts, ctx.i18n.locale),
        dispose: () => handle.detach(),
      };
    },
  });

/** Session shape overlays — id `'illustrations'`. */
export const illustrations = (): Extension =>
  builtIn({
    id: 'illustrations',
    priority: 51,
    setup(ctx: BuiltInExtensionContext) {
      const handle = attachSessionIllustrations({
        host: ctx.host,
        layer: ctx.grid.parentElement ?? ctx.host,
        store: ctx.store,
        history: ctx.history,
        getWb: () => ctx.workbook,
        labels: ctx.i18n.strings.sessionIllustrations,
      });
      return {
        refresh: () => handle.refresh(),
        setStrings: (next) => handle.setLabels(next.sessionIllustrations),
        dispose: () => handle.detach(),
      };
    },
  });

/** PivotTable creation dialog — id `'pivotTableDialog'`. */
export const pivotTableDialog = (): Extension =>
  builtIn({
    id: 'pivotTableDialog',
    priority: 50,
    setup(ctx: BuiltInExtensionContext) {
      const handle = attachPivotTableDialog({
        host: ctx.host,
        store: ctx.store,
        wb: ctx.workbook,
        strings: ctx.i18n.strings,
        history: ctx.history,
        onAfterCreate: () => {
          refreshCells(ctx);
          const wb = ctx.workbook;
          if (wb) autofitPivotTableColumns(ctx.store, wb, { theme: resolveTheme(ctx.host) });
        },
        invalidate: () => ctx.invalidate(),
      });
      return {
        open: () => handle.open(),
        close: () => handle.close(),
        rebindWorkbook: (wb) => handle.bindWorkbook(wb),
        setStrings: (next) => handle.setStrings(next),
        dispose: () => handle.detach(),
      };
    },
  });

/** Workbook Objects side panel — id `'workbookObjects'`. */
export const workbookObjects = (): Extension =>
  builtIn({
    id: 'workbookObjects',
    priority: 50,
    setup(ctx: BuiltInExtensionContext) {
      const handle = attachWorkbookObjectsPanel({
        host: ctx.host,
        wb: ctx.workbook,
        strings: ctx.i18n.strings,
        history: ctx.history,
        store: ctx.store,
        listSessionIllustrations: () => ctx.store.getState().illustrations.illustrations,
        onSelectSessionIllustration: (id) => {
          const panel = Array.from(ctx.host.querySelectorAll<HTMLElement>('.fc-illustration')).find(
            (candidate) => candidate.dataset.illustrationId === id,
          );
          panel?.focus();
        },
        onClearSessionIllustration: (id) =>
          clearSessionIllustration(ctx.store, id, ctx.history, ctx.workbook, {
            origin: 'ribbon',
            commandId: 'workbookObjects.delete',
          }),
        onDuplicateSessionIllustration: (id) =>
          duplicateSessionIllustration(ctx.store, id, ctx.history, ctx.workbook, {
            origin: 'ribbon',
            commandId: 'workbookObjects.duplicate',
          }),
        onUpdateSessionIllustration: (id, patch) =>
          updateSessionIllustration(ctx.store, id, patch, ctx.history, ctx.workbook, {
            origin: 'ribbon',
            commandId: 'workbookObjects.update',
          }),
        subscribeSessionObjects: (listener) => ctx.store.subscribe(listener),
        onAfterPivotEdit: () => {
          const { spilled } = refreshPivotReportsAfterEdit(ctx.store, ctx.workbook, {
            theme: resolveTheme(ctx.host),
          });
          const spill = spilled[0];
          if (spill) showPivotSpillCallout(ctx.host, ctx.store, ctx.i18n.strings, spill);
          ctx.invalidate();
        },
      });
      return {
        open: () => handle.open(),
        openPivotFieldList: (sheetIndex: number, pivotIndex: number) =>
          handle.openPivotFieldList(sheetIndex, pivotIndex),
        openPivotOptions: (sheetIndex: number, pivotIndex: number) =>
          handle.openPivotOptions(sheetIndex, pivotIndex),
        openPivotFieldSettings: (addr: Addr) => handle.openPivotFieldSettings(addr),
        close: () => handle.close(),
        refresh: () => handle.refresh(),
        rebindWorkbook: (wb) => handle.bindWorkbook(wb),
        setStrings: (next) => handle.setStrings(next),
        dispose: () => handle.detach(),
      };
    },
  });

/** View toolbar — id `'viewToolbar'`. */
export const viewToolbar = (): Extension =>
  builtIn({
    id: 'viewToolbar',
    priority: 50,
    setup(ctx: BuiltInExtensionContext) {
      const objects = ctx.resolve<ExtensionHandle & { open: () => void }>('workbookObjects');
      const anchor = ctx.grid.closest<HTMLElement>('.fc-host__sheet-workspace') ?? ctx.grid;
      ctx.host.insertBefore(ctx.viewbar, anchor);
      const handle = attachViewToolbar({
        toolbar: ctx.viewbar,
        store: ctx.store,
        wb: ctx.workbook,
        history: ctx.history as History,
        strings: ctx.i18n.strings,
        onOpenObjects: objects ? () => objects.open() : undefined,
        onChanged: () => ctx.invalidate(),
      });
      return {
        refresh: () => handle.refresh(),
        rebindWorkbook: (wb) => handle.bindWorkbook(wb),
        setStrings: (next) => handle.setStrings(next),
        dispose: () => {
          handle.detach();
          if (ctx.viewbar.parentElement === ctx.host) ctx.host.removeChild(ctx.viewbar);
        },
      };
    },
  });

/** Hover-comment popover — id `'hoverComment'`. */
export const hoverComment = (): Extension =>
  builtIn({
    id: 'hoverComment',
    priority: 50,
    setup(ctx: BuiltInExtensionContext) {
      const handle = attachHover({ grid: ctx.grid, store: ctx.store });
      return { ...handle, dispose: handle.detach };
    },
  });

/** Go To Special dialog — id `'gotoSpecial'`. */
export const gotoSpecial = (): Extension =>
  builtIn({
    id: 'gotoSpecial',
    priority: 50,
    setup(ctx: BuiltInExtensionContext) {
      let handle = attachGoToDialog({
        host: ctx.host,
        store: ctx.store,
        getWb: () => ctx.workbook,
        strings: ctx.i18n.strings,
      });
      return {
        open: () => handle.open(),
        close: () => handle.close(),
        setStrings: (next) => {
          handle.detach();
          handle = attachGoToDialog({
            host: ctx.host,
            store: ctx.store,
            getWb: () => ctx.workbook,
            strings: next,
          });
        },
        dispose: () => handle.detach(),
      };
    },
  });

/** Conditional-formatting rule manager — id `'conditional'`. */
export const conditional = (): Extension =>
  builtIn({
    id: 'conditional',
    priority: 50,
    setup(ctx: BuiltInExtensionContext) {
      let handle = attachConditionalDialog({
        host: ctx.host,
        store: ctx.store,
        history: ctx.history as History,
        strings: ctx.i18n.strings,
      });
      return {
        open: () => handle.open(),
        close: () => handle.close(),
        setStrings: (next) => {
          handle.detach();
          handle = attachConditionalDialog({
            host: ctx.host,
            store: ctx.store,
            history: ctx.history as History,
            strings: next,
          });
        },
        dispose: () => handle.detach(),
      };
    },
  });

/** Calculation options dialog (mode, iteration, R1C1, precision, 1904 dates)
 *  — id `'iterative'`. Always-on in the default mount; surfaced here so
 *  consumers building a custom chrome can wire their own opener. */
export const iterative = (): Extension =>
  builtIn({
    id: 'iterative',
    priority: 50,
    setup(ctx: BuiltInExtensionContext) {
      const attach = (strings: typeof ctx.i18n.strings) =>
        attachIterativeDialog({
          host: ctx.host,
          getWb: () => ctx.workbook,
          strings,
          store: ctx.store,
          history: ctx.history,
        });
      let handle = attach(ctx.i18n.strings);
      return {
        open: () => handle.open(),
        close: () => handle.close(),
        setStrings: (next) => {
          handle.detach();
          handle = attach(next);
        },
        dispose: () => handle.detach(),
      };
    },
  });

/** Page Setup dialog — id `'pageSetup'`. */
export const pageSetup = (): Extension =>
  builtIn({
    id: 'pageSetup',
    priority: 50,
    setup(ctx: BuiltInExtensionContext) {
      let handle = attachPageSetupDialog({
        host: ctx.host,
        store: ctx.store,
        strings: ctx.i18n.strings,
        history: ctx.history as History,
      });
      return {
        open: () => handle.open(),
        close: () => handle.close(),
        setStrings: (next) => {
          handle.detach();
          handle = attachPageSetupDialog({
            host: ctx.host,
            store: ctx.store,
            strings: next,
            history: ctx.history as History,
          });
        },
        dispose: () => handle.detach(),
      };
    },
  });

/** Named-range listing dialog — id `'namedRanges'`. */
export const namedRanges = (): Extension =>
  builtIn({
    id: 'namedRanges',
    priority: 50,
    setup(ctx: BuiltInExtensionContext) {
      let handle = attachNamedRangeDialog({
        host: ctx.host,
        store: ctx.store,
        wb: ctx.workbook,
        history: ctx.history as History,
        strings: ctx.i18n.strings,
      });
      return {
        open: (entry?: DialogEntry) => handle.open(entry),
        close: () => handle.close(),
        rebindWorkbook: (wb) => handle.bindWorkbook(wb),
        setStrings: (next) => {
          handle.detach();
          handle = attachNamedRangeDialog({
            host: ctx.host,
            store: ctx.store,
            wb: ctx.workbook,
            history: ctx.history as History,
            strings: next,
          });
        },
        dispose: () => handle.detach(),
      };
    },
  });

/** Hyperlink (Ctrl+K) dialog — id `'hyperlink'`. */
export const hyperlink = (): Extension =>
  builtIn({
    id: 'hyperlink',
    priority: 50,
    setup(ctx: BuiltInExtensionContext) {
      const buildHandle = (s: typeof ctx.i18n.strings): ReturnType<typeof attachHyperlinkDialog> =>
        attachHyperlinkDialog({
          host: ctx.host,
          store: ctx.store,
          strings: s,
          history: ctx.history as History,
          getWb: () => ctx.workbook,
        });
      let handle = buildHandle(ctx.i18n.strings);
      return {
        open: (entry?: DialogEntry) => handle.open(entry),
        close: () => handle.close(),
        setStrings: (next) => {
          handle.detach();
          handle = buildHandle(next);
        },
        dispose: () => handle.detach(),
      };
    },
  });

/** Comment edit (Shift+F2) dialog — id `'commentDialog'`. */
export const commentDialog = (): Extension =>
  builtIn({
    id: 'commentDialog',
    priority: 50,
    setup(ctx: BuiltInExtensionContext) {
      const buildHandle = (s: typeof ctx.i18n.strings): ReturnType<typeof attachCommentDialog> =>
        attachCommentDialog({
          host: ctx.host,
          store: ctx.store,
          strings: s,
          history: ctx.history as History,
          getWb: () => ctx.workbook,
        });
      let handle = buildHandle(ctx.i18n.strings);
      return {
        open: () => handle.open(),
        close: () => handle.close(),
        setStrings: (next) => {
          handle.detach();
          handle = buildHandle(next);
        },
        dispose: () => handle.detach(),
      };
    },
  });

/** Format-cells dialog (Ctrl+1) — id `'formatDialog'`. */
export const formatDialog = (): Extension =>
  builtIn({
    id: 'formatDialog',
    priority: 50,
    setup(ctx: BuiltInExtensionContext) {
      const buildHandle = (s: typeof ctx.i18n.strings): ReturnType<typeof attachFormatDialog> =>
        attachFormatDialog({
          host: ctx.host,
          store: ctx.store,
          strings: s,
          history: ctx.history as History,
          getWb: () => ctx.workbook,
          getLocale: () => ctx.i18n.locale,
        });
      let handle = buildHandle(ctx.i18n.strings);
      return {
        open: () => handle.open(),
        close: () => handle.close(),
        setStrings: (next) => {
          handle.detach();
          handle = buildHandle(next);
        },
        dispose: () => handle.detach(),
      };
    },
  });

/** Find/Replace dialog (Ctrl+F) — id `'findReplace'`. Per-workbook
 *  instance: rebinds when the engine swaps. */
export const findReplace = (): Extension =>
  builtIn({
    id: 'findReplace',
    priority: 50,
    setup(ctx: BuiltInExtensionContext) {
      const attach = (wb: WorkbookHandle) =>
        attachFindReplace({
          host: ctx.host,
          store: ctx.store,
          wb,
          strings: ctx.i18n.strings,
          history: ctx.history as History,
          onAfterCommit: () => refreshCells(ctx),
        });
      let handle = attach(ctx.workbook);
      return {
        open: () => handle.open(),
        close: () => handle.close(),
        rebindWorkbook: (wb) => {
          handle.detach();
          handle = attach(wb);
        },
        // attachFindReplace's handle exposes setStrings directly — relabels in
        // place without losing query state.
        setStrings: (next) => handle.setStrings(next),
        dispose: () => handle.detach(),
      };
    },
  });

/** Validation-list dropdown — id `'validation'`. Per-workbook. */
export const validation = (): Extension =>
  builtIn({
    id: 'validation',
    priority: 50,
    setup(ctx: BuiltInExtensionContext) {
      const attach = (wb: WorkbookHandle) =>
        attachValidationList({
          grid: ctx.grid,
          store: ctx.store,
          wb,
          onAfterCommit: () => refreshCells(ctx),
        });
      let handle = attach(ctx.workbook);
      return {
        rebindWorkbook: (wb) => {
          handle.detach();
          handle = attach(wb);
        },
        dispose: () => handle.detach(),
      };
    },
  });

/** OS clipboard bridge — id `'clipboard'`. Per-workbook. */
export const clipboard = (): Extension =>
  builtIn({
    id: 'clipboard',
    priority: 50,
    setup(ctx: BuiltInExtensionContext) {
      const attach = (wb: WorkbookHandle) =>
        attachClipboard({
          host: ctx.host,
          history: ctx.history as History,
          store: ctx.store,
          wb,
          onAfterCommit: () => refreshCells(ctx),
        });
      let handle = attach(ctx.workbook);
      return {
        getSnapshot: () => handle.getSnapshot(),
        rebindWorkbook: (wb) => {
          handle.detach();
          handle = attach(wb);
        },
        dispose: () => handle.detach(),
      };
    },
  });

/** Paste-special dialog — id `'pasteSpecial'`. Looks up the clipboard
 *  feature via `ctx.resolve('clipboard')` so a user-replacement
 *  participates. */
export const pasteSpecial = (): Extension =>
  builtIn({
    id: 'pasteSpecial',
    priority: 60,
    setup(ctx: BuiltInExtensionContext) {
      let activeStrings = ctx.i18n.strings;
      const buildHandle = (): ReturnType<typeof attachPasteSpecial> | null => {
        const cb = ctx.resolve<ExtensionHandle & { getSnapshot: () => unknown }>('clipboard');
        if (!cb) return null;
        return attachPasteSpecial({
          host: ctx.host,
          store: ctx.store,
          wb: ctx.workbook,
          strings: activeStrings,
          history: ctx.history as History,
          getSnapshot: () =>
            cb.getSnapshot() as ReturnType<
              typeof attachClipboard
            >['getSnapshot'] extends () => infer R
              ? R
              : never,
          onAfterCommit: () => refreshCells(ctx),
        });
      };
      let handle = buildHandle();
      return {
        open: () => handle?.open(),
        close: () => handle?.close(),
        rebindWorkbook: () => {
          handle?.detach();
          handle = buildHandle();
        },
        setStrings: (next) => {
          activeStrings = next;
          handle?.detach();
          handle = buildHandle();
        },
        dispose: () => handle?.detach(),
      };
    },
  });

/** Right-click context menu — id `'contextMenu'`. Cross-feature
 *  callbacks resolve at call time via `ctx.resolve`, so user
 *  replacements of `formatDialog` / `hyperlink` / `pasteSpecial` are
 *  honored. */
export const contextMenu = (): Extension =>
  builtIn({
    id: 'contextMenu',
    priority: 80,
    setup(ctx: BuiltInExtensionContext) {
      const callOpen = (id: string): void => {
        const handle = ctx.resolve<ExtensionHandle & { open: () => void }>(id);
        handle?.open?.();
      };
      const getClipboardSnapshot = (): ClipboardSnapshot | null => {
        const cb = ctx.resolve<ExtensionHandle & { getSnapshot?: () => ClipboardSnapshot | null }>(
          'clipboard',
        );
        return cb?.getSnapshot?.() ?? null;
      };
      const attach = (wb: WorkbookHandle) =>
        attachContextMenu({
          host: ctx.host,
          grid: ctx.grid,
          store: ctx.store,
          wb,
          strings: ctx.i18n.strings,
          history: ctx.history as History,
          onAfterCommit: () => refreshCells(ctx),
          onFormatDialog: () =>
            ctx
              .resolve<ExtensionHandle & FormatDialogHandle>('formatDialog')
              ?.open(undefined, { entry: { origin: 'contextMenu', commandId: 'formatCells' } }),
          onPasteSpecial: () => callOpen('pasteSpecial'),
          onInsertHyperlink: () =>
            ctx
              .resolve<ExtensionHandle & HyperlinkDialogHandle>('hyperlink')
              ?.open({ origin: 'contextMenu', commandId: 'insertHyperlink' }),
          onEditComment: () => callOpen('commentDialog'),
          onDefineName: () =>
            ctx
              .resolve<ExtensionHandle & NamedRangeDialogHandle>('namedRanges')
              ?.open({ origin: 'contextMenu', commandId: 'defineName' }),
          getLocale: () => ctx.i18n.locale,
          getClipboardSnapshot,
        });
      let detach = attach(ctx.workbook);
      return {
        rebindWorkbook: (wb) => {
          detach();
          detach = attach(wb);
        },
        // attachContextMenu's detacher exposes setStrings; use it directly
        // to update labels without re-attaching all listeners.
        setStrings: (next) => detach.setStrings(next),
        dispose: () => detach(),
      };
    },
  });

/** Watch Window dock — id `'watchWindow'`. Default-off in `resolveFlags`. */
export const watchWindow = (): Extension =>
  builtIn({
    id: 'watchWindow',
    priority: 50,
    setup(ctx: BuiltInExtensionContext) {
      const dock = document.createElement('div');
      dock.dataset.fcWatch = 'dock';
      dock.className = 'fc-host__watchdock';
      ctx.host.appendChild(dock);
      const handle = attachWatchPanel({
        host: dock,
        store: ctx.store,
        getWb: () => ctx.workbook,
        strings: ctx.i18n.strings,
      });
      const unsubWb = ctx.onWorkbookChange(() => handle.refresh());
      return {
        open: () => handle.open(),
        close: () => handle.close(),
        toggle: () => handle.toggle(),
        refresh: () => handle.refresh(),
        setStrings: (next) => {
          const live = handle as typeof handle & { setStrings?: (s: typeof next) => void };
          live.setStrings?.(next);
        },
        dispose: () => {
          unsubWb();
          handle.detach();
          dock.remove();
        },
      };
    },
  });

/** Slicer floating panels — id `'slicer'`. Default-off in `resolveFlags`. */
export const slicer = (): Extension =>
  builtIn({
    id: 'slicer',
    priority: 50,
    setup(ctx: BuiltInExtensionContext) {
      const handle = attachSlicer({
        host: ctx.host,
        store: ctx.store,
        getWb: () => ctx.workbook,
        history: ctx.history as History,
        strings: ctx.i18n.strings,
      });
      const unsubWb = ctx.onWorkbookChange(() => handle.refresh());
      return {
        addSlicer: (input: Parameters<typeof handle.addSlicer>[0]) => handle.addSlicer(input),
        removeSlicer: (id: string) => handle.removeSlicer(id),
        refresh: () => handle.refresh(),
        setStrings: (next) => handle.setStrings(next),
        dispose: () => {
          unsubWb();
          handle.detach();
        },
      };
    },
  });

/** Mouse-wheel scroll handler — id `'wheel'`. */
export const wheel = (): Extension =>
  builtIn({
    id: 'wheel',
    priority: 50,
    setup(ctx: BuiltInExtensionContext) {
      const detach = attachWheel({ grid: ctx.grid, store: ctx.store, getWb: () => ctx.workbook });
      return { dispose: () => detach() };
    },
  });

/** Convenience bundle — the default-on replaceable built-in factories in
 *  one array. Pair with `features: presets.minimal()` (or pick-and-mix
 *  flags) to drop the inline built-ins, then pass this array to
 *  `extensions` for the equivalent surface composed from public factories.
 *  Default-off factories (`watchWindow`, `slicer`, `viewToolbar`) stay
 *  opt-in. */
export const allBuiltIns = (): Extension[] => [
  formatPainter(),
  borderDraw(),
  quickAnalysis(),
  charts(),
  illustrations(),
  pivotTableDialog(),
  statusBar(),
  workbookObjects(),
  hoverComment(),
  gotoSpecial(),
  conditional(),
  iterative(),
  pageSetup(),
  namedRanges(),
  hyperlink(),
  commentDialog(),
  formatDialog(),
  findReplace(),
  validation(),
  clipboard(),
  pasteSpecial(),
  contextMenu(),
  wheel(),
];
