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
import type { History } from '../commands/history.js';
import { setSheetZoom } from '../commands/structure.js';
import { attachClipboard } from '../interact/clipboard.js';
import { attachCommentDialog } from '../interact/comment-dialog.js';
import { attachConditionalDialog } from '../interact/conditional-dialog.js';
import { attachContextMenu } from '../interact/context-menu.js';
import { attachFindReplace } from '../interact/find-replace.js';
import { attachFormatDialog } from '../interact/format-dialog.js';
import { attachFormatPainter } from '../interact/format-painter.js';
import { attachGoToDialog } from '../interact/goto-dialog.js';
import { attachHover } from '../interact/hover.js';
import { attachHyperlinkDialog } from '../interact/hyperlink-dialog.js';
import { attachIterativeDialog } from '../interact/iterative-dialog.js';
import { attachNamedRangeDialog } from '../interact/named-range-dialog.js';
import { attachPageSetupDialog } from '../interact/page-setup-dialog.js';
import { attachPasteSpecial } from '../interact/paste-special.js';
import { attachPivotTableDialog } from '../interact/pivot-table-dialog.js';
import { attachQuickAnalysis } from '../interact/quick-analysis.js';
import { attachSessionCharts } from '../interact/session-charts.js';
import { attachSlicer } from '../interact/slicer.js';
import { attachStatusBar } from '../interact/status-bar.js';
import { attachValidationList } from '../interact/validation.js';
import { attachViewToolbar } from '../interact/view-toolbar.js';
import { attachWatchPanel } from '../interact/watch-panel.js';
import { attachWheel } from '../interact/wheel.js';
import { attachWorkbookObjectsPanel } from '../interact/workbook-objects.js';
import { mutators } from '../store/store.js';
import type { Extension, ExtensionContext, ExtensionHandle } from './types.js';

const refreshCells = (ctx: ExtensionContext): void => {
  const wb = ctx.getWb();
  mutators.replaceCells(ctx.store, wb.cells(ctx.store.getState().data.sheetIndex));
};

/** Format-painter handle. Built-in id `'formatPainter'`. Pair with
 *  `features: { formatPainter: false }` to replace. */
export const formatPainter = (): Extension => ({
  id: 'formatPainter',
  priority: 50,
  setup(ctx) {
    const handle = attachFormatPainter({
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
export const statusBar = (): Extension => ({
  id: 'statusBar',
  priority: 50,
  setup(ctx) {
    let handle!: ReturnType<typeof attachStatusBar>;
    handle = attachStatusBar({
      statusbar: ctx.statusbar,
      store: ctx.store,
      strings: ctx.i18n.strings,
      getEngineLabel: () => {
        const wb = ctx.getWb();
        return wb.isStub ? 'stub' : `formulon ${wb.version}`;
      },
      getCalcMode: () => ctx.getWb().calcMode(),
      onCycleCalcMode: () => {
        const wb = ctx.getWb();
        const cur = wb.calcMode();
        if (cur === null) return;
        wb.setCalcMode(((cur + 1) % 3) as 0 | 1 | 2);
        handle.refresh();
      },
      onRecalc: () => {
        const wb = ctx.getWb();
        wb.recalc();
        refreshCells(ctx);
        ctx.invalidate();
      },
      onZoomChange: (zoom) => {
        setSheetZoom(ctx.store, zoom, ctx.getWb());
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

/** Quick Analysis popover (Ctrl+Q) — id `'quickAnalysis'`. */
export const quickAnalysis = (): Extension => ({
  id: 'quickAnalysis',
  priority: 50,
  setup(ctx) {
    const handle = attachQuickAnalysis({
      host: ctx.host,
      store: ctx.store,
      wb: ctx.getWb(),
      strings: ctx.i18n.strings,
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
export const charts = (): Extension => ({
  id: 'charts',
  priority: 50,
  setup(ctx) {
    const handle = attachSessionCharts({
      host: ctx.host,
      store: ctx.store,
      labels: ctx.i18n.strings.sessionCharts,
    });
    return {
      refresh: () => handle.refresh(),
      setStrings: (next) => handle.setLabels(next.sessionCharts),
      dispose: () => handle.detach(),
    };
  },
});

/** PivotTable creation dialog — id `'pivotTableDialog'`. */
export const pivotTableDialog = (): Extension => ({
  id: 'pivotTableDialog',
  priority: 50,
  setup(ctx) {
    const handle = attachPivotTableDialog({
      host: ctx.host,
      store: ctx.store,
      wb: ctx.getWb(),
      strings: ctx.i18n.strings,
      onAfterCreate: () => refreshCells(ctx),
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
export const workbookObjects = (): Extension => ({
  id: 'workbookObjects',
  priority: 50,
  setup(ctx) {
    const handle = attachWorkbookObjectsPanel({
      host: ctx.host,
      wb: ctx.getWb(),
      strings: ctx.i18n.strings,
    });
    return {
      open: () => handle.open(),
      close: () => handle.close(),
      refresh: () => handle.refresh(),
      rebindWorkbook: (wb) => handle.bindWorkbook(wb),
      setStrings: (next) => handle.setStrings(next),
      dispose: () => handle.detach(),
    };
  },
});

/** View toolbar — id `'viewToolbar'`. */
export const viewToolbar = (): Extension => ({
  id: 'viewToolbar',
  priority: 50,
  setup(ctx) {
    const objects = ctx.resolve<ExtensionHandle & { open: () => void }>('workbookObjects');
    if (ctx.viewbar.parentElement !== ctx.host) ctx.host.insertBefore(ctx.viewbar, ctx.grid);
    const handle = attachViewToolbar({
      toolbar: ctx.viewbar,
      store: ctx.store,
      wb: ctx.getWb(),
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
export const hoverComment = (): Extension => ({
  id: 'hoverComment',
  priority: 50,
  setup(ctx) {
    const handle = attachHover({ grid: ctx.grid, store: ctx.store });
    return { ...handle, dispose: handle.detach };
  },
});

/** Go To Special dialog — id `'gotoSpecial'`. */
export const goToSpecialDialog = (): Extension => ({
  id: 'gotoSpecial',
  priority: 50,
  setup(ctx) {
    let handle = attachGoToDialog({
      host: ctx.host,
      store: ctx.store,
      getWb: ctx.getWb,
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
          getWb: ctx.getWb,
          strings: next,
        });
      },
      dispose: () => handle.detach(),
    };
  },
});

/** Conditional-formatting rule manager — id `'conditional'`. */
export const conditionalDialog = (): Extension => ({
  id: 'conditional',
  priority: 50,
  setup(ctx) {
    let handle = attachConditionalDialog({
      host: ctx.host,
      store: ctx.store,
      strings: ctx.i18n.strings,
    });
    return {
      open: () => handle.open(),
      close: () => handle.close(),
      setStrings: (next) => {
        handle.detach();
        handle = attachConditionalDialog({ host: ctx.host, store: ctx.store, strings: next });
      },
      dispose: () => handle.detach(),
    };
  },
});

/** Iterative-calc settings dialog — id `'iterative'`. Always-on in the
 *  default mount; surfaced here so consumers building a custom chrome
 *  can wire their own opener. */
export const iterativeDialog = (): Extension => ({
  id: 'iterative',
  priority: 50,
  setup(ctx) {
    let handle = attachIterativeDialog({
      host: ctx.host,
      getWb: ctx.getWb,
      strings: ctx.i18n.strings,
    });
    return {
      open: () => handle.open(),
      close: () => handle.close(),
      setStrings: (next) => {
        handle.detach();
        handle = attachIterativeDialog({ host: ctx.host, getWb: ctx.getWb, strings: next });
      },
      dispose: () => handle.detach(),
    };
  },
});

/** Page Setup dialog — id `'pageSetup'`. */
export const pageSetupDialog = (): Extension => ({
  id: 'pageSetup',
  priority: 50,
  setup(ctx) {
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
export const namedRangeDialog = (): Extension => ({
  id: 'namedRanges',
  priority: 50,
  setup(ctx) {
    let handle = attachNamedRangeDialog({
      host: ctx.host,
      wb: ctx.getWb(),
      strings: ctx.i18n.strings,
    });
    return {
      open: () => handle.open(),
      close: () => handle.close(),
      rebindWorkbook: (wb) => handle.bindWorkbook(wb),
      setStrings: (next) => {
        handle.detach();
        handle = attachNamedRangeDialog({ host: ctx.host, wb: ctx.getWb(), strings: next });
      },
      dispose: () => handle.detach(),
    };
  },
});

/** Hyperlink (Ctrl+K) dialog — id `'hyperlink'`. */
export const hyperlinkDialog = (): Extension => ({
  id: 'hyperlink',
  priority: 50,
  setup(ctx) {
    const buildHandle = (s: typeof ctx.i18n.strings): ReturnType<typeof attachHyperlinkDialog> =>
      attachHyperlinkDialog({
        host: ctx.host,
        store: ctx.store,
        strings: s,
        history: ctx.history as History,
        getWb: ctx.getWb,
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

/** Comment edit (Shift+F2) dialog — id `'commentDialog'`. */
export const commentDialog = (): Extension => ({
  id: 'commentDialog',
  priority: 50,
  setup(ctx) {
    const buildHandle = (s: typeof ctx.i18n.strings): ReturnType<typeof attachCommentDialog> =>
      attachCommentDialog({
        host: ctx.host,
        store: ctx.store,
        strings: s,
        history: ctx.history as History,
        getWb: ctx.getWb,
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
export const formatDialog = (): Extension => ({
  id: 'formatDialog',
  priority: 50,
  setup(ctx) {
    const buildHandle = (s: typeof ctx.i18n.strings): ReturnType<typeof attachFormatDialog> =>
      attachFormatDialog({
        host: ctx.host,
        store: ctx.store,
        strings: s,
        history: ctx.history as History,
        getWb: ctx.getWb,
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
export const findReplace = (): Extension => ({
  id: 'findReplace',
  priority: 50,
  setup(ctx) {
    let handle = attachFindReplace({
      host: ctx.host,
      store: ctx.store,
      wb: ctx.getWb(),
      strings: ctx.i18n.strings,
      onAfterCommit: () => refreshCells(ctx),
    });
    return {
      open: () => handle.open(),
      close: () => handle.close(),
      rebindWorkbook: (wb) => {
        handle.detach();
        handle = attachFindReplace({
          host: ctx.host,
          store: ctx.store,
          wb,
          strings: ctx.i18n.strings,
          onAfterCommit: () => refreshCells(ctx),
        });
      },
      // attachFindReplace's handle exposes setStrings directly — relabels in
      // place without losing query state.
      setStrings: (next) => handle.setStrings(next),
      dispose: () => handle.detach(),
    };
  },
});

/** Validation-list dropdown — id `'validation'`. Per-workbook. */
export const validationList = (): Extension => ({
  id: 'validation',
  priority: 50,
  setup(ctx) {
    let handle = attachValidationList({
      grid: ctx.grid,
      store: ctx.store,
      wb: ctx.getWb(),
      onAfterCommit: () => refreshCells(ctx),
    });
    return {
      rebindWorkbook: (wb) => {
        handle.detach();
        handle = attachValidationList({
          grid: ctx.grid,
          store: ctx.store,
          wb,
          onAfterCommit: () => refreshCells(ctx),
        });
      },
      dispose: () => handle.detach(),
    };
  },
});

/** OS clipboard bridge — id `'clipboard'`. Per-workbook. */
export const clipboard = (): Extension => ({
  id: 'clipboard',
  priority: 50,
  setup(ctx) {
    let handle = attachClipboard({
      host: ctx.host,
      store: ctx.store,
      wb: ctx.getWb(),
      onAfterCommit: () => refreshCells(ctx),
    });
    return {
      getSnapshot: () => handle.getSnapshot(),
      rebindWorkbook: (wb) => {
        handle.detach();
        handle = attachClipboard({
          host: ctx.host,
          store: ctx.store,
          wb,
          onAfterCommit: () => refreshCells(ctx),
        });
      },
      dispose: () => handle.detach(),
    };
  },
});

/** Paste-special dialog — id `'pasteSpecial'`. Looks up the clipboard
 *  feature via `ctx.resolve('clipboard')` so a user-replacement
 *  participates. */
export const pasteSpecial = (): Extension => ({
  id: 'pasteSpecial',
  priority: 60,
  setup(ctx) {
    let activeStrings = ctx.i18n.strings;
    const buildHandle = (): ReturnType<typeof attachPasteSpecial> | null => {
      const cb = ctx.resolve<ExtensionHandle & { getSnapshot: () => unknown }>('clipboard');
      if (!cb) return null;
      return attachPasteSpecial({
        host: ctx.host,
        store: ctx.store,
        wb: ctx.getWb(),
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
export const contextMenu = (): Extension => ({
  id: 'contextMenu',
  priority: 80,
  setup(ctx) {
    const callOpen = (id: string): void => {
      const handle = ctx.resolve<ExtensionHandle & { open: () => void }>(id);
      handle?.open?.();
    };
    let detach = attachContextMenu({
      host: ctx.host,
      store: ctx.store,
      wb: ctx.getWb(),
      strings: ctx.i18n.strings,
      history: ctx.history as History,
      onAfterCommit: () => refreshCells(ctx),
      onFormatDialog: () => callOpen('formatDialog'),
      onPasteSpecial: () => callOpen('pasteSpecial'),
      onInsertHyperlink: () => callOpen('hyperlink'),
      onEditComment: () => callOpen('commentDialog'),
    });
    return {
      rebindWorkbook: (wb) => {
        detach();
        detach = attachContextMenu({
          host: ctx.host,
          store: ctx.store,
          wb,
          strings: ctx.i18n.strings,
          history: ctx.history as History,
          onAfterCommit: () => refreshCells(ctx),
          onFormatDialog: () => callOpen('formatDialog'),
          onPasteSpecial: () => callOpen('pasteSpecial'),
          onInsertHyperlink: () => callOpen('hyperlink'),
          onEditComment: () => callOpen('commentDialog'),
        });
      },
      // attachContextMenu's detacher exposes setStrings; use it directly
      // to update labels without re-attaching all listeners.
      setStrings: (next) => detach.setStrings(next),
      dispose: () => detach(),
    };
  },
});

/** Watch Window dock — id `'watchWindow'`. Default-off in `resolveFlags`. */
export const watchWindow = (): Extension => ({
  id: 'watchWindow',
  priority: 50,
  setup(ctx) {
    const dock = document.createElement('div');
    dock.dataset.fcWatch = 'dock';
    dock.className = 'fc-host__watchdock';
    ctx.host.appendChild(dock);
    const handle = attachWatchPanel({
      host: dock,
      store: ctx.store,
      getWb: ctx.getWb,
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
export const slicer = (): Extension => ({
  id: 'slicer',
  priority: 50,
  setup(ctx) {
    const handle = attachSlicer({
      host: ctx.host,
      store: ctx.store,
      getWb: ctx.getWb,
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
export const wheel = (): Extension => ({
  id: 'wheel',
  priority: 50,
  setup(ctx) {
    let detach = attachWheel({ grid: ctx.grid, store: ctx.store, wb: ctx.getWb() });
    return {
      rebindWorkbook: (wb) => {
        detach();
        detach = attachWheel({ grid: ctx.grid, store: ctx.store, wb });
      },
      dispose: () => detach(),
    };
  },
});

/** Convenience bundle — the default-on replaceable built-in factories in
 *  one array. Pair with `features: presets.minimal()` (or pick-and-mix
 *  flags) to drop the inline built-ins, then pass this array to
 *  `extensions` for the equivalent surface composed from public factories.
 *  Default-off factories (`watchWindow`, `slicer`) stay opt-in. */
export const allBuiltIns = (): Extension[] => [
  formatPainter(),
  quickAnalysis(),
  charts(),
  pivotTableDialog(),
  statusBar(),
  workbookObjects(),
  viewToolbar(),
  hoverComment(),
  goToSpecialDialog(),
  conditionalDialog(),
  iterativeDialog(),
  pageSetupDialog(),
  namedRangeDialog(),
  hyperlinkDialog(),
  commentDialog(),
  formatDialog(),
  findReplace(),
  validationList(),
  clipboard(),
  pasteSpecial(),
  contextMenu(),
  wheel(),
];
