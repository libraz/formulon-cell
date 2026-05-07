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
import { attachClipboard } from '../interact/clipboard.js';
import { attachConditionalDialog } from '../interact/conditional-dialog.js';
import { attachContextMenu } from '../interact/context-menu.js';
import { attachFindReplace } from '../interact/find-replace.js';
import { attachFormatDialog } from '../interact/format-dialog.js';
import { attachFormatPainter } from '../interact/format-painter.js';
import { attachHover } from '../interact/hover.js';
import { attachHyperlinkDialog } from '../interact/hyperlink-dialog.js';
import { attachIterativeDialog } from '../interact/iterative-dialog.js';
import { attachNamedRangeDialog } from '../interact/named-range-dialog.js';
import { attachPasteSpecial } from '../interact/paste-special.js';
import { attachStatusBar } from '../interact/status-bar.js';
import { attachValidationList } from '../interact/validation.js';
import { attachWheel } from '../interact/wheel.js';
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
    const handle = attachStatusBar({
      statusbar: ctx.statusbar,
      store: ctx.store,
      strings: ctx.i18n.strings,
      getEngineLabel: () => {
        const wb = ctx.getWb();
        return wb.isStub ? 'stub' : `formulon ${wb.version}`;
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

/** Hover-comment popover — id `'hoverComment'`. */
export const hoverComment = (): Extension => ({
  id: 'hoverComment',
  priority: 50,
  setup(ctx) {
    const handle = attachHover({ grid: ctx.grid, store: ctx.store });
    return { ...handle, dispose: handle.detach };
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
        });
      },
      // attachContextMenu's detacher exposes setStrings; use it directly
      // to update labels without re-attaching all listeners.
      setStrings: (next) => detach.setStrings(next),
      dispose: () => detach(),
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

/** Convenience bundle — every replaceable built-in factory in one
 *  array. Pair with `features: presets.minimal()` (or pick-and-mix
 *  flags) to drop the inline built-ins, then pass this array to
 *  `extensions` for the equivalent surface composed from public
 *  factories. Mostly useful for documentation / smoke tests. */
export const allBuiltIns = (): Extension[] => [
  formatPainter(),
  statusBar(),
  hoverComment(),
  conditionalDialog(),
  iterativeDialog(),
  namedRangeDialog(),
  hyperlinkDialog(),
  formatDialog(),
  findReplace(),
  validationList(),
  clipboard(),
  pasteSpecial(),
  contextMenu(),
  wheel(),
];
