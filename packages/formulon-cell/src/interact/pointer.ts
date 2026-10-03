import type { AutofitOptions } from '../commands/autofit-measurement.js';
import { fillDestFor, fillRange } from '../commands/fill.js';
import { autoFillDownExtent, restrictedFillChanges } from '../commands/fill-plan.js';
import type { History } from '../commands/history.js';
import { hyperlinkAt, isSafeHyperlinkTarget } from '../commands/hyperlinks.js';
import { interactionControllerFor } from '../commands/interaction-controller.js';
import { applyUnmerge, expandRangeWithMerges, mergeAnchorOf } from '../commands/merge.js';
import {
  collapseColGroup,
  collapseRowGroup,
  expandColGroup,
  expandRowGroup,
  isColGroupCollapsed,
  isRowGroupCollapsed,
} from '../commands/outline.js';
import { movePageBreak, resizePrintArea, setPageSetup } from '../commands/page-setup.js';
import { paginationFor } from '../commands/pagination.js';
import { autofitColsWidth, autofitRowsHeight } from '../commands/row-col-layout.js';
import {
  applyLayoutSnapshot,
  captureLayoutSnapshot,
  type LayoutSnapshot,
} from '../commands/slice-history.js';
import { MAX_COL, MAX_ROW } from '../engine/address.js';
import { syncLayoutSizesToEngine } from '../engine/layout-sync.js';
import type { Addr, Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import {
  colLeadingEdge,
  gridOriginY,
  hitTest,
  hitZone,
  layoutForView,
  rowTopEdge,
  type ViewLayout,
} from '../render/geometry.js';
import type { PageBreakHandle, RulerHandle } from '../render/grid/page-view.js';
import { getOutlineToggleHits } from '../render/grid.js';
import {
  isWholeColumnRange,
  isWholeRowRange,
  type SelectionGestureMode,
  selectionContainsAddr,
  selectionCoversRange,
} from '../store/selection-geometry.js';
import {
  mutators,
  type SelectionSlice,
  type SpreadsheetStore,
  type State,
} from '../store/store.js';
import {
  isNavigationAddrAllowed,
  navigationBoundsFor,
  navigationSelectionBoundsFor,
} from './navigation-policy.js';
import {
  indexAt,
  isFillHandleHit,
  marginInchesFor,
  pageBandAt,
  pageBreakHandleAt,
  rulerHandleAt,
  updateCursor,
} from './pointer-targets.js';
import { type RangeInsertTarget, rangeRefOf, refOf } from './range-insert.js';

type DragMode =
  | { kind: 'none' }
  | { kind: 'cell' }
  | {
      kind: 'selection-marquee';
      axis: 'cell';
      base: SelectionSlice;
      mode: SelectionGestureMode;
      start: Addr;
    }
  | {
      kind: 'selection-marquee';
      axis: 'row' | 'col';
      base: SelectionSlice;
      mode: SelectionGestureMode;
      startIndex: number;
      start: Addr;
      bounds: Range;
    }
  | { kind: 'col-header'; anchorCol: number }
  | { kind: 'row-header'; anchorRow: number }
  | {
      kind: 'col-resize';
      col: number;
      leadingEdge: number;
      rtl: boolean;
      preLayout: LayoutSnapshot;
    }
  | { kind: 'row-resize'; row: number; topEdge: number; preLayout: LayoutSnapshot }
  | { kind: 'fill'; src: Range }
  | { kind: 'page-break'; handle: PageBreakHandle; target: number }
  | { kind: 'ruler-margin'; handle: RulerHandle; inches: number }
  | {
      kind: 'range-insert';
      anchor: { row: number; col: number };
      tip: { row: number; col: number };
      base: { row: number; col: number };
      r1c1: boolean;
    };

const fullSheetRange = (sheet: number): Range => ({
  sheet,
  r0: 0,
  c0: 0,
  r1: MAX_ROW,
  c1: MAX_COL,
});

const cloneSelection = (selection: SelectionSlice): SelectionSlice => ({
  active: { ...selection.active },
  anchor: { ...selection.anchor },
  range: { ...selection.range },
  extraRanges: (selection.extraRanges ?? []).map((range) => ({ ...range })),
});

const axisSelectionRange = (
  bounds: Range,
  axis: 'row' | 'col',
  start: number,
  tip: number,
): Range =>
  axis === 'row'
    ? {
        sheet: bounds.sheet,
        r0: Math.min(start, tip),
        c0: bounds.c0,
        r1: Math.max(start, tip),
        c1: bounds.c1,
      }
    : {
        sheet: bounds.sheet,
        r0: bounds.r0,
        c0: Math.min(start, tip),
        r1: bounds.r1,
        c1: Math.max(start, tip),
      };

type SelectionMarqueeDrag = Extract<DragMode, { kind: 'selection-marquee' }>;

const geometryLayout = (state: State): ViewLayout => layoutForView(state);

const isNavigationRangeAllowed = (store: SpreadsheetStore, range: Range): boolean => {
  const bounds = navigationBoundsFor(store);
  if (!bounds) return true;
  return (
    range.sheet === bounds.sheet &&
    range.r0 >= bounds.r0 &&
    range.c0 >= bounds.c0 &&
    range.r1 <= bounds.r1 &&
    range.c1 <= bounds.c1
  );
};

export interface PointerDeps {
  store: SpreadsheetStore;
  wb: WorkbookHandle;
  /** Refresh cached cells after a write — same contract as the inline editor. */
  onAfterCommit?: () => void;
  /** Shared history. When provided, col/row resizes and fill drags push one
   *  entry per drag-end (not per intermediate frame). */
  history?: History | null;
}

export function attachPointer(
  host: HTMLElement,
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  onAfterCommit?: () => void,
  history: History | null = null,
  getEditor: () => RangeInsertTarget | null = () => null,
  /** Locale and theme font that header double-click autofit measures with. */
  getAutofitOptions: () => AutofitOptions = () => ({}),
): () => void {
  let drag: DragMode = { kind: 'none' };
  const unsubscribeSheetChange = store.subscribe((state) => {
    if (drag.kind === 'selection-marquee' && state.data.sheetIndex !== drag.base.range.sheet) {
      drag = { kind: 'none' };
    }
  });

  const localXY = (e: PointerEvent | MouseEvent): { x: number; y: number } => {
    const rect = host.getBoundingClientRect();
    return { x: e.clientX - rect.left, y: e.clientY - rect.top };
  };

  const applySelectionMarquee = (marquee: SelectionMarqueeDrag, x: number, y: number): void => {
    const s = store.getState();
    if (s.data.sheetIndex !== marquee.base.range.sheet) {
      drag = { kind: 'none' };
      return;
    }
    const zone = hitZone(geometryLayout(s), s.viewport, x, y, s.ui.filterRange);
    if (!zone) return;

    if (marquee.axis === 'cell') {
      if (zone.kind !== 'cell') return;
      const tip = mergeAnchorOf(s, { sheet: s.data.sheetIndex, row: zone.row, col: zone.col });
      const requested: Range = {
        sheet: marquee.start.sheet,
        r0: Math.min(marquee.start.row, tip.row),
        c0: Math.min(marquee.start.col, tip.col),
        r1: Math.max(marquee.start.row, tip.row),
        c1: Math.max(marquee.start.col, tip.col),
      };
      mutators.applySelectionRectangle(
        store,
        marquee.base,
        expandRangeWithMerges(s, requested),
        marquee.mode,
        marquee.start,
        tip,
      );
      return;
    }

    if (marquee.axis === 'row') {
      if (zone.kind !== 'row-header' && zone.kind !== 'row-resize') return;
      const requested = axisSelectionRange(marquee.bounds, 'row', marquee.startIndex, zone.row);
      const tip = mergeAnchorOf(s, {
        sheet: marquee.bounds.sheet,
        row: zone.row,
        col: marquee.bounds.c0,
      });
      mutators.applySelectionRectangle(
        store,
        marquee.base,
        expandRangeWithMerges(s, requested),
        marquee.mode,
        marquee.start,
        tip,
      );
      return;
    }

    if (zone.kind !== 'col-header' && zone.kind !== 'col-resize') return;
    const requested = axisSelectionRange(marquee.bounds, 'col', marquee.startIndex, zone.col);
    const tip = mergeAnchorOf(s, {
      sheet: marquee.bounds.sheet,
      row: marquee.bounds.r0,
      col: zone.col,
    });
    mutators.applySelectionRectangle(
      store,
      marquee.base,
      expandRangeWithMerges(s, requested),
      marquee.mode,
      marquee.start,
      tip,
    );
  };

  /** Grow the selection to `addr`, then widen it over any merges it now cuts. */
  const extendSelectionTo = (addr: Addr): void => {
    mutators.extendRangeTo(store, addr);
    const after = store.getState();
    const grown = expandRangeWithMerges(after, after.selection.range);
    if (
      grown.r0 !== after.selection.range.r0 ||
      grown.r1 !== after.selection.range.r1 ||
      grown.c0 !== after.selection.range.c0 ||
      grown.c1 !== after.selection.range.c1
    ) {
      mutators.setRange(store, grown);
    }
  };

  /** Ctrl/Cmd-click on a row or column header: begin a marquee that adds the
   *  band to the selection, or subtracts it when it is already covered. */
  const startHeaderMarquee = (
    s: State,
    axis: 'row' | 'col',
    index: number,
    x: number,
    y: number,
  ): void => {
    const base = cloneSelection(s.selection);
    const bounds = navigationSelectionBoundsFor(store) ?? fullSheetRange(s.data.sheetIndex);
    const initial = axisSelectionRange(bounds, axis, index, index);
    const start = mergeAnchorOf(
      s,
      axis === 'row'
        ? { sheet: bounds.sheet, row: index, col: bounds.c0 }
        : { sheet: bounds.sheet, row: bounds.r0, col: index },
    );
    const marquee: SelectionMarqueeDrag = {
      kind: 'selection-marquee',
      axis,
      base,
      mode: selectionCoversRange(base, initial) ? 'subtract' : 'add',
      startIndex: index,
      start,
      bounds: { ...bounds },
    };
    drag = marquee;
    applySelectionMarquee(marquee, x, y);
  };

  /** Write a fill of `src` into `dest` and promote `dest` to the selection.
   *  Restricted mounts authorize the whole destination as one batch through the
   *  interaction controller; others write straight to the workbook in one undo
   *  step. Merges the destination cuts are unmerged first. */
  const commitFill = (
    s: State,
    src: Range,
    dest: Range,
    opts: { copyOnly: boolean },
  ): { applied: boolean; restricted: boolean } => {
    const controller = interactionControllerFor(store);
    const restricted = controller?.policy !== undefined;
    let applied: boolean;
    if (controller?.policy) {
      const result = controller.execute({
        type: 'cellBatch',
        operation: 'fill',
        origin: 'fillHandle',
        changes: restrictedFillChanges(s, src, dest, opts.copyOnly),
        denied: controller.policy.batchDenied,
      });
      applied = result.status === 'applied';
    } else {
      // Bundle every per-cell write into a single undoable transaction.
      if (history) history.begin();
      applied = false;
      try {
        // Fill cannot tear merged rectangles apart silently.
        applyUnmerge(store, wb, history, dest);
        // Re-read the state: the unmerge above changed the merge maps the fill checks.
        applied = fillRange(store.getState(), wb, src, dest, {
          copyOnly: opts.copyOnly,
          formatting: 'with',
          store,
        });
      } finally {
        if (history) history.end();
      }
    }
    if (applied) {
      onAfterCommit?.();
      mutators.setActive(store, { sheet: dest.sheet, row: dest.r0, col: dest.c0 });
      mutators.extendRangeTo(store, { sheet: dest.sheet, row: dest.r1, col: dest.c1 });
    }
    return { applied, restricted };
  };

  const onDown = (e: PointerEvent): void => {
    if (e.button !== 0) return;
    // Clicks that land on the inline editor itself are not grid clicks — let
    //  the editor handle them natively. Without this, clicking inside an open
    //  formula editor would hit-test to the editing cell and insert that
    //  cell's own ref into the formula.
    if (e.target instanceof Element && e.target.closest('.fc-host__editor')) return;
    const { x, y } = localXY(e);
    const s = store.getState();
    const controller = interactionControllerFor(store);
    const restricted = controller?.policy !== undefined;
    const selectionDisabled = controller?.policy?.selection === false;

    // Capture editor intent BEFORE we touch focus — host.focus() blurs the
    //  textarea, which triggers commit + cancel and tears down the editor
    //  before our cell-zone branch could query it.
    const editor = getEditor();
    const inFormula = editor?.isFormulaEdit();

    // A selection-disabled mount may still use cell clicks while a formula
    // editor is open to insert references. All other pointer entry points
    // must leave the host's fixed active selection untouched.
    if (selectionDisabled && !inFormula) {
      e.preventDefault();
      if (s.ui.editor.kind !== 'idle') host.focus({ preventScroll: true });
      drag = { kind: 'none' };
      return;
    }

    // setPointerCapture throws on synthetic events / certain pointer-id mismatches.
    // Wrap to avoid crashing the handler; the worst-case fallback is that move
    // events stop firing once the pointer leaves the host.
    const tryCapture = (): void => {
      try {
        host.setPointerCapture(e.pointerId);
      } catch {
        /* no-op: still works without capture for in-host drags */
      }
    };

    // Outline toggles in the bracket gutters take precedence over hit-zoning;
    // they sit in territory the hitZone fall-through would otherwise treat as
    // a header/corner click.
    for (const t of getOutlineToggleHits()) {
      if (x < t.rect.x || x > t.rect.x + t.rect.w) continue;
      if (y < t.rect.y || y > t.rect.y + t.rect.h) continue;
      if (restricted) {
        e.preventDefault();
        drag = { kind: 'none' };
        return;
      }
      e.preventDefault();
      if (t.axis === 'row') {
        const collapsed = isRowGroupCollapsed(s.layout, t.i0, t.i1);
        if (collapsed) expandRowGroup(store, history, t.i0, t.i1);
        else collapseRowGroup(store, history, t.i0, t.i1);
      } else {
        const collapsed = isColGroupCollapsed(s.layout, t.i0, t.i1);
        if (collapsed) expandColGroup(store, history, t.i0, t.i1);
        else collapseColGroup(store, history, t.i0, t.i1);
      }
      drag = { kind: 'none' };
      onAfterCommit?.();
      return;
    }

    // Page Layout: dragging a ruler's margin boundary sets that margin, the
    // same affordance the desktop app puts on its rulers.
    const ruler = rulerHandleAt(s, x, y);
    if (ruler) {
      if (restricted) {
        e.preventDefault();
        return;
      }
      e.preventDefault();
      host.focus();
      tryCapture();
      drag = { kind: 'ruler-margin', handle: ruler, inches: marginInchesFor(ruler, x, y) };
      host.style.cursor =
        ruler.side === 'left' || ruler.side === 'right' ? 'col-resize' : 'row-resize';
      return;
    }

    // Page Layout: the header / footer bands live in the page margins, which
    // no cell occupies, so a click there can only mean "edit that slot".
    const band = pageBandAt(s, x, y);
    if (band) {
      if (restricted) {
        e.preventDefault();
        return;
      }
      e.preventDefault();
      host.dispatchEvent(
        new CustomEvent('fc:editpageband', { bubbles: true, detail: { ...band } }),
      );
      drag = { kind: 'none' };
      return;
    }

    // Page Break Preview: the blue boundaries sit on top of the cells and are
    // the only thing in that view a drag from this position could mean.
    const breakHandle = pageBreakHandleAt(s, x, y);
    if (breakHandle) {
      if (restricted) {
        e.preventDefault();
        return;
      }
      e.preventDefault();
      host.focus();
      tryCapture();
      drag = { kind: 'page-break', handle: breakHandle, target: breakHandle.index };
      mutators.setPageBreakDrag(store, {
        axis: breakHandle.axis,
        position: breakHandle.position,
      });
      host.style.cursor = breakHandle.axis === 'row' ? 'row-resize' : 'col-resize';
      return;
    }

    // Fill handle takes precedence over normal cell hit-testing.
    if (isFillHandleHit(x, y)) {
      if (selectionDisabled) {
        e.preventDefault();
        drag = { kind: 'none' };
        return;
      }
      if (restricted) {
        if (!controller) return;
        const permission = controller.canExecute({
          operation: 'fill',
          origin: 'fillHandle',
          effects: [{ kind: 'range', range: s.selection.range }],
        });
        if (!permission.allowed) {
          e.preventDefault();
          return;
        }
      }
      host.focus();
      tryCapture();
      drag = { kind: 'fill', src: { ...s.selection.range } };
      mutators.setFillPreview(store, { ...s.selection.range });
      host.style.cursor = 'crosshair';
      return;
    }

    const layout = geometryLayout(s);
    const zone = hitZone(layout, s.viewport, x, y, s.ui.filterRange);
    if (!zone) return;

    // Range-insert: keep focus on the editor (preventDefault avoids the
    //  pointerdown's default focus-change behavior).
    if (inFormula && zone.kind === 'cell' && editor) {
      e.preventDefault();
      tryCapture();
      const anchor = { row: zone.row, col: zone.col };
      const base = { row: s.selection.active.row, col: s.selection.active.col };
      const r1c1 = s.ui.r1c1 === true;
      editor.insertRefAtCaret(refOf(anchor.row, anchor.col, { r1c1, base }));
      drag = { kind: 'range-insert', anchor, tip: anchor, base, r1c1 };
      return;
    }

    host.focus();
    tryCapture();

    switch (zone.kind) {
      case 'corner':
        mutators.selectAll(store);
        drag = { kind: 'none' };
        return;

      case 'col-header': {
        if (e.shiftKey) {
          const anchorCol = isWholeColumnRange(s.selection.range)
            ? s.selection.anchor.col
            : s.selection.active.col;
          mutators.selectCols(store, anchorCol, zone.col);
          drag = { kind: 'col-header', anchorCol };
          return;
        }
        if (e.ctrlKey || e.metaKey) {
          startHeaderMarquee(s, 'col', zone.col, x, y);
          return;
        }
        mutators.selectCol(store, zone.col);
        drag = { kind: 'col-header', anchorCol: zone.col };
        return;
      }

      case 'col-filter-btn': {
        // Reflect the spreadsheet convention: clicking the chevron does not move the active cell.
        // Bubble out a CustomEvent so the chrome layer (mount.ts) can open
        // the filter dropdown anchored under the chevron.
        e.preventDefault();
        const fr = s.ui.filterRange;
        if (fr) {
          const rect = host.getBoundingClientRect();
          host.dispatchEvent(
            new CustomEvent('fc:openfilter', {
              bubbles: true,
              detail: {
                range: fr,
                col: zone.col,
                anchor: {
                  x: e.clientX - rect.left,
                  y: gridOriginY(geometryLayout(s)) - s.layout.headerRowHeight,
                  h: s.layout.headerRowHeight,
                  clientX: e.clientX,
                  clientY: e.clientY,
                },
              },
            }),
          );
        }
        drag = { kind: 'none' };
        return;
      }

      case 'row-header': {
        if (e.shiftKey) {
          const anchorRow = isWholeRowRange(s.selection.range)
            ? s.selection.anchor.row
            : s.selection.active.row;
          mutators.selectRows(store, anchorRow, zone.row);
          drag = { kind: 'row-header', anchorRow };
          return;
        }
        if (e.ctrlKey || e.metaKey) {
          startHeaderMarquee(s, 'row', zone.row, x, y);
          return;
        }
        mutators.selectRow(store, zone.row);
        drag = { kind: 'row-header', anchorRow: zone.row };
        return;
      }

      case 'col-resize': {
        if (restricted) {
          e.preventDefault();
          drag = { kind: 'none' };
          return;
        }
        drag = {
          kind: 'col-resize',
          col: zone.col,
          // The column's leading edge stays put while the trailing edge
          // follows the pointer. On a right-to-left sheet that leading edge
          // is physically on the right, so the drag runs the other way.
          leadingEdge: colLeadingEdge(geometryLayout(s), s.viewport, zone.col),
          rtl: s.ui.rightToLeft === true,
          preLayout: captureLayoutSnapshot(s),
        };
        return;
      }

      case 'row-resize': {
        if (restricted) {
          e.preventDefault();
          drag = { kind: 'none' };
          return;
        }
        const topEdge = rowTopEdge(geometryLayout(s), s.viewport, zone.row);
        drag = {
          kind: 'row-resize',
          row: zone.row,
          topEdge,
          preLayout: captureLayoutSnapshot(s),
        };
        return;
      }

      case 'cell': {
        const rawAddr = { sheet: s.data.sheetIndex, row: zone.row, col: zone.col };
        // Click on a merged cell body — promote to the merge anchor (spreadsheet parity).
        const addr = mergeAnchorOf(s, rawAddr);
        if (!isNavigationAddrAllowed(store, addr)) {
          drag = { kind: 'none' };
          return;
        }
        const meta = (e.ctrlKey || e.metaKey) && !e.shiftKey;
        // Ctrl/Cmd+click on a hyperlinked cell follows the link (spreadsheet parity).
        // Falls through to multi-range selection when the cell has no link so
        // the modifier stays useful for non-link cells.
        if (meta) {
          const url = hyperlinkAt(s, rawAddr);
          if (url) {
            e.preventDefault();
            openHyperlink(url);
            mutators.setActive(store, addr);
            drag = { kind: 'none' };
            return;
          }
          const base = cloneSelection(s.selection);
          const marquee: SelectionMarqueeDrag = {
            kind: 'selection-marquee',
            axis: 'cell',
            base,
            mode: selectionContainsAddr(base, rawAddr) ? 'subtract' : 'add',
            start: addr,
          };
          drag = marquee;
          applySelectionMarquee(marquee, x, y);
          return;
        }
        if (e.shiftKey) {
          extendSelectionTo(addr);
        } else mutators.setActive(store, addr);
        drag = { kind: 'cell' };
        return;
      }
    }
  };

  const onMove = (e: PointerEvent): void => {
    const { x, y } = localXY(e);

    if (
      interactionControllerFor(store)?.policy?.selection === false &&
      drag.kind !== 'none' &&
      drag.kind !== 'range-insert'
    ) {
      drag = { kind: 'none' };
      mutators.setFillPreview(store, null);
      return;
    }

    if (drag.kind === 'none') {
      updateCursor(host, store, x, y);
      return;
    }

    const s = store.getState();

    switch (drag.kind) {
      case 'selection-marquee': {
        applySelectionMarquee(drag, x, y);
        return;
      }
      case 'col-resize': {
        const w = drag.rtl ? drag.leadingEdge - x : x - drag.leadingEdge;
        mutators.setColWidth(store, drag.col, w);
        host.style.cursor = 'col-resize';
        return;
      }
      case 'row-resize': {
        const h = y - drag.topEdge;
        mutators.setRowHeight(store, drag.row, h);
        host.style.cursor = 'row-resize';
        return;
      }
      case 'col-header': {
        const layout = geometryLayout(s);
        const zone = hitZone(layout, s.viewport, x, y);
        if (zone && (zone.kind === 'col-header' || zone.kind === 'col-resize')) {
          mutators.selectCols(store, drag.anchorCol, zone.col);
        }
        return;
      }
      case 'row-header': {
        const layout = geometryLayout(s);
        const zone = hitZone(layout, s.viewport, x, y);
        if (zone && (zone.kind === 'row-header' || zone.kind === 'row-resize')) {
          mutators.selectRows(store, drag.anchorRow, zone.row);
        }
        return;
      }
      case 'cell': {
        const layout = geometryLayout(s);
        const zone = hitZone(layout, s.viewport, x, y);
        if (zone && zone.kind === 'cell') {
          extendSelectionTo({ sheet: s.data.sheetIndex, row: zone.row, col: zone.col });
        }
        return;
      }
      case 'fill': {
        const cell = hitTest(geometryLayout(s), s.viewport, x, y);
        if (!cell) return;
        const dest = fillDestFor(drag.src, { row: cell.row, col: cell.col });
        if (!isNavigationRangeAllowed(store, dest)) {
          mutators.setFillPreview(store, null);
          return;
        }
        mutators.setFillPreview(store, dest);
        host.style.cursor = 'crosshair';
        return;
      }
      case 'ruler-margin': {
        const horizontal = drag.handle.side === 'left' || drag.handle.side === 'right';
        host.style.cursor = horizontal ? 'col-resize' : 'row-resize';
        drag.inches = marginInchesFor(drag.handle, x, y);
        return;
      }
      case 'page-break': {
        const axis = drag.handle.axis;
        host.style.cursor = axis === 'row' ? 'row-resize' : 'col-resize';
        mutators.setPageBreakDrag(store, { axis, position: axis === 'row' ? y : x });
        const at = indexAt(s, x, y);
        if (at) drag.target = axis === 'row' ? at.row : at.col;
        return;
      }
      case 'range-insert': {
        const cell = hitTest(geometryLayout(s), s.viewport, x, y);
        if (!cell) return;
        if (cell.row === drag.tip.row && cell.col === drag.tip.col) return;
        drag.tip = { row: cell.row, col: cell.col };
        const editor = getEditor();
        if (editor) {
          editor.insertRefAtCaret(
            rangeRefOf(drag.anchor, drag.tip, { r1c1: drag.r1c1, base: drag.base }),
          );
        }
        return;
      }
    }
  };

  /**
   * Apply a released page-boundary drag.
   *
   * A break dropped where the automatic pagination would have put it anyway is
   * removed rather than pinned, so dragging a line back to its original place
   * undoes the override instead of freezing it. Dropping a break before the
   * page it opens is meaningless, so that removes it too — which is how the
   * desktop preview merges two pages.
   */
  const commitPageBreakDrag = (handle: PageBreakHandle, target: number): void => {
    const s = store.getState();
    const sheet = s.data.sheetIndex;
    if (handle.kind === 'printArea') {
      const pagination = paginationFor(s, sheet);
      const origin = handle.axis === 'row' ? pagination.origin.row : pagination.origin.col;
      if (target < origin) return;
      resizePrintArea(store, sheet, handle.axis, target, pagination.content, history);
      return;
    }
    if (target === handle.index) return;
    movePageBreak(store, sheet, handle.axis, handle.index, target, history);
  };

  const onUp = (e: PointerEvent): void => {
    if (host.hasPointerCapture(e.pointerId)) host.releasePointerCapture(e.pointerId);
    if (
      interactionControllerFor(store)?.policy?.selection === false &&
      drag.kind !== 'none' &&
      drag.kind !== 'range-insert'
    ) {
      drag = { kind: 'none' };
      host.style.cursor = '';
      mutators.setPageBreakDrag(store, null);
      mutators.setFillPreview(store, null);
      return;
    }
    if (
      interactionControllerFor(store)?.policy !== undefined &&
      (drag.kind === 'col-resize' ||
        drag.kind === 'row-resize' ||
        drag.kind === 'ruler-margin' ||
        drag.kind === 'page-break')
    ) {
      drag = { kind: 'none' };
      host.style.cursor = '';
      mutators.setPageBreakDrag(store, null);
      mutators.setFillPreview(store, null);
      return;
    }
    if (drag.kind === 'selection-marquee') {
      const { x, y } = localXY(e);
      applySelectionMarquee(drag, x, y);
      drag = { kind: 'none' };
      updateCursor(host, store, x, y);
      return;
    }
    if (drag.kind === 'col-resize' || drag.kind === 'row-resize') {
      // One undo entry per drag, not per pixel: capture pre at drag-start and
      // post here, push the closure pair. Engine-side sync rides on the same
      // snapshot pair so undo/redo replays the resize in the workbook too.
      const before = drag.preLayout;
      const after = captureLayoutSnapshot(store.getState());
      const sheet = store.getState().data.sheetIndex;
      syncLayoutSizesToEngine(wb, store.getState().layout, sheet, before, after);
      if (history && !history.isReplaying()) {
        history.push({
          undo: () => {
            applyLayoutSnapshot(store, before);
            syncLayoutSizesToEngine(wb, store.getState().layout, sheet, after, before);
          },
          redo: () => {
            applyLayoutSnapshot(store, after);
            syncLayoutSizesToEngine(wb, store.getState().layout, sheet, before, after);
          },
        });
      }
    }
    if (drag.kind === 'ruler-margin') {
      const { handle, inches } = drag;
      drag = { kind: 'none' };
      host.style.cursor = '';
      const sheet = store.getState().data.sheetIndex;
      setPageSetup(store, sheet, { margins: { [handle.side]: inches } }, history);
      onAfterCommit?.();
      return;
    }
    if (drag.kind === 'page-break') {
      const { handle, target } = drag;
      drag = { kind: 'none' };
      mutators.setPageBreakDrag(store, null);
      host.style.cursor = '';
      commitPageBreakDrag(handle, target);
      onAfterCommit?.();
      return;
    }
    if (drag.kind === 'fill') {
      const s = store.getState();
      const dest = s.ui.fillPreview;
      mutators.setFillPreview(store, null);
      if (dest) {
        if (!isNavigationRangeAllowed(store, dest)) {
          drag = { kind: 'none' };
          const { x, y } = localXY(e);
          updateCursor(host, store, x, y);
          return;
        }
        // Spreadsheet parity: holding Ctrl/⌘ on release toggles series → tile copy.
        const copyOnly = e.ctrlKey || e.metaKey;
        const { applied, restricted } = commitFill(s, drag.src, dest, { copyOnly });
        if (applied && !restricted) {
          host.dispatchEvent(
            new CustomEvent('fc:autofilloptions', {
              bubbles: true,
              detail: {
                src: drag.src,
                dest,
                mode: copyOnly ? 'copy' : 'series',
                clientX: e.clientX,
                clientY: e.clientY,
              },
            }),
          );
        }
      }
    }
    drag = { kind: 'none' };
    const { x, y } = localXY(e);
    updateCursor(host, store, x, y);
  };

  /** An interrupted gesture is abandoned: live previews revert and nothing is committed. */
  const onCancel = (e: PointerEvent): void => {
    if (host.hasPointerCapture(e.pointerId)) host.releasePointerCapture(e.pointerId);
    if (drag.kind === 'selection-marquee') {
      const base = drag.base;
      store.setState((s) => ({ ...s, selection: base }));
    } else if (drag.kind === 'col-resize' || drag.kind === 'row-resize') {
      applyLayoutSnapshot(store, drag.preLayout);
    }
    drag = { kind: 'none' };
    mutators.setPageBreakDrag(store, null);
    mutators.setFillPreview(store, null);
    host.style.cursor = '';
  };

  const onLeave = (): void => {
    if (drag.kind === 'none') host.style.cursor = '';
  };

  const onDblClick = (e: MouseEvent): void => {
    const { x, y } = localXY(e);
    const s = store.getState();
    if (interactionControllerFor(store)?.policy?.selection === false) {
      e.preventDefault();
      e.stopPropagation();
      return;
    }

    // Fill-handle takes precedence — spreadsheet-style "double-click to flash-fill
    // down to match the neighbour column's contiguous run."
    if (isFillHandleHit(x, y)) {
      e.preventDefault();
      e.stopPropagation();
      const src = { ...s.selection.range };
      const dest = autoFillDownExtent(s, src);
      if (!dest) return;
      if (!isNavigationRangeAllowed(store, dest)) return;
      commitFill(s, src, dest, { copyOnly: false });
      return;
    }

    const zone = hitZone(geometryLayout(s), s.viewport, x, y);
    if (!zone) return;

    if (zone.kind === 'col-resize') {
      e.preventDefault();
      e.stopPropagation();
      if (interactionControllerFor(store)?.policy !== undefined) return;
      autofitColsWidth(store, history, zone.col, zone.col, wb, getAutofitOptions());
      return;
    }
    if (zone.kind === 'row-resize') {
      e.preventDefault();
      e.stopPropagation();
      if (interactionControllerFor(store)?.policy !== undefined) return;
      autofitRowsHeight(store, history, zone.row, zone.row, wb, getAutofitOptions());
      return;
    }
  };

  host.addEventListener('pointerdown', onDown);
  host.addEventListener('pointermove', onMove);
  host.addEventListener('pointerup', onUp);
  host.addEventListener('pointercancel', onCancel);
  host.addEventListener('pointerleave', onLeave);
  host.addEventListener('dblclick', onDblClick);

  return () => {
    unsubscribeSheetChange();
    host.removeEventListener('pointerdown', onDown);
    host.removeEventListener('pointermove', onMove);
    host.removeEventListener('pointerup', onUp);
    host.removeEventListener('pointercancel', onCancel);
    host.removeEventListener('pointerleave', onLeave);
    host.removeEventListener('dblclick', onDblClick);
    host.style.cursor = '';
  };
}

/**
 * Opens a hyperlink in a new tab. Restricted to safe protocols (http(s)://,
 * mailto:, tel:) so a hostile cell value can't smuggle a `javascript:` URL.
 */
function openHyperlink(url: string): void {
  const trimmed = url.trim();
  if (trimmed.length === 0) return;
  if (!isSafeHyperlinkTarget(trimmed)) return;
  if (typeof window === 'undefined' || typeof window.open !== 'function') return;
  window.open(trimmed, '_blank', 'noopener,noreferrer');
}
