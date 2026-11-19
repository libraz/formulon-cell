import type { AutofitOptions } from '../commands/autofit-measurement.js';
import { clearSelectedContents } from '../commands/clear-contents.js';
import { inputLocaleForWorkbook } from '../commands/coerce-input.js';
import { isPopulated } from '../commands/current-region.js';
import { userCustomLists } from '../commands/custom-lists.js';
import {
  executeFillHandle,
  fillDestFor,
  fillLastValue,
  fillOutputRange,
  fillShrinkDropped,
  fillShrinkFor,
} from '../commands/fill.js';
import { autoFillDownExtent } from '../commands/fill-plan.js';
import { headerFilterRanges, hiddenFilterButtonTest } from '../commands/filter-ranges.js';
import { tableOverlayById } from '../commands/format-as-table.js';
import type { History } from '../commands/history.js';
import { hyperlinkAt, isInDocumentHyperlinkTarget } from '../commands/hyperlinks.js';
import { interactionControllerFor } from '../commands/interaction-controller.js';
import { fullFidelityCellWrites, policyPermits } from '../commands/interaction-policy.js';
import { expandRangeWithMerges, mergeAnchorOf } from '../commands/merge.js';
import { moveSelectionTo, moveTargetFits, moveTargetRange } from '../commands/move-cells.js';
import {
  collapseColGroup,
  collapseRowGroup,
  expandColGroup,
  expandRowGroup,
  isColGroupCollapsed,
  isRowGroupCollapsed,
  showOutlineLevel,
} from '../commands/outline.js';
import { movePageBreak, resizePrintArea, setPageSetup } from '../commands/page-setup.js';
import { pageSetupWrite } from '../commands/page-setup-actions.js';
import { paginationFor } from '../commands/pagination.js';
import { refreshPivotReportsAfterEdit } from '../commands/pivot-report-update.js';
import { pivotTableAtRange } from '../commands/pivot-table.js';
import { blockedByProtection, deferForAllowedRangeUnlock } from '../commands/protection.js';
import { layoutAllowed, layoutWrite } from '../commands/row-col-actions.js';
import {
  autofitHeaderTracks,
  headerGestureSpans,
  previewHeaderResize,
  type TrackSpan,
} from '../commands/row-col-layout.js';
import {
  applyLayoutSnapshot,
  captureLayoutSnapshot,
  type LayoutSnapshot,
} from '../commands/slice-history.js';
import { formatA1Range, MAX_COL, MAX_ROW } from '../engine/address.js';
import { syncLayoutSizesToEngine } from '../engine/layout-sync.js';
import type { Addr, Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { GridPopupsStrings } from '../i18n/strings/grid-popups/types.js';
import { defaultStrings } from '../i18n/strings.js';
import {
  cellRect,
  colWidth,
  frozenColsWidth,
  frozenRowsHeight,
  gridOriginX,
  gridOriginY,
  hitTest,
  hitZone,
  layoutForView,
  type Rect,
  rowHeight,
} from '../render/geometry.js';
import type { PageBreakHandle, RulerHandle } from '../render/grid/page-view.js';
import { gridHitState } from '../render/grid.js';
import { filterButtonRect } from '../render/painters/filter-button.js';
import { renderScale } from '../store/render-scale.js';
import { isRowHidden } from '../store/row-visibility.js';
import {
  axisSelectionRange,
  cloneSelection,
  fullSheetRange,
  isWholeColumnRange,
  isWholeRowRange,
  type SelectionGestureMode,
  sameRange,
  selectionContainsAddr,
  selectionCoversRange,
} from '../store/selection-geometry.js';
import {
  mutators,
  type SelectionSlice,
  type SpreadsheetStore,
  type State,
} from '../store/store.js';
import { firstVisibleCol, firstVisibleRow } from '../store/visible-axis.js';
import { jumpToDataEdge } from './data-edge-navigation.js';
import { type AutoscrollBounds, createDragAutoscroll } from './drag-autoscroll.js';
import { createFillValueTooltip } from './fill-value-tooltip.js';
import { resizeCursor } from './grid-cursors.js';
import { followHyperlink, type HyperlinkFollowDeps } from './hyperlink-follow.js';
import {
  isNavigationAddrAllowed,
  isNavigationRangeAllowed,
  navigationSelectionBoundsFor,
  scrollActiveIntoView,
} from './navigation-policy.js';
import { pivotRowItemAt, togglePivotItemDetail } from './pivot-detail.js';
import { movePivotItem, type PivotItemSlots, pivotItemSlotsAt } from './pivot-item-order.js';
import { movePivotTable } from './pivot-move.js';
import { capturePointerSafely, releaseHeldPointer } from './pointer-capture.js';
import {
  indexAt,
  isFillHandleHit,
  isSelectionBorderHit,
  marginInchesFor,
  pageBandAt,
  pageBreakHandleAt,
  pivotButtonAt,
  type RefBoxCorner,
  refBoxHitAt,
  rulerHandleAt,
  selectionBorderEdgeAt,
  updateCursor,
} from './pointer-targets.js';
import { type RangeInsertTarget, rangeRefOf, refOf } from './range-insert.js';
import { createResizeSizeTooltip } from './resize-size-tooltip.js';
import {
  applyTableResize,
  resizableTableGripAt,
  TABLE_RESIZE_COMMAND,
  tableResizeRange,
} from './table-resize.js';

type ResizeTipStrings = GridPopupsStrings['resizeTip'];

type DragMode =
  | { kind: 'none' }
  | { kind: 'cell'; link?: ClickedLink }
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
      /** Tracks the gesture resizes: the whole selected columns, or just `col`. */
      spans: TrackSpan[];
      /** Pointer x at press and the track's screen width then. */
      startPointer: number;
      startSize: number;
      rtl: boolean;
      preLayout: LayoutSnapshot;
    }
  | {
      kind: 'row-resize';
      row: number;
      spans: TrackSpan[];
      startPointer: number;
      startSize: number;
      preLayout: LayoutSnapshot;
    }
  | { kind: 'fill'; src: Range; menu?: boolean }
  | { kind: 'table-resize'; id: string }
  /** Moving or resizing a formula-edit reference box; `press` is the cell grabbed. */
  | {
      kind: 'ref-box';
      index: number;
      part: 'edge' | 'corner';
      corner?: RefBoxCorner;
      orig: Range;
      press: { row: number; col: number };
      base: { row: number; col: number };
      r1c1: boolean;
      started: boolean;
    }
  /** Moving the selection by its frame; `grab` is the grabbed cell's offset
   *  from the block's top-left. */
  | {
      kind: 'move';
      src: Range;
      grab: { row: number; col: number };
      menu?: boolean;
      /** Client position of a right-button press, to tell a click from a drag. */
      pressed?: { x: number; y: number };
    }
  /** An item of a report dragged by its cell border; `after` is the row it is dropped below. */
  | { kind: 'pivot-item'; slots: PivotItemSlots; fromRow: number; after: number }
  | { kind: 'page-break'; handle: PageBreakHandle; target: number }
  | { kind: 'ruler-margin'; handle: RulerHandle; inches: number }
  | {
      kind: 'range-insert';
      anchor: { row: number; col: number };
      tip: { row: number; col: number };
      base: { row: number; col: number };
      r1c1: boolean;
    };

type SelectionMarqueeDrag = Extract<DragMode, { kind: 'selection-marquee' }>;

const lastVisibleCol = (s: State): number =>
  Math.min(
    MAX_COL,
    s.viewport.colStart + s.viewport.colCount - 1,
    s.viewport.navigationRange?.c1 ?? MAX_COL,
  );

const lastVisibleRow = (s: State): number =>
  Math.min(
    MAX_ROW,
    s.viewport.rowStart + s.viewport.rowCount - 1,
    s.viewport.navigationRange?.r1 ?? MAX_ROW,
  );

export interface PointerDeps {
  store: SpreadsheetStore;
  wb: WorkbookHandle;
  /** Refresh cached cells after a write — same contract as the inline editor. */
  onAfterCommit?: () => void;
  /** Shared history. When provided, col/row resizes and fill drags push one
   *  entry per drag-end (not per intermediate frame). */
  history?: History | null;
}

/** A press on a link cell held at least this long selects instead of following (measured 0.5-0.6 s). */
const HYPERLINK_HOLD_SELECT_MS = 500;

/** A link cell pressed without a modifier, followed on release. */
interface ClickedLink {
  url: string;
  range: Range;
  at: number;
}

/** The cell at the window's top-left corner; frozen panes pin it to the sheet origin. */
function topLeftVisibleCell(s: State): { row: number; col: number } {
  return {
    row: s.layout.freezeRows > 0 ? 0 : s.viewport.rowStart,
    col: s.layout.freezeCols > 0 ? 0 : s.viewport.colStart,
  };
}

/** The range a reference-box drag produces with the pointer over `cell`: an edge
 *  drag shifts the box by the cell delta from the press, a corner drag keeps the
 *  opposite corner fixed. */
const refBoxRangeFor = (
  drag: Extract<DragMode, { kind: 'ref-box' }>,
  cell: { row: number; col: number },
): Range => {
  const { orig } = drag;
  if (drag.part === 'edge') {
    const dRow = Math.max(-orig.r0, Math.min(MAX_ROW - orig.r1, cell.row - drag.press.row));
    const dCol = Math.max(-orig.c0, Math.min(MAX_COL - orig.c1, cell.col - drag.press.col));
    return {
      sheet: orig.sheet,
      r0: orig.r0 + dRow,
      c0: orig.c0 + dCol,
      r1: orig.r1 + dRow,
      c1: orig.c1 + dCol,
    };
  }
  const fixedRow = drag.corner === 'nw' || drag.corner === 'ne' ? orig.r1 : orig.r0;
  const fixedCol = drag.corner === 'nw' || drag.corner === 'sw' ? orig.c1 : orig.c0;
  return {
    sheet: orig.sheet,
    r0: Math.min(fixedRow, cell.row),
    c0: Math.min(fixedCol, cell.col),
    r1: Math.max(fixedRow, cell.row),
    c1: Math.max(fixedCol, cell.col),
  };
};

export function attachPointer(
  host: HTMLElement,
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  onAfterCommit?: () => void,
  history: History | null = null,
  getEditor: () => RangeInsertTarget | null = () => null,
  /** Locale and theme font that header double-click autofit measures with. */
  getAutofitOptions: () => AutofitOptions = () => ({}),
  /** Sheet activation and invalid-target reporting for in-document links. */
  hyperlinkFollow: Pick<HyperlinkFollowDeps, 'activateSheet' | 'onInvalidTarget'> = {},
  /** Asked before a drag-move overwrites data; a host without one overwrites,
   *  as the desktop does with its alerts off. */
  confirmMoveOverwrite: () => Promise<boolean> = async () => true,
  /** Templates for the live size readout of a header-border drag. */
  getResizeTipStrings: () => ResizeTipStrings = () => defaultStrings.resizeTip,
): () => void {
  const hits = gridHitState(host);
  const fillTooltip = createFillValueTooltip(host);
  const resizeTooltip = createResizeSizeTooltip(host);
  let drag: DragMode = { kind: 'none' };
  const unsubscribeSheetChange = store.subscribe((state) => {
    if (drag.kind === 'selection-marquee' && state.data.sheetIndex !== drag.base.range.sheet) {
      drag = { kind: 'none' };
    }
  });

  // A header border drag or double-click sizes the tracks as Row Height / Column Width do.
  const headerResizeContext = (axis: 'row' | 'col') =>
    ({
      store,
      history,
      origin: 'editor',
      commandId: axis === 'row' ? 'rowHeight' : 'colWidth',
    }) as const;

  const localXY = (
    e: PointerEvent | MouseEvent,
    rect: DOMRect = host.getBoundingClientRect(),
  ): { x: number; y: number } => ({ x: e.clientX - rect.left, y: e.clientY - rect.top });

  /** Client coordinates of a cell rect's trailing corner, for anchoring popups to the pointer's frame. */
  const clientAnchorOf = (
    e: PointerEvent | MouseEvent,
    rect: Rect,
    at: { x: number; y: number } = localXY(e),
  ): { x: number; y: number } => ({
    x: e.clientX - at.x + rect.x + rect.w,
    y: e.clientY - at.y + rect.y + rect.h,
  });

  const refreshCursor = (e: PointerEvent | MouseEvent): void => {
    const { x, y } = localXY(e);
    updateCursor(host, store, x, y);
  };

  const applySelectionMarquee = (marquee: SelectionMarqueeDrag, x: number, y: number): void => {
    const s = store.getState();
    if (s.data.sheetIndex !== marquee.base.range.sheet) {
      drag = { kind: 'none' };
      return;
    }
    const zone = hitZone(layoutForView(s), s.viewport, x, y, headerFilterRanges(s), {
      filterButtonHidden: hiddenFilterButtonTest(s),
    });
    if (!zone) return;

    let requested: Range;
    let tip: Addr;
    if (marquee.axis === 'cell') {
      if (zone.kind !== 'cell') return;
      tip = mergeAnchorOf(s, { sheet: s.data.sheetIndex, row: zone.row, col: zone.col });
      requested = {
        sheet: marquee.start.sheet,
        r0: Math.min(marquee.start.row, tip.row),
        c0: Math.min(marquee.start.col, tip.col),
        r1: Math.max(marquee.start.row, tip.row),
        c1: Math.max(marquee.start.col, tip.col),
      };
    } else if (marquee.axis === 'row') {
      if (zone.kind !== 'row-header' && zone.kind !== 'row-resize') return;
      requested = axisSelectionRange(marquee.bounds, 'row', marquee.startIndex, zone.row);
      tip = mergeAnchorOf(s, {
        sheet: marquee.bounds.sheet,
        row: zone.row,
        col: marquee.bounds.c0,
      });
    } else {
      if (zone.kind !== 'col-header' && zone.kind !== 'col-resize') return;
      requested = axisSelectionRange(marquee.bounds, 'col', marquee.startIndex, zone.col);
      tip = mergeAnchorOf(s, {
        sheet: marquee.bounds.sheet,
        row: marquee.bounds.r0,
        col: zone.col,
      });
    }
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
    if (!sameRange(grown, after.selection.range)) mutators.setRange(store, grown);
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
    src: Range,
    dest: Range,
    opts: { copyOnly: boolean; announce?: { clientX: number; clientY: number } },
  ): { applied: boolean; restricted: boolean } => {
    const restricted = hostPolicy() !== undefined;
    // A password-guarded range in the way is unlocked first; the fill then repeats.
    if (
      !restricted &&
      deferForAllowedRangeUnlock(store, [fillOutputRange(src, dest)], () =>
        commitFill(src, dest, opts),
      )
    ) {
      return { applied: false, restricted };
    }
    const applied =
      executeFillHandle({ store, workbook: wb, history, src, dest, copyOnly: opts.copyOnly })
        .status === 'applied';
    if (applied) {
      onAfterCommit?.();
      mutators.setActive(store, { sheet: dest.sheet, row: dest.r0, col: dest.c0 });
      mutators.extendRangeTo(store, { sheet: dest.sheet, row: dest.r1, col: dest.c1 });
      // Raised here, not by the caller, so a fill that waited for a password gets its options too.
      if (!restricted && opts.announce) {
        host.dispatchEvent(
          new CustomEvent('fc:autofilloptions', {
            bubbles: true,
            detail: {
              src,
              dest,
              mode: opts.copyOnly ? 'copy' : 'series',
              clientX: opts.announce.clientX,
              clientY: opts.announce.clientY,
            },
          }),
        );
      }
    }
    return { applied, restricted };
  };

  /** Pulling the fill handle back inside the source clears the dropped cells'
   *  contents (formats stay) and shrinks the selection to what is kept. */
  const commitFillShrink = (s: State, src: Range, kept: Range): void => {
    const dropped = fillShrinkDropped(src, kept);
    const active = s.selection.active;
    mutators.setActive(store, { sheet: dropped.sheet, row: dropped.r0, col: dropped.c0 });
    mutators.extendRangeTo(store, { sheet: dropped.sheet, row: dropped.r1, col: dropped.c1 });
    const result = clearSelectedContents({
      store,
      workbook: wb,
      history,
      origin: 'fillHandle',
      commandId: 'fill-shrink',
    });
    mutators.setActive(store, active);
    mutators.setRange(store, kept);
    if (result.status === 'applied') onAfterCommit?.();
  };

  /** Pointer travel, in client px, within which a right press is still a click. */
  const RIGHT_CLICK_SLOP_PX = 4;

  let releaseContextMenuSwallow: (() => void) | null = null;
  const swallowContextMenu = (): void => {
    releaseContextMenuSwallow?.();
    const doc = host.ownerDocument;
    const swallow = (ev: Event): void => {
      ev.preventDefault();
      ev.stopPropagation();
    };
    doc.addEventListener('contextmenu', swallow, true);
    releaseContextMenuSwallow = () => {
      doc.removeEventListener('contextmenu', swallow, true);
      releaseContextMenuSwallow = null;
    };
  };

  // A drag starts only when the policy allows its operation; the cells and the sheet's
  // protection are decided when it drops.
  const hostPolicy = () => interactionControllerFor(store)?.policy;
  const selectionDisabled = (): boolean => hostPolicy()?.selection === false;
  /** A border drag moves whole cells, formats and formulas included. */
  const moveDragAllowed = (): boolean => fullFidelityCellWrites(hostPolicy());
  /** A right-button fill drag ends in a menu whose fills carry formats and formulas. */
  const fillMenuDragAllowed = (): boolean =>
    fullFidelityCellWrites(hostPolicy()) && policyPermits(hostPolicy(), 'fill');
  const pageDragAllowed = (): boolean => policyPermits(hostPolicy(), 'pageSetup');

  /** Take focus and pointer capture for a gesture the host owns. */
  const grab = (e: PointerEvent): void => {
    host.focus();
    capturePointerSafely(host, e.pointerId);
  };

  /** A secondary press moves the selection before the context menu opens:
   *  a cell outside every selected range becomes the selection, the corner
   *  selects the sheet. Headers are resolved by the menu itself. */
  const onSecondaryDown = (e: PointerEvent): void => {
    const s = store.getState();
    if (selectionDisabled()) return;
    if (s.ui.editor.kind !== 'idle') return;
    const { x, y } = localXY(e);
    if (isFillHandleHit(hits, x, y) && fillMenuDragAllowed()) {
      // A right-button fill drag ends in a menu of fill kinds, so the context
      // menu the same press would raise is swallowed.
      grab(e);
      drag = { kind: 'fill', src: { ...s.selection.range }, menu: true };
      mutators.setFillPreview(store, { ...s.selection.range });
      host.style.cursor = 'crosshair';
      swallowContextMenu();
      return;
    }
    // A right-button drag by the selection frame ends in a menu of drop actions.
    if (moveDragAllowed() && isSelectionBorderHit(s, x, y)) {
      const grabbed = hitTest(layoutForView(s), s.viewport, x, y);
      if (grabbed) {
        grab(e);
        const src = { ...s.selection.range };
        drag = {
          kind: 'move',
          src,
          grab: { row: grabbed.row - src.r0, col: grabbed.col - src.c0 },
          menu: true,
          pressed: { x: e.clientX, y: e.clientY },
        };
        host.style.cursor = 'move';
        swallowContextMenu();
        return;
      }
    }
    const zone = hitZone(layoutForView(s), s.viewport, x, y, [], { resizeHandles: false });
    if (!zone) return;
    if (zone.kind === 'corner') {
      mutators.selectAll(store, topLeftVisibleCell(s));
      return;
    }
    if (zone.kind !== 'cell') return;
    const addr = { sheet: s.data.sheetIndex, row: zone.row, col: zone.col };
    if (selectionContainsAddr(s.selection, addr)) return;
    if (!isNavigationAddrAllowed(store, addr)) return;
    mutators.setActive(store, addr);
  };

  const onDown = (e: PointerEvent): void => {
    if (e.button === 2) {
      onSecondaryDown(e);
      return;
    }
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

    // Capture editor intent BEFORE we touch focus — host.focus() blurs the
    //  textarea, which triggers commit + cancel and tears down the editor
    //  before our cell-zone branch could query it.
    const editor = getEditor();
    const inFormula = editor?.isFormulaEdit();

    // A selection-disabled mount may still use cell clicks while a formula
    // editor is open to insert references. All other pointer entry points
    // must leave the host's fixed active selection untouched.
    if (selectionDisabled() && !inFormula) {
      e.preventDefault();
      if (s.ui.editor.kind !== 'idle') host.focus({ preventScroll: true });
      drag = { kind: 'none' };
      return;
    }

    // Outline toggles and level buttons in the gutters take precedence over hit-zoning;
    // they sit in territory the hitZone fall-through would otherwise treat as
    // a header/corner click.
    for (const t of hits.outlineToggles) {
      if (x < t.rect.x || x > t.rect.x + t.rect.w) continue;
      if (y < t.rect.y || y > t.rect.y + t.rect.h) continue;
      e.preventDefault();
      drag = { kind: 'none' };
      // A gutter toggle expands or collapses the outline, as Show/Hide Detail does.
      const collapsed =
        t.kind !== 'level' &&
        (t.axis === 'row'
          ? isRowGroupCollapsed(s.layout, t.i0, t.i1)
          : isColGroupCollapsed(s.layout, t.i0, t.i1));
      const ctx = {
        store,
        history,
        origin: 'editor',
        commandId: t.kind === 'level' || collapsed ? 'outlineShowDetail' : 'outlineHideDetail',
      } as const;
      const toggled = layoutWrite(ctx, t.axis, false, () => {
        if (t.kind === 'level') showOutlineLevel(store, history, t.axis, t.level, wb);
        else if (t.axis === 'row') {
          if (collapsed) expandRowGroup(store, history, t.i0, t.i1, wb);
          else collapseRowGroup(store, history, t.i0, t.i1, wb);
        } else if (collapsed) expandColGroup(store, history, t.i0, t.i1, wb);
        else collapseColGroup(store, history, t.i0, t.i1, wb);
        return true;
      });
      if (toggled) onAfterCommit?.();
      return;
    }

    // Page Layout: dragging a ruler's margin boundary sets that margin, the
    // same affordance the desktop app puts on its rulers.
    const ruler = rulerHandleAt(s, hits, x, y);
    if (ruler) {
      e.preventDefault();
      if (!pageDragAllowed()) return;
      grab(e);
      drag = { kind: 'ruler-margin', handle: ruler, inches: marginInchesFor(ruler, x, y) };
      host.style.cursor =
        ruler.side === 'left' || ruler.side === 'right' ? 'col-resize' : 'row-resize';
      return;
    }

    // Page Layout: the header / footer bands live in the page margins, which
    // no cell occupies, so a click there can only mean "edit that slot".
    const band = pageBandAt(s, hits, x, y);
    if (band) {
      e.preventDefault();
      if (!pageDragAllowed()) return;
      host.dispatchEvent(
        new CustomEvent('fc:editpageband', { bubbles: true, detail: { ...band } }),
      );
      drag = { kind: 'none' };
      return;
    }

    // Page Break Preview: the blue boundaries sit on top of the cells and are
    // the only thing in that view a drag from this position could mean.
    const breakHandle = pageBreakHandleAt(s, hits, x, y);
    if (breakHandle) {
      e.preventDefault();
      if (!pageDragAllowed()) return;
      grab(e);
      drag = { kind: 'page-break', handle: breakHandle, target: breakHandle.index };
      mutators.setPageBreakDrag(store, {
        axis: breakHandle.axis,
        position: breakHandle.position,
      });
      host.style.cursor = breakHandle.axis === 'row' ? 'row-resize' : 'col-resize';
      return;
    }

    // Fill handle takes precedence over normal cell hit-testing.
    if (isFillHandleHit(hits, x, y)) {
      if (selectionDisabled()) {
        e.preventDefault();
        drag = { kind: 'none' };
        return;
      }
      if (restricted) {
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
      grab(e);
      drag = { kind: 'fill', src: { ...s.selection.range } };
      mutators.setFillPreview(store, { ...s.selection.range });
      host.style.cursor = 'crosshair';
      return;
    }

    // A table's corner grip resizes the table; the target range previews in the
    // same band the fill drag uses.
    if (!inFormula) {
      const table = resizableTableGripAt(store, x, y);
      if (table) {
        e.preventDefault();
        grab(e);
        drag = { kind: 'table-resize', id: table.id };
        mutators.setFillPreview(store, { ...table.range });
        host.style.cursor = 'nwse-resize';
        return;
      }
    }

    // The selection frame moves the block. The cell under the pointer is the
    // grab point, so a grab on the bottom edge counts from the row below it.
    if (moveDragAllowed() && !inFormula && isSelectionBorderHit(s, x, y)) {
      const grabbed = hitTest(layoutForView(s), s.viewport, x, y);
      const { r0, c0, r1, c1 } = s.selection.range;
      // An item of a report is reordered, never moved as a cell.
      const itemSlots =
        r0 === r1 && c0 === c1 ? pivotItemSlotsAt(s, wb, s.data.sheetIndex, r0, c0) : null;
      if (itemSlots) {
        e.preventDefault();
        grab(e);
        drag = { kind: 'pivot-item', slots: itemSlots, fromRow: r0, after: r0 };
        return;
      }
      if (grabbed) {
        e.preventDefault();
        grab(e);
        const src = { ...s.selection.range };
        drag = {
          kind: 'move',
          src,
          grab: { row: grabbed.row - src.r0, col: grabbed.col - src.c0 },
        };
        host.style.cursor = 'move';
        return;
      }
    }

    // A reference box of the formula being edited moves by its edge and
    // resizes by its corner handles.
    if (inFormula && editor?.replaceReferenceAt) {
      const box = refBoxHitAt(s, x, y);
      const ref = box ? s.ui.editorRefs[box.index] : undefined;
      const pressed = box ? hitTest(layoutForView(s), s.viewport, x, y) : null;
      if (box && ref && pressed) {
        e.preventDefault();
        capturePointerSafely(host, e.pointerId);
        const orig = { sheet: s.data.sheetIndex, r0: ref.r0, c0: ref.c0, r1: ref.r1, c1: ref.c1 };
        drag = {
          kind: 'ref-box',
          index: box.index,
          part: box.part,
          corner: box.corner,
          orig,
          press: { row: pressed.row, col: pressed.col },
          base: { ...editor.editedCell() },
          r1c1: s.ui.r1c1 === true,
          started: false,
        };
        return;
      }
    }

    const layout = layoutForView(s);
    const zone = hitZone(layout, s.viewport, x, y, headerFilterRanges(s), {
      filterButtonHidden: hiddenFilterButtonTest(s),
    });
    if (!zone) return;

    // Range-insert: keep focus on the editor (preventDefault avoids the
    //  pointerdown's default focus-change behavior).
    if (inFormula && zone.kind === 'cell' && editor) {
      e.preventDefault();
      capturePointerSafely(host, e.pointerId);
      const rawAnchor = { row: zone.row, col: zone.col };
      const mergedAnchor = mergeAnchorOf(s, {
        sheet: s.data.sheetIndex,
        ...rawAnchor,
      });
      const anchor = { row: mergedAnchor.row, col: mergedAnchor.col };
      const { row, col } = editor.editedCell();
      const base = { row, col };
      const r1c1 = s.ui.r1c1 === true;
      editor.insertRefAtCaret(refOf(anchor.row, anchor.col, { r1c1, base }), { phase: 'start' });
      drag = { kind: 'range-insert', anchor, tip: anchor, base, r1c1 };
      return;
    }

    grab(e);

    switch (zone.kind) {
      case 'corner':
        mutators.selectAll(store, topLeftVisibleCell(s));
        drag = { kind: 'none' };
        return;

      case 'col-header': {
        if (e.shiftKey) {
          const anchorCol = isWholeColumnRange(s.selection.range)
            ? s.selection.anchor.col
            : s.selection.active.col;
          mutators.selectCols(store, anchorCol, zone.col, s.selection.active.row);
          drag = { kind: 'col-header', anchorCol };
          return;
        }
        if (e.ctrlKey || e.metaKey) {
          startHeaderMarquee(s, 'col', zone.col, x, y);
          return;
        }
        // The active cell lands on the column's top visible row; the view does not move.
        mutators.selectCol(store, zone.col, firstVisibleRow(s));
        drag = { kind: 'col-header', anchorCol: zone.col };
        return;
      }

      case 'col-filter-btn': {
        // Reflect the spreadsheet convention: clicking the chevron does not move the active cell.
        // Bubble out a CustomEvent so the chrome layer (mount.ts) can open
        // the filter dropdown anchored under the chevron.
        e.preventDefault();
        const rect = host.getBoundingClientRect();
        // The dropdown hangs from the bottom of the in-cell button.
        const button = filterButtonRect(
          cellRect(layoutForView(s), s.viewport, zone.range.r0, zone.col),
          s.ui.rightToLeft === true,
        );
        host.dispatchEvent(
          new CustomEvent('fc:openfilter', {
            bubbles: true,
            detail: {
              range: zone.range,
              col: zone.col,
              anchor: {
                x: e.clientX - rect.left,
                y: button.y,
                h: button.h,
                clientX: e.clientX,
                clientY: e.clientY,
              },
            },
          }),
        );
        drag = { kind: 'none' };
        return;
      }

      case 'row-header': {
        if (e.shiftKey) {
          const anchorRow = isWholeRowRange(s.selection.range)
            ? s.selection.anchor.row
            : s.selection.active.row;
          mutators.selectRows(store, anchorRow, zone.row, s.selection.active.col);
          drag = { kind: 'row-header', anchorRow };
          return;
        }
        if (e.ctrlKey || e.metaKey) {
          startHeaderMarquee(s, 'row', zone.row, x, y);
          return;
        }
        // The active cell lands on the row's leftmost visible column; the view does not move.
        mutators.selectRow(store, zone.row, firstVisibleCol(s));
        drag = { kind: 'row-header', anchorRow: zone.row };
        return;
      }

      case 'col-resize': {
        // A protected sheet refuses the drag unless it allows formatting columns.
        if (
          !layoutAllowed(headerResizeContext('col'), 'col') ||
          blockedByProtection(store, s.data.sheetIndex, 'col-resize', 'formatColumns')
        ) {
          e.preventDefault();
          drag = { kind: 'none' };
          return;
        }
        drag = {
          kind: 'col-resize',
          col: zone.col,
          spans: headerGestureSpans(s, 'col', zone.col),
          startPointer: x,
          startSize: colWidth(layout, zone.col, s.viewport),
          rtl: s.ui.rightToLeft === true,
          preLayout: captureLayoutSnapshot(s),
        };
        return;
      }

      case 'row-resize': {
        if (
          !layoutAllowed(headerResizeContext('row'), 'row') ||
          blockedByProtection(store, s.data.sheetIndex, 'row-resize', 'formatRows')
        ) {
          e.preventDefault();
          drag = { kind: 'none' };
          return;
        }
        drag = {
          kind: 'row-resize',
          row: zone.row,
          spans: headerGestureSpans(s, 'row', zone.row),
          startPointer: y,
          startSize: rowHeight(layout, zone.row, s.viewport),
          preLayout: captureLayoutSnapshot(s),
        };
        return;
      }

      case 'cell': {
        // The Mac desktop treats the selection border as a handle: a plain
        // click on it leaves the selection, so a double-click can jump from it.
        if (
          !e.shiftKey &&
          !e.metaKey &&
          !e.ctrlKey &&
          selectionBorderEdgeAt(host, s, x, y) !== null
        ) {
          drag = { kind: 'none' };
          return;
        }
        const rawAddr = { sheet: s.data.sheetIndex, row: zone.row, col: zone.col };
        const pivotButton = pivotButtonAt(s, wb, host, zone.row, zone.col, x, y);
        // A report's drop-down button opens its filter panel; the active cell stays.
        if (pivotButton?.kind === 'field') {
          e.preventDefault();
          const rect = host.getBoundingClientRect();
          const { rect: button } = pivotButton;
          host.dispatchEvent(
            new CustomEvent('fc:openpivotfilter', {
              bubbles: true,
              detail: {
                addr: rawAddr,
                anchor: { x: rect.left + button.x, y: rect.top + button.y + button.h },
              },
            }),
          );
          drag = { kind: 'none' };
          return;
        }
        // A row label's expand/collapse button turns its item over; the active cell stays.
        if (pivotButton?.kind === 'outline') {
          e.preventDefault();
          const ref = pivotRowItemAt(wb, rawAddr);
          if (ref) {
            togglePivotItemDetail(
              {
                wb,
                history,
                store,
                origin: 'editor',
                commandId: ref.showDetail ? 'pivotCollapse' : 'pivotExpand',
                onChanged: () => refreshPivotReportsAfterEdit(store, wb),
              },
              ref,
            );
            onAfterCommit?.();
          }
          drag = { kind: 'none' };
          return;
        }
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
            const followed = followHyperlink({ store, wb, ...hyperlinkFollow }, url);
            if (!followed || !isInDocumentHyperlinkTarget(url)) mutators.setActive(store, addr);
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
          drag = { kind: 'cell' };
          return;
        }
        mutators.setActive(store, addr);
        const url = e.altKey ? null : hyperlinkAt(s, rawAddr);
        drag = url
          ? {
              kind: 'cell',
              link: { url, range: store.getState().selection.range, at: e.timeStamp },
            }
          : { kind: 'cell' };
        return;
      }
    }
  };

  /** Re-apply the active drag at a point, as a pointer move there would. */
  let applyDragAt: (x: number, y: number) => void = () => {};

  const autoscroll = createDragAutoscroll({
    scrollBy: (dRow, dCol) => mutators.scrollBy(store, dRow, dCol),
    onScrolled: ({ x, y }) => {
      const point = autoscrollTarget(x, y);
      if (point) applyDragAt(point.x, point.y);
      else autoscroll.stop();
    },
  });

  /** Directions the active drag may scroll in, or null when it never autoscrolls. */
  const autoscrollAxes = (): { row: boolean; col: boolean } | null => {
    if (drag.kind === 'cell') return { row: true, col: true };
    if (drag.kind === 'col-header') return { row: false, col: true };
    if (drag.kind === 'row-header') return { row: true, col: false };
    if (drag.kind === 'selection-marquee') {
      return {
        row: drag.axis !== 'col',
        col: drag.axis !== 'row',
      };
    }
    return null;
  };

  /** Body area the pointer must stay inside to avoid scrolling, in host coordinates. */
  const autoscrollBounds = (
    rect: DOMRect = host.getBoundingClientRect(),
  ): AutoscrollBounds | null => {
    if (rect.width <= 0 || rect.height <= 0) return null;
    const s = store.getState();
    const layout = layoutForView(s);
    const originX = gridOriginX(layout) + frozenColsWidth(layout, s.viewport);
    return {
      left: layout.rtl ? 0 : originX,
      right: layout.rtl ? rect.width - originX : rect.width,
      top: gridOriginY(layout) + frozenRowsHeight(layout, s.viewport),
      bottom: rect.height,
    };
  };

  /** The pointer pulled back into the grid body (header strip on the axis a header drag keeps). */
  const autoscrollTarget = (x: number, y: number): { x: number; y: number } | null => {
    const axes = autoscrollAxes();
    const bounds = autoscrollBounds();
    if (!axes || !bounds) return null;
    const s = store.getState();
    const layout = layoutForView(s);
    const clampX = Math.min(bounds.right - 1, Math.max(bounds.left, x));
    const clampY = Math.min(bounds.bottom - 1, Math.max(bounds.top, y));
    return {
      x: axes.col ? clampX : gridOriginX(layout) - 1,
      y: axes.row ? clampY : gridOriginY(layout) - 1,
    };
  };

  const updateAutoscroll = (x: number, y: number, rect: DOMRect): void => {
    const axes = autoscrollAxes();
    const bounds = autoscrollBounds(rect);
    if (!axes || !bounds) {
      autoscroll.stop();
      return;
    }
    autoscroll.update({ x, y }, bounds, axes);
  };

  const onMove = (e: PointerEvent): void => {
    const rect = host.getBoundingClientRect();
    const { x, y } = localXY(e, rect);

    if (selectionDisabled() && drag.kind !== 'none' && drag.kind !== 'range-insert') {
      drag = { kind: 'none' };
      mutators.setFillPreview(store, null);
      return;
    }

    if (drag.kind === 'none') {
      autoscroll.stop();
      updateCursor(host, store, x, y);
      return;
    }

    updateAutoscroll(x, y, rect);
    applyDrag(e, x, y);
  };

  /** Show the dragged track's current size beside the pointer. */
  const showResizeTip = (axis: 'col' | 'row', track: number, e: PointerEvent | null): void => {
    if (!e) return;
    const s = store.getState();
    const strings = getResizeTipStrings();
    const hidden = axis === 'col' ? s.layout.hiddenCols.has(track) : isRowHidden(s.layout, track);
    const px = hidden
      ? 0
      : axis === 'col'
        ? (s.layout.colWidths.get(track) ?? s.layout.defaultColWidth)
        : (s.layout.rowHeights.get(track) ?? s.layout.defaultRowHeight);
    const size = hidden
      ? 0
      : axis === 'col'
        ? wb.columnDisplayChars(s.data.sheetIndex, px)
        : wb.rowDisplayPt(px);
    const template =
      axis === 'col'
        ? hidden
          ? strings.widthHidden
          : strings.width
        : hidden
          ? strings.heightHidden
          : strings.height;
    resizeTooltip.show(
      template
        .replace('{size}', size.toFixed(2))
        .replace('{pixels}', String(Math.round(px * renderScale(s.layout)))),
      e.clientX,
      e.clientY,
    );
  };

  // Header and cell drags rewrite the selection only when the pointer moves to
  // another target while the selection is as the last write left it.
  let lastDragWrite: {
    drag: DragMode;
    key: string;
    selection: State['selection'];
    merges: State['merges'];
  } | null = null;
  const applyOncePerTarget = (s: State, key: string, write: () => void): void => {
    const last = lastDragWrite;
    if (
      last &&
      last.drag === drag &&
      last.key === key &&
      last.selection === s.selection &&
      last.merges === s.merges
    ) {
      return;
    }
    write();
    const after = store.getState();
    lastDragWrite = { drag, key, selection: after.selection, merges: after.merges };
  };

  // The tooltip value only changes with the destination, not with every move.
  let fillValueCache: {
    drag: DragMode;
    dest: Range;
    data: State['data'];
    format: State['format'];
    value: string | null;
  } | null = null;
  const fillValueFor = (s: State, owner: DragMode, src: Range, dest: Range): string | null => {
    const cached = fillValueCache;
    if (
      cached &&
      cached.drag === owner &&
      cached.data === s.data &&
      cached.format === s.format &&
      sameRange(cached.dest, dest)
    ) {
      return cached.value;
    }
    const value = fillLastValue(s, src, dest, inputLocaleForWorkbook(wb), userCustomLists(store));
    fillValueCache = { drag: owner, dest: { ...dest }, data: s.data, format: s.format, value };
    return value;
  };

  const applyDrag = (e: PointerEvent | null, x: number, y: number): void => {
    const s = store.getState();

    switch (drag.kind) {
      case 'selection-marquee': {
        applySelectionMarquee(drag, x, y);
        return;
      }
      case 'col-resize': {
        // Screen px → canonical px; the layout is rebuilt from the pre-drag
        // snapshot on every move, and one history entry lands on pointerup.
        // The size follows the pointer 1:1 from where the border was grabbed;
        // on a right-to-left sheet the trailing edge runs the other way.
        const w = Math.round(
          drag.startSize + (drag.rtl ? drag.startPointer - x : x - drag.startPointer),
        );
        const scale = renderScale(s.layout, s.viewport);
        previewHeaderResize(
          store,
          'col',
          drag.preLayout,
          drag.spans,
          w / scale,
          Math.max(0, -w) / scale,
        );
        host.style.cursor = resizeCursor(host, 'col');
        showResizeTip('col', drag.col, e);
        return;
      }
      case 'row-resize': {
        const h = Math.round(drag.startSize + y - drag.startPointer);
        const scale = renderScale(s.layout, s.viewport);
        previewHeaderResize(
          store,
          'row',
          drag.preLayout,
          drag.spans,
          h / scale,
          Math.max(0, -h) / scale,
        );
        host.style.cursor = resizeCursor(host, 'row');
        showResizeTip('row', drag.row, e);
        return;
      }
      case 'col-header': {
        const layout = layoutForView(s);
        const zone = hitZone(layout, s.viewport, x, y, [], { resizeHandles: false });
        const end =
          zone && (zone.kind === 'col-header' || zone.kind === 'cell')
            ? zone.col
            : zone === null
              ? lastVisibleCol(s)
              : null;
        if (end !== null) {
          const anchorCol = drag.anchorCol;
          applyOncePerTarget(s, `col:${end}`, () =>
            mutators.selectCols(store, anchorCol, end, s.selection.active.row),
          );
        }
        return;
      }
      case 'row-header': {
        const layout = layoutForView(s);
        const zone = hitZone(layout, s.viewport, x, y, [], { resizeHandles: false });
        const end =
          zone && (zone.kind === 'row-header' || zone.kind === 'cell')
            ? zone.row
            : zone === null
              ? lastVisibleRow(s)
              : null;
        if (end !== null) {
          const anchorRow = drag.anchorRow;
          applyOncePerTarget(s, `row:${end}`, () =>
            mutators.selectRows(store, anchorRow, end, s.selection.active.col),
          );
        }
        return;
      }
      case 'cell': {
        const layout = layoutForView(s);
        const zone = hitZone(layout, s.viewport, x, y);
        // Over a header strip the drag reaches the nearest visible edge cell.
        const edge = topLeftVisibleCell(s);
        const sheet = s.data.sheetIndex;
        let target: Addr | null = null;
        if (zone && zone.kind === 'cell') {
          target = { sheet, row: zone.row, col: zone.col };
        } else if (zone && (zone.kind === 'row-header' || zone.kind === 'row-resize')) {
          target = { sheet, row: zone.row, col: edge.col };
        } else if (zone && (zone.kind === 'col-header' || zone.kind === 'col-resize')) {
          target = { sheet, row: edge.row, col: zone.col };
        } else if (zone?.kind === 'corner') {
          target = { sheet, row: edge.row, col: edge.col };
        }
        if (target) {
          const to = target;
          applyOncePerTarget(s, `${to.sheet}:${to.row}:${to.col}`, () => extendSelectionTo(to));
        }
        return;
      }
      case 'fill': {
        const layout = layoutForView(s);
        const cell = hitTest(layout, s.viewport, x, y);
        if (!cell) return;
        const corner = cellRect(layout, s.viewport, drag.src.r1, drag.src.c1);
        const kept = drag.menu
          ? null
          : fillShrinkFor(
              drag.src,
              { row: cell.row, col: cell.col },
              { dx: x - (corner.x + corner.w), dy: y - (corner.y + corner.h) },
            );
        // Pulling back inside the source drops its trailing cells; no value tip.
        if (kept) {
          mutators.setFillPreview(store, kept);
          host.style.cursor = 'crosshair';
          fillTooltip.hide();
          return;
        }
        const dest = fillDestFor(drag.src, { row: cell.row, col: cell.col });
        if (!isNavigationRangeAllowed(store, dest)) {
          mutators.setFillPreview(store, null);
          return;
        }
        mutators.setFillPreview(store, dest);
        host.style.cursor = 'crosshair';
        const last = fillValueFor(s, drag, drag.src, dest);
        if (last === null || !e) {
          fillTooltip.hide();
        } else {
          const edge = cellRect(layout, s.viewport, dest.r1, dest.c1);
          fillTooltip.show(last, clientAnchorOf(e, edge, { x, y }).x, e.clientY);
        }
        return;
      }
      case 'table-resize': {
        const cell = hitTest(layoutForView(s), s.viewport, x, y);
        const table = tableOverlayById(s, drag.id);
        if (!cell || !table) return;
        mutators.setFillPreview(store, tableResizeRange(s, table, cell));
        host.style.cursor = 'nwse-resize';
        return;
      }
      case 'move': {
        const cell = hitTest(layoutForView(s), s.viewport, x, y);
        if (!cell) return;
        const topLeft = {
          row: Math.max(0, cell.row - drag.grab.row),
          col: Math.max(0, cell.col - drag.grab.col),
        };
        const target = moveTargetRange(drag.src, topLeft);
        const moved = target.r0 !== drag.src.r0 || target.c0 !== drag.src.c0;
        mutators.setMovePreview(store, moved && moveTargetFits(target) ? target : null);
        host.style.cursor = 'move';
        return;
      }
      case 'pivot-item': {
        // The pointer rides the dragged item's bottom border, one pixel inside its row.
        const cell = hitTest(layoutForView(s), s.viewport, x, y - 1);
        if (!cell) return;
        const { slots } = drag;
        drag.after = Math.min(Math.max(cell.row, slots.firstRow - 1), slots.lastRow);
        const band = {
          sheet: slots.sheet,
          r0: drag.after,
          c0: slots.col,
          r1: drag.after,
          c1: slots.c1,
        };
        mutators.setMovePreview(store, band);
        if (!e) return;
        const edge = cellRect(layoutForView(s), s.viewport, band.r1, band.c1);
        fillTooltip.show(formatA1Range(band), clientAnchorOf(e, edge, { x, y }).x, e.clientY);
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
      case 'ref-box': {
        const cell = hitTest(layoutForView(s), s.viewport, x, y);
        const editor = getEditor();
        if (!cell || !editor?.replaceReferenceAt) return;
        const next = refBoxRangeFor(drag, cell);
        editor.replaceReferenceAt(
          drag.index,
          rangeRefOf(
            { row: next.r0, col: next.c0 },
            { row: next.r1, col: next.c1 },
            { r1c1: drag.r1c1, base: drag.base },
          ),
          { phase: drag.started ? 'update' : 'start' },
        );
        drag.started = true;
        return;
      }
      case 'range-insert': {
        const cell = hitTest(layoutForView(s), s.viewport, x, y);
        if (!cell) return;
        const mergedCell = mergeAnchorOf(s, { sheet: s.data.sheetIndex, ...cell });
        const tip = { row: mergedCell.row, col: mergedCell.col };
        if (tip.row === drag.tip.row && tip.col === drag.tip.col) return;
        drag.tip = tip;
        const editor = getEditor();
        if (editor) {
          editor.insertRefAtCaret(
            rangeRefOf(drag.anchor, drag.tip, { r1c1: drag.r1c1, base: drag.base }),
            { phase: 'update' },
          );
        }
        return;
      }
    }
  };

  /** A page-setup change made by dragging on the grid (page breaks, print area, ruler margins). */
  const gridPageSetup = (commandId: string) => ({
    store,
    history,
    origin: 'editor' as const,
    commandId,
  });

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
      const pagination = paginationFor(s, sheet, { engine: wb });
      const origin = handle.axis === 'row' ? pagination.origin.row : pagination.origin.col;
      if (target < origin) return;
      pageSetupWrite(gridPageSetup('printArea'), sheet, undefined, () =>
        resizePrintArea(store, sheet, handle.axis, target, pagination.content, history),
      );
      return;
    }
    if (target === handle.index) return;
    pageSetupWrite(gridPageSetup('pageBreaks'), sheet, undefined, () =>
      movePageBreak(store, sheet, handle.axis, handle.index, target, history),
    );
  };

  applyDragAt = (x, y) => applyDrag(null, x, y);

  const onUp = (e: PointerEvent): void => {
    autoscroll.stop();
    releaseHeldPointer(host, e.pointerId);
    if (
      (selectionDisabled() && drag.kind !== 'none' && drag.kind !== 'range-insert') ||
      (hostPolicy() !== undefined &&
        (drag.kind === 'col-resize' ||
          drag.kind === 'row-resize' ||
          drag.kind === 'ruler-margin' ||
          drag.kind === 'page-break'))
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
      resizeTooltip.hide();
      // One undo entry per drag, not per pixel: capture pre at drag-start and
      // post here, push the closure pair. Engine-side sync rides on the same
      // snapshot pair so undo/redo replays the resize in the workbook too.
      const before = drag.preLayout;
      const after = captureLayoutSnapshot(store.getState());
      const sheet = store.getState().data.sheetIndex;
      const axis = drag.kind === 'col-resize' ? 'col' : 'row';
      const committed = layoutWrite(headerResizeContext(axis), axis, false, () => {
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
        return true;
      });
      // A resize refused while the drag ran leaves the sizes it started from.
      if (!committed) applyLayoutSnapshot(store, before);
    }
    if (drag.kind === 'ruler-margin') {
      const { handle, inches } = drag;
      drag = { kind: 'none' };
      host.style.cursor = '';
      const sheet = store.getState().data.sheetIndex;
      pageSetupWrite(gridPageSetup('marginsPreset'), sheet, undefined, () =>
        setPageSetup(store, sheet, { margins: { [handle.side]: inches } }, history),
      );
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
    if (drag.kind === 'pivot-item') {
      const { slots, fromRow, after } = drag;
      mutators.setMovePreview(store, null);
      fillTooltip.hide();
      drag = { kind: 'none' };
      if (movePivotItem({ store, workbook: wb, history }, slots, fromRow, after)) onAfterCommit?.();
      return;
    }
    if (drag.kind === 'move') {
      const target = store.getState().ui.movePreview;
      mutators.setMovePreview(store, null);
      const { src, menu, pressed } = drag;
      drag = { kind: 'none' };
      host.style.cursor = '';
      if (menu) {
        const click =
          !target &&
          pressed !== undefined &&
          Math.hypot(e.clientX - pressed.x, e.clientY - pressed.y) <= RIGHT_CLICK_SLOP_PX;
        if (click) {
          // A right press that never became a drag is a plain right-click: let the
          // cell menu open for the current selection, then swallow the native event.
          releaseContextMenuSwallow?.();
          host.dispatchEvent(
            new MouseEvent('contextmenu', {
              bubbles: true,
              cancelable: true,
              clientX: e.clientX,
              clientY: e.clientY,
            }),
          );
          swallowContextMenu();
        }
        finishSecondaryDrag();
        if (target) {
          host.dispatchEvent(
            new CustomEvent('fc:movemenu', {
              bubbles: true,
              detail: { src, dest: { ...target }, clientX: e.clientX, clientY: e.clientY },
            }),
          );
        }
        return;
      }
      const report = pivotTableAtRange(wb, src);
      if (target && report !== null) {
        const moved = movePivotTable(
          { store, workbook: wb, history, origin: 'editor', commandId: 'pivotMove' },
          { sheet: src.sheet, pivotIndex: report },
          { sheet: target.sheet, row: target.r0, col: target.c0 },
        );
        if (moved) onAfterCommit?.();
        return;
      }
      if (target) {
        void moveSelectionTo(
          { store, wb, history, commandId: 'dragMove', confirmOverwrite: confirmMoveOverwrite },
          { row: target.r0, col: target.c0 },
        ).then((moved) => {
          if (moved) onAfterCommit?.();
        });
      }
      return;
    }
    if (drag.kind === 'table-resize') {
      const s = store.getState();
      const range = s.ui.fillPreview;
      const table = tableOverlayById(s, drag.id);
      mutators.setFillPreview(store, null);
      drag = { kind: 'none' };
      if (range && table && isNavigationRangeAllowed(store, range)) {
        const ctx = {
          store,
          history,
          workbook: wb,
          origin: 'editor',
          commandId: TABLE_RESIZE_COMMAND,
        } as const;
        if (applyTableResize(ctx, table, range)) onAfterCommit?.();
      }
      refreshCursor(e);
      return;
    }
    if (drag.kind === 'fill') {
      if (drag.menu) finishSecondaryDrag();
      const s = store.getState();
      const dest = s.ui.fillPreview;
      mutators.setFillPreview(store, null);
      fillTooltip.hide();
      const unchanged = drag.menu === true && dest !== null && sameRange(dest, drag.src);
      const shrunk =
        dest !== null &&
        !drag.menu &&
        dest.r0 === drag.src.r0 &&
        dest.c0 === drag.src.c0 &&
        (dest.r1 < drag.src.r1 || dest.c1 < drag.src.c1);
      if (dest && shrunk) {
        commitFillShrink(s, drag.src, dest);
      } else if (dest && !unchanged) {
        if (!isNavigationRangeAllowed(store, dest)) {
          drag = { kind: 'none' };
          refreshCursor(e);
          return;
        }
        // The chip hangs off the last filled cell, not the pointer.
        const filled = fillOutputRange(drag.src, dest);
        const last = cellRect(layoutForView(s), s.viewport, filled.r1, filled.c1);
        const anchor = clientAnchorOf(e, last);
        if (drag.menu) {
          host.dispatchEvent(
            new CustomEvent('fc:autofillmenu', {
              bubbles: true,
              detail: {
                src: { ...drag.src },
                dest: { ...dest },
                clientX: e.clientX,
                clientY: e.clientY,
                anchorX: anchor.x,
                anchorY: anchor.y,
              },
            }),
          );
          drag = { kind: 'none' };
          refreshCursor(e);
          return;
        }
        // Spreadsheet parity: holding Ctrl/⌘ on release toggles series → tile copy.
        const copyOnly = e.ctrlKey || e.metaKey;
        commitFill(drag.src, dest, {
          copyOnly,
          announce: { clientX: anchor.x, clientY: anchor.y },
        });
      }
    }
    if (drag.kind === 'cell' && drag.link) followClickedLink(drag.link, e.timeStamp);
    drag = { kind: 'none' };
    refreshCursor(e);
  };

  /** A plain click on a link cell follows it once released; a long press or a
   *  drag onto other cells only selects. */
  const followClickedLink = (link: ClickedLink, upAt: number): void => {
    if (upAt - link.at >= HYPERLINK_HOLD_SELECT_MS) return;
    const now = store.getState().selection.range;
    const was = link.range;
    if (!sameRange(now, was)) return;
    followHyperlink({ store, wb, ...hyperlinkFollow }, link.url);
  };

  /** The browser may raise `contextmenu` after the release, so the swallow outlives it by a tick. */
  const finishSecondaryDrag = (): void => {
    const release = releaseContextMenuSwallow;
    if (release) setTimeout(release, 0);
  };

  /** An interrupted gesture is abandoned: live previews revert and nothing is committed. */
  const onCancel = (e: PointerEvent): void => {
    finishSecondaryDrag();
    autoscroll.stop();
    releaseHeldPointer(host, e.pointerId);
    if (drag.kind === 'selection-marquee') {
      const base = drag.base;
      store.setState((s) => ({ ...s, selection: base }));
    } else if (drag.kind === 'col-resize' || drag.kind === 'row-resize') {
      applyLayoutSnapshot(store, drag.preLayout);
    }
    drag = { kind: 'none' };
    mutators.setPageBreakDrag(store, null);
    mutators.setFillPreview(store, null);
    mutators.setMovePreview(store, null);
    fillTooltip.hide();
    resizeTooltip.hide();
    host.style.cursor = '';
  };

  const onLeave = (): void => {
    if (drag.kind === 'none') host.style.cursor = '';
  };

  const onDblClick = (e: MouseEvent): void => {
    const { x, y } = localXY(e);
    const s = store.getState();
    if (selectionDisabled()) {
      e.preventDefault();
      e.stopPropagation();
      return;
    }

    // Fill-handle takes precedence — spreadsheet-style "double-click to flash-fill
    // down to match the neighbour column's contiguous run."
    if (isFillHandleHit(hits, x, y)) {
      e.preventDefault();
      e.stopPropagation();
      const src = { ...s.selection.range };
      const dest = autoFillDownExtent(s, src);
      if (!dest) return;
      if (!isNavigationRangeAllowed(store, dest)) return;
      commitFill(src, dest, { copyOnly: false });
      return;
    }

    // Double-clicking the selection border jumps to the edge of the data region
    // in that direction, as Cmd+Arrow does, instead of editing the cell.
    const edge = selectionBorderEdgeAt(host, s, x, y);
    if (edge) {
      e.preventDefault();
      e.stopPropagation();
      if (edge === 'corner') return;
      const colDir = s.ui.rightToLeft === true ? -1 : 1;
      const [dRow, dCol] =
        edge === 'bottom'
          ? [1, 0]
          : edge === 'top'
            ? [-1, 0]
            : edge === 'right'
              ? [0, colDir]
              : [0, -colDir];
      // The jump starts from the frame cell under the pointer, not the active
      // cell; a blank frame cell just becomes the selection.
      const under = hitTest(layoutForView(s), s.viewport, x, y);
      if (!under) return;
      const range = s.selection.range;
      const origin: Addr = {
        sheet: s.data.sheetIndex,
        row:
          dRow === 0
            ? Math.max(range.r0, Math.min(range.r1, under.row))
            : dRow > 0
              ? range.r1
              : range.r0,
        col:
          dCol === 0
            ? Math.max(range.c0, Math.min(range.c1, under.col))
            : dCol > 0
              ? range.c1
              : range.c0,
      };
      const target = isPopulated(s, origin.sheet, origin.row, origin.col)
        ? jumpToDataEdge(s, origin, dRow, dCol)
        : origin;
      if (!isNavigationAddrAllowed(store, target)) return;
      mutators.setActive(store, target);
      scrollActiveIntoView(store, target);
      return;
    }

    const zone = hitZone(layoutForView(s), s.viewport, x, y);
    if (!zone) return;

    if (zone.kind === 'col-resize' || zone.kind === 'row-resize') {
      e.preventDefault();
      e.stopPropagation();
      const axis = zone.kind === 'col-resize' ? 'col' : 'row';
      const track = zone.kind === 'col-resize' ? zone.col : zone.row;
      layoutWrite(headerResizeContext(axis), axis, undefined, () =>
        autofitHeaderTracks(
          store,
          history,
          axis,
          track,
          headerGestureSpans(s, axis, track),
          wb,
          getAutofitOptions(),
        ),
      );
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
    autoscroll.stop();
    host.removeEventListener('pointerdown', onDown);
    host.removeEventListener('pointermove', onMove);
    host.removeEventListener('pointerup', onUp);
    host.removeEventListener('pointercancel', onCancel);
    host.removeEventListener('pointerleave', onLeave);
    host.removeEventListener('dblclick', onDblClick);
    fillTooltip.dispose();
    resizeTooltip.dispose();
    host.style.cursor = '';
  };
}
