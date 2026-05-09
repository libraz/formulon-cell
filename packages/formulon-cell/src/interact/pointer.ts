import { fillDestFor, fillRange } from '../commands/fill.js';
import { formatNumber } from '../commands/format.js';
import {
  applyLayoutSnapshot,
  captureLayoutSnapshot,
  type History,
  type LayoutSnapshot,
} from '../commands/history.js';
import { applyUnmerge, expandRangeWithMerges, mergeAnchorOf } from '../commands/merge.js';
import {
  collapseColGroup,
  collapseRowGroup,
  expandColGroup,
  expandRowGroup,
  isColGroupCollapsed,
  isRowGroupCollapsed,
} from '../commands/outline.js';
import { syncLayoutSizesToEngine } from '../engine/layout-sync.js';
import type { Range } from '../engine/types.js';
import { formatCell } from '../engine/value.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { hitTest, hitZone } from '../render/geometry.js';
import { getFillHandleRect, getOutlineToggleHits } from '../render/grid.js';
import { type CellFormat, mutators, type SpreadsheetStore, type State } from '../store/store.js';

type StoreCellEntry = State['data']['cells'] extends Map<string, infer Cell> ? Cell : never;

type DragMode =
  | { kind: 'none' }
  | { kind: 'cell' }
  | { kind: 'col-header'; anchorCol: number }
  | { kind: 'row-header'; anchorRow: number }
  | { kind: 'col-resize'; col: number; leftEdge: number; preLayout: LayoutSnapshot }
  | { kind: 'row-resize'; row: number; topEdge: number; preLayout: LayoutSnapshot }
  | { kind: 'fill'; src: Range }
  | {
      kind: 'range-insert';
      anchor: { row: number; col: number };
      tip: { row: number; col: number };
    };

const colLetters = (col: number): string => {
  let n = col + 1;
  let s = '';
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
};

const refOf = (row: number, col: number): string => `${colLetters(col)}${row + 1}`;
const rangeRefOf = (a: { row: number; col: number }, b: { row: number; col: number }): string => {
  if (a.row === b.row && a.col === b.col) return refOf(a.row, a.col);
  const r0 = Math.min(a.row, b.row);
  const r1 = Math.max(a.row, b.row);
  const c0 = Math.min(a.col, b.col);
  const c1 = Math.max(a.col, b.col);
  return `${refOf(r0, c0)}:${refOf(r1, c1)}`;
};

const MAX_ROW = 1048575;
const MAX_COL = 16383;
const FILTER_DROPDOWN_RESERVED_WIDTH = 20;

export interface PointerDeps {
  store: SpreadsheetStore;
  wb: WorkbookHandle;
  /** Refresh cached cells after a write — same contract as the inline editor. */
  onAfterCommit?: () => void;
  /** Shared history. When provided, col/row resizes and fill drags push one
   *  entry per drag-end (not per intermediate frame). */
  history?: History | null;
}

/** Capability surfaced by the inline editor used by the pointer layer to
 *  detect a live formula edit and inject clicked cell references. */
export interface RangeInsertTarget {
  isFormulaEdit: () => boolean;
  insertRefAtCaret: (ref: string) => void;
}

export function attachPointer(
  host: HTMLElement,
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  onAfterCommit?: () => void,
  history: History | null = null,
  getEditor: () => RangeInsertTarget | null = () => null,
): () => void {
  let drag: DragMode = { kind: 'none' };
  const measureCanvas = document.createElement('canvas');
  const measureCtx = measureCanvas.getContext('2d');

  const localXY = (e: PointerEvent | MouseEvent): { x: number; y: number } => {
    const rect = host.getBoundingClientRect();
    return { x: e.clientX - rect.left, y: e.clientY - rect.top };
  };

  const isFillHandleHit = (x: number, y: number): boolean => {
    const rect = getFillHandleRect();
    if (!rect) return false;
    // Pad by a couple of pixels so the handle is comfortable to grab.
    const pad = 3;
    return (
      x >= rect.x - pad &&
      x <= rect.x + rect.w + pad &&
      y >= rect.y - pad &&
      y <= rect.y + rect.h + pad
    );
  };

  const onDown = (e: PointerEvent): void => {
    if (e.button !== 0) return;
    const { x, y } = localXY(e);
    const s = store.getState();

    // Capture editor intent BEFORE we touch focus — host.focus() blurs the
    //  textarea, which triggers commit + cancel and tears down the editor
    //  before our cell-zone branch could query it.
    const editor = getEditor();
    const inFormula = editor?.isFormulaEdit();

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

    // Fill handle takes precedence over normal cell hit-testing.
    if (isFillHandleHit(x, y)) {
      host.focus();
      tryCapture();
      drag = { kind: 'fill', src: { ...s.selection.range } };
      mutators.setFillPreview(store, { ...s.selection.range });
      host.style.cursor = 'crosshair';
      return;
    }

    const zone = hitZone(s.layout, s.viewport, x, y, s.ui.filterRange);
    if (!zone) return;

    // Range-insert: keep focus on the editor (preventDefault avoids the
    //  pointerdown's default focus-change behavior).
    if (inFormula && zone.kind === 'cell' && editor) {
      e.preventDefault();
      tryCapture();
      const anchor = { row: zone.row, col: zone.col };
      editor.insertRefAtCaret(refOf(anchor.row, anchor.col));
      drag = { kind: 'range-insert', anchor, tip: anchor };
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
                  y: s.layout.outlineColGutter,
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
        mutators.selectRow(store, zone.row);
        drag = { kind: 'row-header', anchorRow: zone.row };
        return;
      }

      case 'col-resize': {
        const leftEdge =
          s.layout.outlineRowGutter +
          s.layout.headerColWidth +
          colXFromState(
            s.layout.colWidths,
            s.layout.defaultColWidth,
            s.viewport.colStart,
            zone.col,
          );
        drag = {
          kind: 'col-resize',
          col: zone.col,
          leftEdge,
          preLayout: captureLayoutSnapshot(s),
        };
        return;
      }

      case 'row-resize': {
        const topEdge =
          s.layout.outlineColGutter +
          s.layout.headerRowHeight +
          rowYFromState(
            s.layout.rowHeights,
            s.layout.defaultRowHeight,
            s.viewport.rowStart,
            zone.row,
          );
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
          // Disjoint additive selection. Drag is suppressed so the user
          // doesn't accidentally turn a click into a multi-range marquee.
          mutators.addExtraCell(store, addr);
          drag = { kind: 'none' };
          return;
        }
        if (e.shiftKey) {
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
        } else mutators.setActive(store, addr);
        drag = { kind: 'cell' };
        return;
      }
    }
  };

  const onMove = (e: PointerEvent): void => {
    const { x, y } = localXY(e);

    if (drag.kind === 'none') {
      updateCursor(host, store, x, y);
      return;
    }

    const s = store.getState();

    switch (drag.kind) {
      case 'col-resize': {
        const w = x - drag.leftEdge;
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
        const zone = hitZone(s.layout, s.viewport, x, y);
        if (zone && (zone.kind === 'col-header' || zone.kind === 'col-resize')) {
          mutators.extendRangeTo(store, { sheet: s.data.sheetIndex, row: MAX_ROW, col: zone.col });
        }
        return;
      }
      case 'row-header': {
        const zone = hitZone(s.layout, s.viewport, x, y);
        if (zone && (zone.kind === 'row-header' || zone.kind === 'row-resize')) {
          mutators.extendRangeTo(store, { sheet: s.data.sheetIndex, row: zone.row, col: MAX_COL });
        }
        return;
      }
      case 'cell': {
        const zone = hitZone(s.layout, s.viewport, x, y);
        if (zone && zone.kind === 'cell') {
          mutators.extendRangeTo(store, {
            sheet: s.data.sheetIndex,
            row: zone.row,
            col: zone.col,
          });
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
        }
        return;
      }
      case 'fill': {
        const cell = hitTest(s.layout, s.viewport, x, y);
        if (!cell) return;
        const dest = fillDestFor(drag.src, { row: cell.row, col: cell.col });
        mutators.setFillPreview(store, dest);
        host.style.cursor = 'crosshair';
        return;
      }
      case 'range-insert': {
        const cell = hitTest(s.layout, s.viewport, x, y);
        if (!cell) return;
        if (cell.row === drag.tip.row && cell.col === drag.tip.col) return;
        drag.tip = { row: cell.row, col: cell.col };
        const editor = getEditor();
        if (editor) {
          editor.insertRefAtCaret(rangeRefOf(drag.anchor, drag.tip));
        }
        return;
      }
    }
  };

  const onUp = (e: PointerEvent): void => {
    if (host.hasPointerCapture(e.pointerId)) host.releasePointerCapture(e.pointerId);
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
    if (drag.kind === 'fill') {
      const s = store.getState();
      const dest = s.ui.fillPreview;
      mutators.setFillPreview(store, null);
      if (dest) {
        // Spreadsheet parity: holding Ctrl/⌘ on release toggles series → tile copy.
        const copyOnly = e.ctrlKey || e.metaKey;
        // Bundle every per-cell write into a single undoable transaction.
        if (history) history.begin();
        let wrote = false;
        try {
          // Strip any merges that intersect the fill destination — fill cannot
          // tear merged rectangles apart silently.
          applyUnmerge(store, wb, history, dest);
          wrote = fillRange(s, wb, drag.src, dest, { copyOnly });
        } finally {
          if (history) history.end();
        }
        if (wrote) {
          onAfterCommit?.();
          // Promote dest as the new selection.
          mutators.setActive(store, { sheet: dest.sheet, row: dest.r0, col: dest.c0 });
          mutators.extendRangeTo(store, { sheet: dest.sheet, row: dest.r1, col: dest.c1 });
        }
      }
    }
    drag = { kind: 'none' };
    const { x, y } = localXY(e);
    updateCursor(host, store, x, y);
  };

  const onLeave = (): void => {
    if (drag.kind === 'none') host.style.cursor = '';
  };

  const onDblClick = (e: MouseEvent): void => {
    const { x, y } = localXY(e);
    const s = store.getState();

    // Fill-handle takes precedence — spreadsheet-style "double-click to flash-fill
    // down to match the neighbour column's contiguous run."
    if (isFillHandleHit(x, y)) {
      e.preventDefault();
      e.stopPropagation();
      const src = { ...s.selection.range };
      const dest = autoFillDownExtent(s, src);
      if (!dest) return;
      if (history) history.begin();
      let wrote = false;
      try {
        wrote = fillRange(s, wb, src, dest);
      } finally {
        if (history) history.end();
      }
      if (wrote) {
        onAfterCommit?.();
        mutators.setActive(store, { sheet: dest.sheet, row: dest.r0, col: dest.c0 });
        mutators.extendRangeTo(store, { sheet: dest.sheet, row: dest.r1, col: dest.c1 });
      }
      return;
    }

    const zone = hitZone(s.layout, s.viewport, x, y);
    if (!zone) return;

    if (zone.kind === 'col-resize') {
      e.preventDefault();
      e.stopPropagation();
      const before = captureLayoutSnapshot(s);
      const w = autofitColWidth(store, zone.col, measureCtx);
      mutators.setColWidth(store, zone.col, w);
      pushLayoutDelta(history, store, wb, before);
      return;
    }
    if (zone.kind === 'row-resize') {
      e.preventDefault();
      e.stopPropagation();
      const before = captureLayoutSnapshot(s);
      const h = autofitRowHeight(store, zone.row, measureCtx);
      mutators.setRowHeight(store, zone.row, h);
      pushLayoutDelta(history, store, wb, before);
      return;
    }
  };

  host.addEventListener('pointerdown', onDown);
  host.addEventListener('pointermove', onMove);
  host.addEventListener('pointerup', onUp);
  host.addEventListener('pointercancel', onUp);
  host.addEventListener('pointerleave', onLeave);
  host.addEventListener('dblclick', onDblClick);

  return () => {
    host.removeEventListener('pointerdown', onDown);
    host.removeEventListener('pointermove', onMove);
    host.removeEventListener('pointerup', onUp);
    host.removeEventListener('pointercancel', onUp);
    host.removeEventListener('pointerleave', onLeave);
    host.removeEventListener('dblclick', onDblClick);
    host.style.cursor = '';
  };
}

function pushLayoutDelta(
  history: History | null,
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  before: LayoutSnapshot,
): void {
  const s = store.getState();
  const after = captureLayoutSnapshot(s);
  const sheet = s.data.sheetIndex;
  syncLayoutSizesToEngine(wb, s.layout, sheet, before, after);
  if (!history || history.isReplaying()) return;
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

function updateCursor(host: HTMLElement, store: SpreadsheetStore, x: number, y: number): void {
  const handle = getFillHandleRect();
  if (handle) {
    const pad = 3;
    if (
      x >= handle.x - pad &&
      x <= handle.x + handle.w + pad &&
      y >= handle.y - pad &&
      y <= handle.y + handle.h + pad
    ) {
      host.style.cursor = 'crosshair';
      return;
    }
  }
  const s = store.getState();
  const zone = hitZone(s.layout, s.viewport, x, y, s.ui.filterRange);
  if (!zone) {
    host.style.cursor = '';
    return;
  }
  if (zone.kind === 'col-resize') host.style.cursor = 'col-resize';
  else if (zone.kind === 'row-resize') host.style.cursor = 'row-resize';
  else if (zone.kind === 'col-filter-btn') host.style.cursor = 'pointer';
  else host.style.cursor = '';
}

function colXFromState(
  widths: Map<number, number>,
  def: number,
  colStart: number,
  col: number,
): number {
  let x = 0;
  for (let c = colStart; c < col; c += 1) x += widths.get(c) ?? def;
  return x;
}

function rowYFromState(
  heights: Map<number, number>,
  def: number,
  rowStart: number,
  row: number,
): number {
  let y = 0;
  for (let r = rowStart; r < row; r += 1) y += heights.get(r) ?? def;
  return y;
}

function autofitColWidth(
  store: SpreadsheetStore,
  col: number,
  ctx: CanvasRenderingContext2D | null,
): number {
  const s = store.getState();
  const sheet = s.data.sheetIndex;
  const padding = 16;
  const minWidth = 48;
  let max = 0;

  for (const [key, cell] of s.data.cells) {
    const parsed = parseCellKey(key);
    if (!parsed || parsed.sheet !== sheet || parsed.col !== col) continue;
    const text = autofitDisplayText(s, key, cell);
    if (!text) continue;
    const fmt = s.format.formats.get(key);
    const fontSize = fmt?.fontSize ?? 13;
    const font = autofitFont(fmt);
    if (ctx) ctx.font = font;
    const width =
      maxExplicitLineWidth(text, ctx, fontSize) +
      (isFilterHeaderCell(s, parsed.sheet, parsed.row, parsed.col)
        ? FILTER_DROPDOWN_RESERVED_WIDTH
        : 0);
    if (width > max) max = width;
  }

  return Math.max(minWidth, Math.ceil(max) + padding);
}

function parseCellKey(key: string): { sheet: number; row: number; col: number } | null {
  const parts = key.split(':');
  if (parts.length !== 3) return null;
  const sheet = Number(parts[0]);
  const row = Number(parts[1]);
  const col = Number(parts[2]);
  if (!Number.isInteger(sheet) || !Number.isInteger(row) || !Number.isInteger(col)) return null;
  return { sheet, row, col };
}

function autofitDisplayText(state: State, key: string, cell: StoreCellEntry): string {
  if (state.ui.showFormulas && cell.formula) return cell.formula;
  const fmt = state.format.formats.get(key);
  if (cell.value.kind === 'number' && fmt?.numFmt) {
    return formatNumber(cell.value.value, fmt.numFmt);
  }
  return formatCell(cell.value);
}

function isFilterHeaderCell(state: State, sheet: number, row: number, col: number): boolean {
  const fr = state.ui.filterRange;
  return !!fr && fr.sheet === sheet && fr.r0 === row && col >= fr.c0 && col <= fr.c1;
}

function autofitRowHeight(
  store: SpreadsheetStore,
  row: number,
  ctx: CanvasRenderingContext2D | null,
): number {
  const s = store.getState();
  const sheet = s.data.sheetIndex;
  let max = s.layout.defaultRowHeight;

  for (const [key, cell] of s.data.cells) {
    const parsed = parseCellKey(key);
    if (!parsed || parsed.sheet !== sheet || parsed.row !== row) continue;
    const text = autofitDisplayText(s, key, cell);
    if (!text) continue;
    const fmt = s.format.formats.get(key);
    const fontSize = fmt?.fontSize ?? 13;
    if (ctx) ctx.font = autofitFont(fmt);
    const lineHeight = Math.round(fontSize * 1.28);
    const colW = s.layout.colWidths.get(parsed.col) ?? s.layout.defaultColWidth;
    const lines = autofitLineCount(text, fmt?.wrap === true, colW, ctx, fontSize);
    const height = Math.ceil(lines * lineHeight + 8);
    if (height > max) max = height;
  }

  return max;
}

function autofitFont(format: CellFormat | undefined): string {
  const styleSlant = format?.italic ? 'italic ' : '';
  const weight = format?.bold ? 700 : 400;
  const size = format?.fontSize ?? 13;
  const family = format?.fontFamily ?? 'system-ui, sans-serif';
  return `${styleSlant}${weight} ${size}px ${fontFamilyCss(family)}`;
}

function fontFamilyCss(family: string): string {
  return family
    .split(',')
    .map((part) => {
      const trimmed = part.trim();
      if (/^["'].*["']$/.test(trimmed) || /^[a-z-]+$/i.test(trimmed)) return trimmed;
      return `"${trimmed.replace(/"/g, '\\"')}"`;
    })
    .join(', ');
}

function maxExplicitLineWidth(
  text: string,
  ctx: CanvasRenderingContext2D | null,
  fontSize: number,
): number {
  let max = 0;
  for (const line of text.split(/\r\n|\r|\n/)) {
    const measured = ctx ? ctx.measureText(line).width : 0;
    const width = measured > 0 ? measured : line.length * fontSize * 0.54;
    if (width > max) max = width;
  }
  return max;
}

function autofitLineCount(
  text: string,
  wrap: boolean,
  colWidth: number,
  ctx: CanvasRenderingContext2D | null,
  fontSize: number,
): number {
  const paragraphs = text.split(/\r\n|\r|\n/);
  if (!wrap) return Math.max(1, paragraphs.length);

  const available = Math.max(1, colWidth - 12);
  let lines = 0;
  for (const paragraph of paragraphs) {
    if (paragraph.length === 0) {
      lines += 1;
      continue;
    }
    const wrapped = wrapAutofitParagraph(paragraph, available, ctx, fontSize);
    lines += wrapped;
  }
  return Math.max(1, lines);
}

function wrapAutofitParagraph(
  paragraph: string,
  maxWidth: number,
  ctx: CanvasRenderingContext2D | null,
  fontSize: number,
): number {
  const words = paragraph.split(/(\s+)/);
  let line = '';
  let count = 0;
  for (const word of words) {
    const candidate = line + word;
    if (measureAutofitText(candidate, ctx, fontSize) <= maxWidth || line === '') {
      line = candidate;
    } else {
      count += 1;
      line = word.trimStart();
    }
  }
  return count + (line ? 1 : 0);
}

function measureAutofitText(
  text: string,
  ctx: CanvasRenderingContext2D | null,
  fontSize: number,
): number {
  const measured = ctx ? ctx.measureText(text).width : 0;
  return measured > 0 ? measured : text.length * fontSize * 0.54;
}

/**
 * Desktop spreadsheets "double-click the fill handle" rule: extend the source range downward
 * by the contiguous-data run of the immediate left-then-right neighbour
 * column. Returns null when neither neighbour has a usable run (so we don't
 * flash a no-op fill).
 *
 * The neighbour run starts at the row just below the source bottom edge
 * (src.r1 + 1) and ends at the last non-blank row in that column. We require
 * at least one non-blank cell at row src.r1 + 1 — without it, spreadsheets don't
 * expand either, so a stray cell ten rows down won't trigger an unexpected
 * fill.
 */
function autoFillDownExtent(state: State, src: Range): Range | null {
  const sheet = src.sheet;
  const start = src.r1 + 1;
  if (start > MAX_ROW) return null;
  const probeCols: number[] = [];
  if (src.c0 > 0) probeCols.push(src.c0 - 1); // left first (desktop spreadsheets preference)
  if (src.c1 < MAX_COL) probeCols.push(src.c1 + 1);

  let endRow = -1;
  for (const col of probeCols) {
    if (!hasCellAt(state, sheet, start, col)) continue;
    let r = start;
    while (r <= MAX_ROW && hasCellAt(state, sheet, r, col)) r += 1;
    const last = r - 1;
    if (last > endRow) endRow = last;
  }
  if (endRow < start) return null;
  return { sheet, r0: src.r0, c0: src.c0, r1: endRow, c1: src.c1 };
}

function hasCellAt(state: State, sheet: number, row: number, col: number): boolean {
  const cell = state.data.cells.get(`${sheet}:${row}:${col}`);
  if (!cell) return false;
  if (cell.formula) return true;
  return cell.value.kind !== 'blank';
}

function hyperlinkAt(
  state: State,
  addr: { sheet: number; row: number; col: number },
): string | null {
  const fmt = state.format.formats.get(`${addr.sheet}:${addr.row}:${addr.col}`);
  const url = fmt?.hyperlink;
  if (typeof url !== 'string' || url.length === 0) return null;
  return url;
}

/**
 * Opens a hyperlink in a new tab. Restricted to safe protocols (http(s)://,
 * mailto:, tel:) so a hostile cell value can't smuggle a `javascript:` URL.
 */
function openHyperlink(url: string): void {
  const trimmed = url.trim();
  if (trimmed.length === 0) return;
  const lower = trimmed.toLowerCase();
  const ok =
    lower.startsWith('http://') ||
    lower.startsWith('https://') ||
    lower.startsWith('mailto:') ||
    lower.startsWith('tel:');
  if (!ok) return;
  if (typeof window === 'undefined' || typeof window.open !== 'function') return;
  window.open(trimmed, '_blank', 'noopener,noreferrer');
}
