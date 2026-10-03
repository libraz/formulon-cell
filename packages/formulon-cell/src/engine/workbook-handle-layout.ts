import type { EngineCapabilities, Range, Workbook } from './types.js';
import { SheetVisibility } from './types.js';
import type { WorkbookHandle } from './workbook-handle.js';

type WorkbookHandleCtor = { prototype: WorkbookHandle };
type WorkbookHandleInternals = {
  wb: Workbook;
  assertAlive(): void;
};

declare module './workbook-handle.js' {
  interface WorkbookHandle extends WorkbookHandleLayoutMethods {}
}

function internals(handle: unknown): WorkbookHandleInternals {
  return handle as WorkbookHandleInternals;
}

function assertAlive(handle: unknown): void {
  internals(handle).assertAlive();
}

function wb(handle: unknown): Workbook {
  return internals(handle).wb;
}

/** Resolve a sheet view's tab state. `visibility` is authoritative when the
 *  engine reports it; an engine that carries only the two-state `tabHidden`
 *  cannot distinguish `veryHidden`, so such a sheet reads as `Hidden`. */
function sheetVisibilityOf(view: { tabHidden: number; visibility?: number }): SheetVisibility {
  if (view.visibility === SheetVisibility.VeryHidden) return SheetVisibility.VeryHidden;
  if (view.visibility === SheetVisibility.Hidden || view.tabHidden !== 0) {
    return SheetVisibility.Hidden;
  }
  return SheetVisibility.Visible;
}

export abstract class WorkbookHandleLayoutMethods {
  declare readonly capabilities: EngineCapabilities;

  /** Persist a column-width override on `[first, last]` for `sheet`.
   *  No-op (returns false) when the engine doesn't expose `setColumnWidth`
   *  — i.e. the stub fallback. The UI store stays the source of truth in
   *  that case so paint still reflects the drag. */
  setColumnWidth(sheet: number, first: number, last: number, width: number): boolean {
    assertAlive(this);
    if (!this.capabilities.colRowSize) return false;
    const s = wb(this).setColumnWidth(sheet, first, last, width);
    return s.ok;
  }

  /** Persist a row-height override at `row` for `sheet`. See `setColumnWidth`
   *  for the no-op-on-stub rationale. */
  setRowHeight(sheet: number, row: number, height: number): boolean {
    assertAlive(this);
    if (!this.capabilities.colRowSize) return false;
    const s = wb(this).setRowHeight(sheet, row, height);
    return s.ok;
  }

  /** Snapshot of column overrides on `sheet`. Empty array under the stub.
   *  The returned objects own no engine memory — the underlying vector
   *  handle is released before this method returns. */
  getColumnLayouts(
    sheet: number,
  ): { first: number; last: number; width: number; hidden: boolean; outlineLevel: number }[] {
    assertAlive(this);
    if (!this.capabilities.colRowSize) return [];
    const r = wb(this).getSheetColumns(sheet);
    const out: {
      first: number;
      last: number;
      width: number;
      hidden: boolean;
      outlineLevel: number;
    }[] = [];
    if (!r.status.ok) return out;
    for (const e of r.columns) {
      out.push({
        first: e.first,
        last: e.last,
        width: e.width,
        hidden: e.hidden !== 0,
        outlineLevel: e.outlineLevel,
      });
    }
    return out;
  }

  /** Persist frozen-pane counts on `sheet`. No-op (returns false) under stub. */
  setSheetFreeze(sheet: number, freezeRows: number, freezeCols: number): boolean {
    assertAlive(this);
    if (!this.capabilities.freeze) return false;
    const s = wb(this).setSheetFreeze(sheet, freezeRows, freezeCols);
    return s.ok;
  }

  /** Persist sheet zoom percentage (10..400, engine clamps). No-op under stub. */
  setSheetZoom(sheet: number, zoomScale: number): boolean {
    assertAlive(this);
    if (!this.capabilities.sheetZoom) return false;
    const s = wb(this).setSheetZoom(sheet, zoomScale);
    return s.ok;
  }

  /** Toggle the tab-hidden flag on `sheet`. Returns false on engine failure
   *  or when the engine doesn't expose `setSheetTabHidden`.
   *
   *  This is the two-state view: `true` on an already very-hidden sheet leaves
   *  it very-hidden, and `false` reveals it from either hidden state. Use
   *  `setSheetVisibility` to move between the two hidden states. */
  setSheetTabHidden(sheet: number, hidden: boolean): boolean {
    assertAlive(this);
    if (!this.capabilities.sheetTabHidden) return false;
    const s = wb(this).setSheetTabHidden(sheet, hidden);
    return s.ok;
  }

  /** Set `sheet`'s tab to one of the three OOXML visibility states. Returns
   *  false when the engine only carries the two-state `setSheetTabHidden`,
   *  which cannot express `veryHidden`. */
  setSheetVisibility(sheet: number, visibility: SheetVisibility): boolean {
    assertAlive(this);
    if (!this.capabilities.sheetVisibility) return false;
    const s = wb(this).setSheetVisibility(sheet, visibility as number);
    return s.ok;
  }

  /** Set the hidden flag on `[first, last]` columns. No-op under stub or
   *  when the engine doesn't expose `setColumnHidden`. */
  setColumnHidden(sheet: number, first: number, last: number, hidden: boolean): boolean {
    assertAlive(this);
    if (!this.capabilities.hiddenRowsCols) return false;
    const s = wb(this).setColumnHidden(sheet, first, last, hidden);
    return s.ok;
  }

  /** Set the hidden flag on `row`. */
  setRowHidden(sheet: number, row: number, hidden: boolean): boolean {
    assertAlive(this);
    if (!this.capabilities.hiddenRowsCols) return false;
    const s = wb(this).setRowHidden(sheet, row, hidden);
    return s.ok;
  }

  /** Set the outline level on `[first, last]` columns (0..7). */
  setColumnOutline(sheet: number, first: number, last: number, level: number): boolean {
    assertAlive(this);
    if (!this.capabilities.outlines) return false;
    const s = wb(this).setColumnOutline(sheet, first, last, level);
    return s.ok;
  }

  /** Set the outline level on `row` (0..7). */
  setRowOutline(sheet: number, row: number, level: number): boolean {
    assertAlive(this);
    if (!this.capabilities.outlines) return false;
    const s = wb(this).setRowOutline(sheet, row, level);
    return s.ok;
  }

  /** Snapshot of `sheet`'s view, including display flags. Returns null when
   *  the engine doesn't expose `getSheetView` (i.e. the stub or an older bundle). */
  getSheetView(sheet: number): {
    zoomScale: number;
    freezeRows: number;
    freezeCols: number;
    tabHidden: boolean;
    visibility: SheetVisibility;
    showGridLines: boolean;
    showRowColHeaders: boolean;
    showZeros: boolean;
    rightToLeft: boolean;
  } | null {
    assertAlive(this);
    if (!this.capabilities.sheetView) return null;
    const r = wb(this).getSheetView(sheet);
    if (!r.status.ok) return null;
    return {
      zoomScale: r.view.zoomScale,
      freezeRows: r.view.freezeRows,
      freezeCols: r.view.freezeCols,
      tabHidden: r.view.tabHidden !== 0,
      // Engines predating three-state visibility carry only `tabHidden`.
      visibility: sheetVisibilityOf(r.view),
      showGridLines: r.view.showGridLines !== 0,
      showRowColHeaders: r.view.showRowColHeaders !== 0,
      showZeros: r.view.showZeros !== 0,
      rightToLeft: r.view.rightToLeft !== 0,
    };
  }

  setSheetShowGridLines(sheet: number, show: boolean): boolean {
    assertAlive(this);
    if (!this.capabilities.sheetViewFlags) return false;
    return wb(this).setSheetShowGridLines(sheet, show).ok;
  }

  setSheetShowRowColHeaders(sheet: number, show: boolean): boolean {
    assertAlive(this);
    if (!this.capabilities.sheetViewFlags) return false;
    return wb(this).setSheetShowRowColHeaders(sheet, show).ok;
  }

  setSheetShowZeros(sheet: number, show: boolean): boolean {
    assertAlive(this);
    if (!this.capabilities.sheetViewFlags) return false;
    return wb(this).setSheetShowZeros(sheet, show).ok;
  }

  setSheetRightToLeft(sheet: number, rightToLeft: boolean): boolean {
    assertAlive(this);
    if (!this.capabilities.sheetViewFlags) return false;
    return wb(this).setSheetRightToLeft(sheet, rightToLeft).ok;
  }

  /** Insert `count` blank rows at `row` on `sheet`. The engine rewrites
   *  cross-workbook formula refs to follow the shift. Returns false on
   *  engines without `insertDeleteRowsCols`. NOT routed through the
   *  per-cell journal — callers wrap this in their own history entry. */
  engineInsertRows(sheet: number, row: number, count: number): boolean {
    assertAlive(this);
    if (!this.capabilities.insertDeleteRowsCols) return false;
    const s = wb(this).insertRows(sheet, row, count);
    return s.ok;
  }

  /** Delete `count` rows starting at `row` on `sheet`. Refs that fall
   *  inside the deleted interval collapse to `#REF!`. */
  engineDeleteRows(sheet: number, row: number, count: number): boolean {
    assertAlive(this);
    if (!this.capabilities.insertDeleteRowsCols) return false;
    const s = wb(this).deleteRows(sheet, row, count);
    return s.ok;
  }

  /** Insert `count` blank columns at `col` on `sheet`. */
  engineInsertCols(sheet: number, col: number, count: number): boolean {
    assertAlive(this);
    if (!this.capabilities.insertDeleteRowsCols) return false;
    const s = wb(this).insertCols(sheet, col, count);
    return s.ok;
  }

  /** Delete `count` columns starting at `col` on `sheet`. */
  engineDeleteCols(sheet: number, col: number, count: number): boolean {
    assertAlive(this);
    if (!this.capabilities.insertDeleteRowsCols) return false;
    const s = wb(this).deleteCols(sheet, col, count);
    return s.ok;
  }

  /** Append `range` as a merge on `sheet`. Returns false on engine failure or
   *  when `capabilities.merges` is off. The cell content inside the range is
   *  the caller's responsibility (spreadsheets keep top-left, blanks the rest). */
  engineAddMerge(sheet: number, range: Range): boolean {
    assertAlive(this);
    if (!this.capabilities.merges) return false;
    const s = wb(this).addMerge(sheet, {
      firstRow: range.r0,
      firstCol: range.c0,
      lastRow: range.r1,
      lastCol: range.c1,
    });
    return s.ok;
  }

  /** Remove every merge on `sheet` overlapping `range` (inclusive). No-op when
   *  nothing overlaps. Returns false on engine failure or capability off. */
  engineRemoveMerge(sheet: number, range: Range): boolean {
    assertAlive(this);
    if (!this.capabilities.merges) return false;
    const s = wb(this).removeMerge(sheet, {
      firstRow: range.r0,
      firstCol: range.c0,
      lastRow: range.r1,
      lastCol: range.c1,
    });
    return s.ok;
  }

  /** Drop every merge on `sheet`. Returns false on engine failure or capability
   *  off. */
  engineClearMerges(sheet: number): boolean {
    assertAlive(this);
    if (!this.capabilities.merges) return false;
    const s = wb(this).clearMerges(sheet);
    return s.ok;
  }

  /** Snapshot of every merge on `sheet` as inclusive `Range` records. Empty
   *  array under stub or when `capabilities.merges` is off. */
  getMerges(sheet: number): Range[] {
    assertAlive(this);
    if (!this.capabilities.merges) return [];
    const arr = wb(this).getMerges(sheet);
    if (!arr.status.ok) return [];
    return arr.map((m) => ({
      sheet,
      r0: m.firstRow,
      c0: m.firstCol,
      r1: m.lastRow,
      c1: m.lastCol,
    }));
  }

  /** Snapshot of row overrides on `sheet`. See `getColumnLayouts`. */
  getRowLayouts(
    sheet: number,
  ): { row: number; height: number; hidden: boolean; outlineLevel: number }[] {
    assertAlive(this);
    if (!this.capabilities.colRowSize) return [];
    const r = wb(this).getSheetRowOverrides(sheet);
    const out: { row: number; height: number; hidden: boolean; outlineLevel: number }[] = [];
    if (!r.status.ok) return out;
    for (const e of r.rows) {
      out.push({
        row: e.row,
        height: e.height,
        hidden: e.hidden !== 0,
        outlineLevel: e.outlineLevel,
      });
    }
    return out;
  }
}

export function installLayoutMethods(target: WorkbookHandleCtor): void {
  for (const key of Object.getOwnPropertyNames(WorkbookHandleLayoutMethods.prototype)) {
    if (key === 'constructor') continue;
    const descriptor = Object.getOwnPropertyDescriptor(WorkbookHandleLayoutMethods.prototype, key);
    if (!descriptor) continue;
    Object.defineProperty(target.prototype, key, descriptor);
  }
}
