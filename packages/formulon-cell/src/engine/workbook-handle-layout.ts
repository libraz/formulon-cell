import {
  CellShift,
  GeometryMode,
  type SheetFormatDefaultsPatch,
  SheetViewMode,
  StatusCode,
  type WidthModel,
  WidthPlatform,
} from '@libraz/formulon';
import type { WorkbookViewMode } from '../store/types.js';
import { rangeOfRect, rectOfRange } from './address.js';
import { displayCharsForScreenPx, screenPxForDisplayChars } from './column-display-width.js';
import { macDigitAdvance } from './mac-digit-metrics.js';
import { macDefaultRowPt } from './mac-row-model.js';
import { attemptWrite, formulonErrorCode, orFallback } from './result.js';
import { sheetHasView } from './sheet-kind.js';
import type { EngineCapabilities, Range } from './types.js';
import { SheetKind, SheetVisibility } from './types.js';
import type { WorkbookHandle } from './workbook-handle.js';
import {
  assertAlive,
  type WorkbookHandleInternals as SharedInternals,
  wb,
  ws,
} from './workbook-handle-internals.js';

type WorkbookHandleCtor = { prototype: WorkbookHandle };
type WorkbookHandleInternals = SharedInternals & {
  geometryPlatform?: WorkbookHandle['geometryPlatform'];
};

const CSS_PX_PER_POINT = 96 / 72;
const POINTS_PER_CSS_PX = 72 / 96;

/** Calibri 11's screen digit width and padding (96-DPI px), used when the
 *  engine reports no width model. */
const FALLBACK_DIGIT_WIDTH_PX = 7;
const FALLBACK_COLUMN_PADDING_PX = 5;

type LayoutCol = {
  first: number;
  last: number;
  width: number;
  hidden: boolean;
  outlineLevel: number;
  hasWidth?: boolean;
  collapsed?: boolean;
};

type LayoutRow = {
  row: number;
  height: number;
  hidden: boolean;
  outlineLevel: number;
  hasHeight?: boolean;
  customHeight?: boolean;
  collapsed?: boolean;
};

declare module './workbook-handle.js' {
  interface WorkbookHandle extends WorkbookHandleLayoutMethods {}
}

function internals(handle: unknown): WorkbookHandleInternals {
  return handle as WorkbookHandleInternals;
}

/** Whether rows and the row-number strip follow the Mac layout; column widths come from the
 *  engine's profile instead. */
function usesMacLayout(handle: unknown): boolean {
  return internals(handle).geometryPlatform === 'mac';
}

/** `sheet`'s display width model, or null when the engine cannot report it. */
function displayWidthModel(handle: unknown, sheet: number): WidthModel | null {
  if ((handle as WorkbookHandle).capabilities.geometry !== true) return null;
  return orFallback(null, () => ws(handle, sheet).getWidthModel(GeometryMode.Display));
}

/** `sheet`'s Normal font as the engine names it, or null when it cannot report it. */
function sheetNormalFont(handle: unknown, sheet: number): { name: string; size: number } | null {
  const model = displayWidthModel(handle, sheet);
  if (!model || !Number.isFinite(model.normalFontSizePt)) return null;
  return { name: model.normalFontName, size: model.normalFontSizePt };
}

/** `sheet`'s digit width and column padding in screen pixels, and the CSS px per screen pixel:
 *  1 at 96 DPI, 4/3 under a Mac profile, where the engine lays a pixel out as a point. */
function screenColumnMetrics(
  handle: unknown,
  sheet: number,
): { mdw: number; padding: number; cssPxPerScreenPx: number } {
  const model = displayWidthModel(handle, sheet);
  if (!model) {
    return {
      mdw: FALLBACK_DIGIT_WIDTH_PX,
      padding: FALLBACK_COLUMN_PADDING_PX,
      cssPxPerScreenPx: 1,
    };
  }
  const screenPxPerPoint = model.platform === WidthPlatform.Mac ? 1 : CSS_PX_PER_POINT;
  return {
    mdw: Math.round(model.pointsPerChar * screenPxPerPoint),
    padding: Math.round(model.paddingPt * screenPxPerPoint),
    cssPxPerScreenPx: CSS_PX_PER_POINT / screenPxPerPoint,
  };
}

function columnCharsToPx(handle: unknown, sheet: number, chars: number): number | null {
  if (!Number.isFinite(chars)) return null;
  return orFallback(null, () => {
    const pt = ws(handle, sheet).convertColumnCharsToPt(chars, GeometryMode.Display);
    return Number.isFinite(pt) ? pt * CSS_PX_PER_POINT : null;
  });
}

function columnPxToChars(handle: unknown, sheet: number, width: number): number | null {
  if (!Number.isFinite(width)) return null;
  return orFallback(null, () => {
    const chars = ws(handle, sheet).convertColumnPtToChars(
      width * POINTS_PER_CSS_PX,
      GeometryMode.Display,
    );
    return Number.isFinite(chars) ? chars : null;
  });
}

/** Map the engine's `<sheetView view>` mode to the store's name. */
function workbookViewOf(viewMode: SheetViewMode): WorkbookViewMode {
  if (viewMode === SheetViewMode.PageBreakPreview) return 'pageBreakPreview';
  if (viewMode === SheetViewMode.PageLayout) return 'pageLayout';
  return 'normal';
}

function engineViewModeOf(mode: WorkbookViewMode): SheetViewMode {
  if (mode === 'pageBreakPreview') return SheetViewMode.PageBreakPreview;
  if (mode === 'pageLayout') return SheetViewMode.PageLayout;
  return SheetViewMode.Normal;
}

export abstract class WorkbookHandleLayoutMethods {
  declare readonly capabilities: EngineCapabilities;

  /** Persist a column-width override on `[first, last]` for `sheet`.
   *  No-op (returns false) when the engine doesn't expose `setColumnWidth`
   *  — i.e. the stub fallback. The UI store stays the source of truth in
   *  that case so paint still reflects the drag. The public width is
   *  canonical CSS px under the handle's `geometryPlatform` column model;
   *  formulon's geometry-aware engines store the native character width. */
  setColumnWidth(sheet: number, first: number, last: number, width: number): boolean {
    assertAlive(this);
    if (!this.capabilities.colRowSize) return false;
    const nativeWidth =
      this.capabilities.geometry === true ? columnPxToChars(this, sheet, width) : width;
    if (nativeWidth === null) return false;
    return attemptWrite(() => ws(this, sheet).setColumnWidth(first, last, nativeWidth));
  }

  /** Persist a row-height override at `row` for `sheet`. See `setColumnWidth`
   *  for the no-op-on-stub rationale. The public height is CSS px; geometry
   *  aware engines store row heights in points. */
  setRowHeight(sheet: number, row: number, height: number): boolean {
    assertAlive(this);
    if (!this.capabilities.colRowSize) return false;
    const nativeHeight = this.capabilities.geometry === true ? height * POINTS_PER_CSS_PX : height;
    if (!Number.isFinite(nativeHeight)) return false;
    return attemptWrite(() => ws(this, sheet).setRowHeight(row, row, nativeHeight));
  }

  /** Remove explicit column-width overrides on `[first, last]` so the columns
   *  follow their sheet default. Returns false under the stub. */
  clearColumnWidth(sheet: number, first: number, last: number): boolean {
    assertAlive(this);
    if (this.capabilities.columnWidthClear !== true) return false;
    return attemptWrite(() => ws(this, sheet).resetColumnWidth(first, last));
  }

  /** Remove an explicit row-height override so the row follows its sheet
   *  default. Returns false under the stub. */
  clearRowHeight(sheet: number, row: number): boolean {
    assertAlive(this);
    if (this.capabilities.rowHeightClear !== true) return false;
    return attemptWrite(() => ws(this, sheet).resetRowHeight(row, row));
  }

  /** Snapshot of column overrides on `sheet`. Empty array under the stub.
   *  The returned objects own no engine memory — the underlying vector
   *  handle is released before this method returns. */
  getColumnLayouts(sheet: number): LayoutCol[] {
    assertAlive(this);
    if (!this.capabilities.colRowSize) return [];
    const out: LayoutCol[] = [];
    for (const e of orFallback([], () => ws(this, sheet).listColumnOverrides())) {
      const hasWidth = e.widthChars !== null;
      let width = 0;
      if (e.widthChars !== null) {
        const px =
          this.capabilities.geometry === true
            ? columnCharsToPx(this, sheet, e.widthChars)
            : e.widthChars;
        if (px === null) return [];
        width = px;
      }
      out.push({
        first: e.first,
        last: e.last,
        width,
        hidden: e.hidden,
        outlineLevel: e.outlineLevel,
        hasWidth,
        collapsed: e.collapsed,
      });
    }
    return out;
  }

  /** Return sheet defaults in the store's CSS-pixel units. Columns follow the
   *  `geometryPlatform` model: Mac derives an unstated default from
   *  `baseColWidth`, the default profile uses the engine's effective width.
   *  Mac rows take the Normal font's height when it is tabulated, unless
   *  `customHeight` pins the stated default; otherwise rows without a stated
   *  default use 102/7 points. On Mac `headerDigitAdvance` is the Normal
   *  font's unrounded digit advance in points, which sizes the row-number
   *  strip at every zoom; it is undefined elsewhere. Null means the engine
   *  lacks geometry support or a native read/conversion failed. */
  getSheetDefaultLayoutPx(sheet: number): {
    defaultColWidth: number;
    defaultRowHeight: number;
    headerDigitAdvance: number | undefined;
  } | null {
    assertAlive(this);
    if (this.capabilities.geometry !== true) return null;
    const defaults = orFallback(null, () => ws(this, sheet).getFormatDefaults());
    if (!defaults) return null;

    const colWidth = columnCharsToPx(
      this,
      sheet,
      defaults.defaultColumnWidthChars ?? defaults.effectiveDefaultColumnWidthChars,
    );
    const statedRowPoints = defaults.defaultRowHeightPt ?? 102 / 7;
    const pinned = defaults.defaultRowHeightPt !== null && defaults.customHeight;
    const font = usesMacLayout(this) && !pinned ? sheetNormalFont(this, sheet) : null;
    const rowPoints = (font && macDefaultRowPt(font.name, font.size)) ?? statedRowPoints;
    if (colWidth === null || !Number.isFinite(rowPoints)) return null;
    const rowHeight = rowPoints * CSS_PX_PER_POINT;
    if (!Number.isFinite(rowHeight)) return null;
    const normal = usesMacLayout(this) ? sheetNormalFont(this, sheet) : null;
    const headerDigitAdvance = normal ? macDigitAdvance(normal.name, normal.size) : undefined;
    return { defaultColWidth: colWidth, defaultRowHeight: rowHeight, headerDigitAdvance };
  }

  /** Width a column `widthPx` CSS px wide shows in the Column Width dialog,
   *  in characters of `sheet`'s Normal font. */
  columnDisplayChars(sheet: number, widthPx: number): number {
    assertAlive(this);
    const { mdw, padding, cssPxPerScreenPx } = screenColumnMetrics(this, sheet);
    return displayCharsForScreenPx(widthPx / cssPxPerScreenPx, mdw, padding);
  }

  /** CSS px of a column the Column Width dialog sets to `chars`, snapped to
   *  whole screen pixels. */
  columnPxForDisplayChars(sheet: number, chars: number): number {
    assertAlive(this);
    const { mdw, padding, cssPxPerScreenPx } = screenColumnMetrics(this, sheet);
    return screenPxForDisplayChars(chars, mdw, padding) * cssPxPerScreenPx;
  }

  /** Height a row `heightPx` CSS px tall shows in the Row Height dialog, in
   *  points, to the quarter point the desktop spreadsheet stores. */
  rowDisplayPt(heightPx: number): number {
    return Math.round(heightPx * POINTS_PER_CSS_PX * 4) / 4;
  }

  /** CSS px of a row the Row Height dialog sets to `pt`, snapped to a quarter
   *  point; painting snaps to screen pixels separately. */
  rowPxForDisplayPt(pt: number): number {
    return (Math.round(pt * 4) / 4) * CSS_PX_PER_POINT;
  }

  /** Persist frozen-pane counts on `sheet`. No-op (returns false) under stub. */
  setSheetFreeze(sheet: number, freezeRows: number, freezeCols: number): boolean {
    assertAlive(this);
    if (!this.capabilities.freeze) return false;
    return attemptWrite(() => ws(this, sheet).updateView({ freezeRows, freezeCols }));
  }

  /** Persist sheet zoom percentage (10..400, engine clamps). No-op under stub. */
  setSheetZoom(sheet: number, zoomScale: number): boolean {
    assertAlive(this);
    if (!this.capabilities.sheetZoom) return false;
    return attemptWrite(() => ws(this, sheet).updateView({ zoomPct: zoomScale }));
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
    return attemptWrite(() => {
      const sh = ws(this, sheet);
      const current = sh.getVisibility();
      if (hidden) {
        if (current === SheetVisibility.Visible) sh.setVisibility(SheetVisibility.Hidden);
      } else if (current !== SheetVisibility.Visible) {
        sh.setVisibility(SheetVisibility.Visible);
      }
    });
  }

  /** Set `sheet`'s tab to one of the three OOXML visibility states. Returns
   *  false when the engine only carries the two-state `setSheetTabHidden`,
   *  which cannot express `veryHidden`. */
  setSheetVisibility(sheet: number, visibility: SheetVisibility): boolean {
    assertAlive(this);
    if (!this.capabilities.sheetVisibility) return false;
    return attemptWrite(() => ws(this, sheet).setVisibility(visibility));
  }

  /** Set the hidden flag on `[first, last]` columns. No-op under stub or
   *  when the engine doesn't expose `setColumnHidden`. */
  setColumnHidden(sheet: number, first: number, last: number, hidden: boolean): boolean {
    assertAlive(this);
    if (!this.capabilities.hiddenRowsCols) return false;
    return attemptWrite(() => ws(this, sheet).setColumnHidden(first, last, hidden));
  }

  /** Set the hidden flag on `row`. */
  setRowHidden(sheet: number, row: number, hidden: boolean): boolean {
    assertAlive(this);
    if (!this.capabilities.hiddenRowsCols) return false;
    const ok = attemptWrite(() => ws(this, sheet).setRowHidden(row, row, hidden));
    if (ok) (this as unknown as WorkbookHandle).noteVisibilityChanged();
    return ok;
  }

  /** Set the outline level on `[first, last]` columns (0..7). */
  setColumnOutline(sheet: number, first: number, last: number, level: number): boolean {
    assertAlive(this);
    if (!this.capabilities.outlines) return false;
    return attemptWrite(() => ws(this, sheet).setColumnOutline(first, last, level));
  }

  /** Set the outline level on `row` (0..7). */
  setRowOutline(sheet: number, row: number, level: number): boolean {
    assertAlive(this);
    if (!this.capabilities.outlines) return false;
    return attemptWrite(() => ws(this, sheet).setRowOutline(row, row, level));
  }

  /** Set or clear the `collapsed` flag on the summary columns `[first, last]`. */
  setColumnCollapsed(sheet: number, first: number, last: number, collapsed: boolean): boolean {
    assertAlive(this);
    if (!this.capabilities.outlineCollapse) return false;
    return attemptWrite(() => ws(this, sheet).setColumnCollapsed(first, last, collapsed));
  }

  /** The type of `sheet`: a `SheetKind` ordinal. Every sheet reads as a worksheet when the engine
   *  does not report kinds (the stub). */
  sheetKind(sheet: number): SheetKind {
    assertAlive(this);
    if (!this.capabilities.sheetKind) return SheetKind.Worksheet;
    return orFallback(SheetKind.Worksheet, () => ws(this, sheet).kind ?? SheetKind.Worksheet);
  }

  /** Set or clear the `collapsed` flag on the summary row `row`. */
  setRowCollapsed(sheet: number, row: number, collapsed: boolean): boolean {
    assertAlive(this);
    if (!this.capabilities.outlineCollapse) return false;
    return attemptWrite(() => ws(this, sheet).setRowCollapsed(row, row, collapsed));
  }

  /** `sheet`'s tab state, which every sheet type carries; null when the engine cannot report it. */
  sheetVisibility(sheet: number): SheetVisibility | null {
    assertAlive(this);
    if (!this.capabilities.sheetView) return null;
    return orFallback(null, () => ws(this, sheet).getVisibility());
  }

  /** Snapshot of `sheet`'s view, including display flags. Returns null when
   *  the engine doesn't expose `getSheetView` (i.e. the stub or an older bundle) or
   *  `sheet` is neither a worksheet nor a chartsheet. */
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
    /** Absent when the engine cannot write the mode, so there is nothing to read back. */
    viewMode?: WorkbookViewMode;
    /** Scroll origin (`topLeftCell`); the frozen band's first cell when frozen. */
    topLeftRow?: number;
    topLeftCol?: number;
    /** First visible cell of the scrolling pane of a frozen or split view. */
    paneTopLeftRow?: number;
    paneTopLeftCol?: number;
    /** Absent when the engine cannot write the flag. */
    showFormulas?: boolean;
    showOutlineSymbols?: boolean;
    /** Unfrozen split position in screen twips (header included); 0 for none. */
    splitX?: number;
    splitY?: number;
  } | null {
    assertAlive(this);
    if (!this.capabilities.sheetView) return null;
    if (!sheetHasView(this, sheet)) return null;
    const sh = ws(this, sheet);
    const read = orFallback(null, () => ({ view: sh.getView(), visibility: sh.getVisibility() }));
    if (!read) return null;
    const { view } = read;
    const visibility: SheetVisibility = read.visibility;
    return {
      zoomScale: view.zoomPct,
      freezeRows: view.freezeRows,
      freezeCols: view.freezeCols,
      tabHidden: visibility !== SheetVisibility.Visible,
      visibility,
      showGridLines: view.showGridLines,
      showRowColHeaders: view.showRowColumnHeaders,
      showZeros: view.showZeros,
      rightToLeft: view.rightToLeft,
      viewMode: this.capabilities.sheetViewMode ? workbookViewOf(view.viewMode) : undefined,
      topLeftRow: view.topLeft.row,
      topLeftCol: view.topLeft.col,
      paneTopLeftRow: view.paneTopLeft.row,
      paneTopLeftCol: view.paneTopLeft.col,
      showFormulas: this.capabilities.sheetShowFormulas ? view.showFormulas : undefined,
      showOutlineSymbols: this.capabilities.outlineSymbols ? view.showOutlineSymbols : undefined,
      splitX: this.capabilities.sheetSplit ? view.splitXTwips : undefined,
      splitY: this.capabilities.sheetSplit ? view.splitYTwips : undefined,
    };
  }

  /** Persist the sheet's view mode (`<sheetView view="...">`). No-op under
   *  engines without the setter. */
  setSheetViewMode(sheet: number, mode: WorkbookViewMode): boolean {
    assertAlive(this);
    if (!this.capabilities.sheetViewMode) return false;
    return attemptWrite(() => ws(this, sheet).updateView({ viewMode: engineViewModeOf(mode) }));
  }

  /** Mark `sheet`'s tab as selected or not (`<sheetView tabSelected>`). */
  setSheetTabSelected(sheet: number, selected: boolean): boolean {
    assertAlive(this);
    if (!this.capabilities.sheetTabSelected) return false;
    return attemptWrite(() => ws(this, sheet).updateView({ tabSelected: selected }));
  }

  /** Index of the sheet the workbook opens on, or null when the engine has no
   *  active-sheet record. */
  activeSheet(): number | null {
    assertAlive(this);
    if (!this.capabilities.activeSheet) return null;
    return wb(this).activeSheet().index;
  }

  /** Make `sheet` the open-on sheet. Rejected for hidden or out-of-range sheets. */
  setActiveSheet(sheet: number): boolean {
    assertAlive(this);
    if (!this.capabilities.activeSheet) return false;
    return attemptWrite(() => wb(this).setActiveSheet(sheet));
  }

  /** The selection saved for `sheet`'s active pane; null without engine support. */
  getSheetSelection(sheet: number): {
    activeRow: number;
    activeCol: number;
    ranges: Range[];
  } | null {
    assertAlive(this);
    if (!this.capabilities.sheetSelection) return null;
    const r = orFallback(null, () => ws(this, sheet).getSelection());
    if (!r) return null;
    return {
      activeRow: r.active.row,
      activeCol: r.active.col,
      ranges: r.ranges.map((m) => rangeOfRect(sheet, m)),
    };
  }

  /** Replace the selection saved for `sheet`'s active pane. */
  setSheetSelection(
    sheet: number,
    selection: { activeRow: number; activeCol: number; ranges: readonly Range[] },
  ): boolean {
    assertAlive(this);
    if (!this.capabilities.sheetSelection) return false;
    return attemptWrite(() =>
      ws(this, sheet).setSelection({
        active: { row: selection.activeRow, col: selection.activeCol },
        ranges: selection.ranges.map(rectOfRange),
      }),
    );
  }

  /** Set the saved scroll origin (`<sheetView topLeftCell>`). */
  setSheetTopLeft(sheet: number, row: number, col: number): boolean {
    assertAlive(this);
    if (!this.capabilities.sheetScroll) return false;
    return attemptWrite(() => ws(this, sheet).updateView({ topLeft: { row, col } }));
  }

  /** Set the first visible cell of a frozen view's scrolling pane. */
  setSheetPaneTopLeft(sheet: number, row: number, col: number): boolean {
    assertAlive(this);
    if (!this.capabilities.sheetScroll) return false;
    return attemptWrite(() => ws(this, sheet).updateView({ paneTopLeft: { row, col } }));
  }

  setSheetShowGridLines(sheet: number, show: boolean): boolean {
    assertAlive(this);
    if (!this.capabilities.sheetViewFlags) return false;
    return attemptWrite(() => ws(this, sheet).updateView({ showGridLines: show }));
  }

  setSheetShowRowColHeaders(sheet: number, show: boolean): boolean {
    assertAlive(this);
    if (!this.capabilities.sheetViewFlags) return false;
    return attemptWrite(() => ws(this, sheet).updateView({ showRowColumnHeaders: show }));
  }

  setSheetShowZeros(sheet: number, show: boolean): boolean {
    assertAlive(this);
    if (!this.capabilities.sheetViewFlags) return false;
    return attemptWrite(() => ws(this, sheet).updateView({ showZeros: show }));
  }

  setSheetRightToLeft(sheet: number, rightToLeft: boolean): boolean {
    assertAlive(this);
    if (!this.capabilities.sheetViewFlags) return false;
    return attemptWrite(() => ws(this, sheet).updateView({ rightToLeft }));
  }

  setSheetShowFormulas(sheet: number, show: boolean): boolean {
    assertAlive(this);
    if (!this.capabilities.sheetShowFormulas) return false;
    return attemptWrite(() => ws(this, sheet).updateView({ showFormulas: show }));
  }

  setSheetShowOutlineSymbols(sheet: number, show: boolean): boolean {
    assertAlive(this);
    if (!this.capabilities.outlineSymbols) return false;
    return attemptWrite(() => ws(this, sheet).updateView({ showOutlineSymbols: show }));
  }

  /** Split the view at `xTwips` / `yTwips` of screen space, headers included;
   *  `0, 0` removes the split. Replaces any freeze. */
  setSheetSplit(sheet: number, xTwips: number, yTwips: number): boolean {
    assertAlive(this);
    if (!this.capabilities.sheetSplit) return false;
    return attemptWrite(() =>
      ws(this, sheet).updateView({
        splitXTwips: Math.round(xTwips),
        splitYTwips: Math.round(yTwips),
      }),
    );
  }

  /** Whether `sheet` hides every row that does not say otherwise (`zeroHeight`). */
  sheetRowsHiddenByDefault(sheet: number): boolean {
    assertAlive(this);
    if (this.capabilities.geometry !== true) return false;
    return orFallback(false, () => ws(this, sheet).getFormatDefaults().zeroHeight);
  }

  /** Merge `patch` into `sheet`'s `<sheetFormatPr>`; an omitted key keeps its value. */
  patchSheetFormatDefaults(sheet: number, patch: SheetFormatDefaultsPatch): boolean {
    assertAlive(this);
    if (!this.capabilities.sheetFormatDefaults) return false;
    return attemptWrite(() => ws(this, sheet).updateFormatDefaults(patch));
  }

  /** Insert `count` blank rows at `row` on `sheet`. The engine rewrites
   *  cross-workbook formula refs to follow the shift. Returns false on
   *  engines without `insertDeleteRowsCols`. NOT routed through the
   *  per-cell journal — callers wrap this in their own history entry. */
  engineInsertRows(sheet: number, row: number, count: number): boolean {
    assertAlive(this);
    if (!this.capabilities.insertDeleteRowsCols) return false;
    return attemptWrite(() => ws(this, sheet).insertRows(row, row + count - 1));
  }

  /** Delete `count` rows starting at `row` on `sheet`. Refs that fall
   *  inside the deleted interval collapse to `#REF!`. */
  engineDeleteRows(sheet: number, row: number, count: number): boolean {
    assertAlive(this);
    if (!this.capabilities.insertDeleteRowsCols) return false;
    return attemptWrite(() => ws(this, sheet).deleteRows(row, row + count - 1));
  }

  /** Insert `count` blank columns at `col` on `sheet`. */
  engineInsertCols(sheet: number, col: number, count: number): boolean {
    assertAlive(this);
    if (!this.capabilities.insertDeleteRowsCols) return false;
    return attemptWrite(() => ws(this, sheet).insertColumns(col, col + count - 1));
  }

  /** Delete `count` columns starting at `col` on `sheet`. */
  engineDeleteCols(sheet: number, col: number, count: number): boolean {
    assertAlive(this);
    if (!this.capabilities.insertDeleteRowsCols) return false;
    return attemptWrite(() => ws(this, sheet).deleteColumns(col, col + count - 1));
  }

  /** Insert (`Down` / `Right`) or delete (`Up` / `Left`) the cells of `range`, shifting the band. False when
   *  the engine refuses the shift (a table, multi-cell array, data table or PivotTable in the way, or a value
   *  pushed off the grid); the workbook is then unchanged. Other failures throw. */
  engineShiftCells(range: Range, shift: CellShift): boolean {
    assertAlive(this);
    const target = ws(this, range.sheet);
    const rect = rectOfRange(range);
    try {
      if (shift === CellShift.Down || shift === CellShift.Right) target.insertCells(rect, shift);
      else target.deleteCells(rect, shift);
      return true;
    } catch (e) {
      if (formulonErrorCode(e) === StatusCode.PreconditionFailed) return false;
      throw e;
    }
  }

  /** Append `range` as a merge on `sheet`. Returns false on engine failure or
   *  when `capabilities.merges` is off. The cell content inside the range is
   *  the caller's responsibility (spreadsheets keep top-left, blanks the rest). */
  engineAddMerge(sheet: number, range: Range): boolean {
    assertAlive(this);
    if (!this.capabilities.merges) return false;
    return attemptWrite(() => {
      ws(this, sheet).addMerge(rectOfRange(range));
    });
  }

  /** Remove every merge on `sheet` overlapping `range` (inclusive). No-op when
   *  nothing overlaps. Returns false on engine failure or capability off. */
  engineRemoveMerge(sheet: number, range: Range): boolean {
    assertAlive(this);
    if (!this.capabilities.merges) return false;
    const rect = rectOfRange(range);
    const sh = ws(this, sheet);
    // The engine reports NotFound for a range no merge overlaps.
    if (sh.getMergesInRange(rect).length === 0) return true;
    return attemptWrite(() => sh.removeMerge(rect));
  }

  /** Drop every merge on `sheet`. Returns false on engine failure or capability
   *  off. */
  engineClearMerges(sheet: number): boolean {
    assertAlive(this);
    if (!this.capabilities.merges) return false;
    return attemptWrite(() => ws(this, sheet).clearMerges());
  }

  /** Snapshot of every merge on `sheet` as inclusive `Range` records. Empty
   *  array under stub or when `capabilities.merges` is off. */
  getMerges(sheet: number): Range[] {
    assertAlive(this);
    if (!this.capabilities.merges) return [];
    return orFallback([], () => ws(this, sheet).listMerges()).map((m) => rangeOfRect(sheet, m));
  }

  /** Snapshot of row overrides on `sheet`. See `getColumnLayouts`. */
  getRowLayouts(sheet: number): LayoutRow[] {
    assertAlive(this);
    if (!this.capabilities.colRowSize) return [];
    const out: LayoutRow[] = [];
    for (const e of orFallback([], () => ws(this, sheet).listRowOverrides())) {
      const hasHeight = e.heightPt !== null;
      const points = e.heightPt ?? 0;
      const height = this.capabilities.geometry === true ? points * CSS_PX_PER_POINT : points;
      if (!Number.isFinite(height)) return [];
      out.push({
        row: e.row,
        height,
        hidden: e.hidden,
        outlineLevel: e.outlineLevel,
        hasHeight,
        customHeight: e.customHeight,
        collapsed: e.collapsed,
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
