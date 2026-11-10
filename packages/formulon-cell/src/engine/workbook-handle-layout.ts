import {
  CellShift,
  GeometryMode,
  type Sheet,
  type SheetFormatDefaults,
  type SheetFormatDefaultsPatch,
  SheetViewMode,
  StatusCode,
  type WidthModel,
} from '@libraz/formulon';
import type { WorkbookViewMode } from '../store/types.js';
import { displayCharsForScreenPx, screenPxForDisplayChars } from './column-display-width.js';
import {
  macColumnPaddingPt,
  macDefaultColumnPt,
  macDigitAdvance,
  macMdw,
  macPtToRaw,
  macRawToPt,
} from './mac-column-model.js';
import { macDefaultRowPt } from './mac-row-model.js';
import { formulonErrorCode, isFormulonError, orFallback } from './result.js';
import type { EngineCapabilities, Range, Workbook } from './types.js';
import { SheetVisibility } from './types.js';
import type { WorkbookHandle } from './workbook-handle.js';
import { shiftPivotOwnedAreas } from './workbook-handle-pivot.js';

type WorkbookHandleCtor = { prototype: WorkbookHandle };
type WorkbookHandleInternals = {
  wb: Workbook;
  geometryPlatform?: WorkbookHandle['geometryPlatform'];
  assertAlive(): void;
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

function assertAlive(handle: unknown): void {
  internals(handle).assertAlive();
}

function wb(handle: unknown): Workbook {
  return internals(handle).wb;
}

function ws(handle: unknown, sheet: number): Sheet {
  return wb(handle).sheet(sheet);
}

/** Run an engine call; an engine refusal is a soft `false`, anything else rethrows. */
function attempt(fn: () => void): boolean {
  try {
    fn();
    return true;
  } catch (e) {
    if (!isFormulonError(e)) throw e;
    return false;
  }
}

function usesMacColumnModel(handle: unknown): boolean {
  return internals(handle).geometryPlatform === 'mac';
}

/** `sheet`'s display width model, or null when the engine cannot report it. */
function displayWidthModel(handle: unknown, sheet: number): WidthModel | null {
  if ((handle as WorkbookHandle).capabilities.geometry !== true) return null;
  try {
    return ws(handle, sheet).getWidthModel(GeometryMode.Display);
  } catch (e) {
    if (!isFormulonError(e)) throw e;
    return null;
  }
}

/** `sheet`'s Normal font as displayed — a theme-linked record takes the face
 *  the theme resolves it to — or null when the engine cannot report it. */
function sheetNormalFont(handle: unknown, sheet: number): { name: string; size: number } | null {
  const model = displayWidthModel(handle, sheet);
  if (!model || !Number.isFinite(model.normalFontSizePt)) return null;
  const normal = (handle as WorkbookHandle).getResolvedFontRecord(0);
  const name = normal?.scheme ? normal.name : model.normalFontName;
  return { name, size: model.normalFontSizePt };
}

/** CSS px per platform screen pixel: 1 at 96 DPI, 4/3 where Mac draws 1 pt as 1 px. */
function cssPxPerScreenPx(handle: unknown): number {
  return usesMacColumnModel(handle) ? CSS_PX_PER_POINT : 1;
}

/** `sheet`'s digit width and column padding in platform screen pixels. */
function screenColumnMetrics(handle: unknown, sheet: number): { mdw: number; padding: number } {
  if (usesMacColumnModel(handle)) {
    const mdw = macSheetMdw(handle, sheet) ?? FALLBACK_DIGIT_WIDTH_PX;
    return { mdw, padding: macColumnPaddingPt(mdw) };
  }
  const model = displayWidthModel(handle, sheet);
  if (!model) return { mdw: FALLBACK_DIGIT_WIDTH_PX, padding: FALLBACK_COLUMN_PADDING_PX };
  return {
    mdw: Math.round(model.pointsPerChar * CSS_PX_PER_POINT),
    padding: Math.round(model.paddingPt * CSS_PX_PER_POINT),
  };
}

/** Mac maximum digit width of `sheet`'s raw Normal font, or null when the
 *  engine cannot report it. */
function macSheetMdw(handle: unknown, sheet: number): number | null {
  const font = sheetNormalFont(handle, sheet);
  if (!font) return null;
  const mdw = macMdw(font.name, font.size);
  return mdw > 0 ? mdw : null;
}

function columnCharsToPx(handle: unknown, sheet: number, chars: number): number | null {
  if (!Number.isFinite(chars)) return null;
  if (usesMacColumnModel(handle)) {
    const mdw = macSheetMdw(handle, sheet);
    return mdw === null ? null : macRawToPt(chars, mdw) * CSS_PX_PER_POINT;
  }
  try {
    const pt = ws(handle, sheet).convertColumnCharsToPt(chars, GeometryMode.Display);
    return Number.isFinite(pt) ? pt * CSS_PX_PER_POINT : null;
  } catch (e) {
    if (!isFormulonError(e)) throw e;
    return null;
  }
}

function columnPxToChars(handle: unknown, sheet: number, width: number): number | null {
  if (!Number.isFinite(width)) return null;
  if (usesMacColumnModel(handle)) {
    const mdw = macSheetMdw(handle, sheet);
    return mdw === null ? null : macPtToRaw(width * POINTS_PER_CSS_PX, mdw);
  }
  try {
    const chars = ws(handle, sheet).convertColumnPtToChars(
      width * POINTS_PER_CSS_PX,
      GeometryMode.Display,
    );
    return Number.isFinite(chars) ? chars : null;
  } catch (e) {
    if (!isFormulonError(e)) throw e;
    return null;
  }
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
    return attempt(() => ws(this, sheet).setColumnWidth(first, last, nativeWidth));
  }

  /** Persist a row-height override at `row` for `sheet`. See `setColumnWidth`
   *  for the no-op-on-stub rationale. The public height is CSS px; geometry
   *  aware engines store row heights in points. */
  setRowHeight(sheet: number, row: number, height: number): boolean {
    assertAlive(this);
    if (!this.capabilities.colRowSize) return false;
    const nativeHeight = this.capabilities.geometry === true ? height * POINTS_PER_CSS_PX : height;
    if (!Number.isFinite(nativeHeight)) return false;
    return attempt(() => ws(this, sheet).setRowHeight(row, row, nativeHeight));
  }

  /** Remove explicit column-width overrides on `[first, last]` so the columns
   *  follow their sheet default. Returns false under the stub. */
  clearColumnWidth(sheet: number, first: number, last: number): boolean {
    assertAlive(this);
    if (this.capabilities.columnWidthClear !== true) return false;
    return attempt(() => ws(this, sheet).resetColumnWidth(first, last));
  }

  /** Remove an explicit row-height override so the row follows its sheet
   *  default. Returns false under the stub. */
  clearRowHeight(sheet: number, row: number): boolean {
    assertAlive(this);
    if (this.capabilities.rowHeightClear !== true) return false;
    return attempt(() => ws(this, sheet).resetRowHeight(row, row));
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
    let defaults: SheetFormatDefaults;
    try {
      defaults = ws(this, sheet).getFormatDefaults();
    } catch (e) {
      if (!isFormulonError(e)) throw e;
      return null;
    }

    let colWidth: number | null;
    if (usesMacColumnModel(this) && defaults.defaultColumnWidthChars === null) {
      const mdw = macSheetMdw(this, sheet);
      colWidth =
        mdw === null
          ? null
          : macDefaultColumnPt(defaults.baseColumnWidthChars, mdw) * CSS_PX_PER_POINT;
    } else {
      const chars = defaults.defaultColumnWidthChars ?? defaults.effectiveDefaultColumnWidthChars;
      colWidth = columnCharsToPx(this, sheet, chars);
    }
    const statedRowPoints = defaults.defaultRowHeightPt ?? 102 / 7;
    const pinned = defaults.defaultRowHeightPt !== null && defaults.customHeight;
    const font = usesMacColumnModel(this) && !pinned ? sheetNormalFont(this, sheet) : null;
    const rowPoints = (font && macDefaultRowPt(font.name, font.size)) ?? statedRowPoints;
    if (colWidth === null || !Number.isFinite(rowPoints)) return null;
    const rowHeight = rowPoints * CSS_PX_PER_POINT;
    if (!Number.isFinite(rowHeight)) return null;
    const normal = usesMacColumnModel(this) ? sheetNormalFont(this, sheet) : null;
    const headerDigitAdvance = normal ? macDigitAdvance(normal.name, normal.size) : undefined;
    return { defaultColWidth: colWidth, defaultRowHeight: rowHeight, headerDigitAdvance };
  }

  /** Width a column `widthPx` CSS px wide shows in the Column Width dialog,
   *  in characters of `sheet`'s Normal font. */
  columnDisplayChars(sheet: number, widthPx: number): number {
    assertAlive(this);
    const { mdw, padding } = screenColumnMetrics(this, sheet);
    return displayCharsForScreenPx(widthPx / cssPxPerScreenPx(this), mdw, padding);
  }

  /** CSS px of a column the Column Width dialog sets to `chars`, snapped to
   *  whole screen pixels. */
  columnPxForDisplayChars(sheet: number, chars: number): number {
    assertAlive(this);
    const { mdw, padding } = screenColumnMetrics(this, sheet);
    return screenPxForDisplayChars(chars, mdw, padding) * cssPxPerScreenPx(this);
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
    return attempt(() => ws(this, sheet).updateView({ freezeRows, freezeCols }));
  }

  /** Persist sheet zoom percentage (10..400, engine clamps). No-op under stub. */
  setSheetZoom(sheet: number, zoomScale: number): boolean {
    assertAlive(this);
    if (!this.capabilities.sheetZoom) return false;
    return attempt(() => ws(this, sheet).updateView({ zoomPct: zoomScale }));
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
    return attempt(() => {
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
    return attempt(() => ws(this, sheet).setVisibility(visibility));
  }

  /** Set the hidden flag on `[first, last]` columns. No-op under stub or
   *  when the engine doesn't expose `setColumnHidden`. */
  setColumnHidden(sheet: number, first: number, last: number, hidden: boolean): boolean {
    assertAlive(this);
    if (!this.capabilities.hiddenRowsCols) return false;
    return attempt(() => ws(this, sheet).setColumnHidden(first, last, hidden));
  }

  /** Set the hidden flag on `row`. */
  setRowHidden(sheet: number, row: number, hidden: boolean): boolean {
    assertAlive(this);
    if (!this.capabilities.hiddenRowsCols) return false;
    const ok = attempt(() => ws(this, sheet).setRowHidden(row, row, hidden));
    if (ok) (this as unknown as WorkbookHandle).noteVisibilityChanged();
    return ok;
  }

  /** Set the outline level on `[first, last]` columns (0..7). */
  setColumnOutline(sheet: number, first: number, last: number, level: number): boolean {
    assertAlive(this);
    if (!this.capabilities.outlines) return false;
    return attempt(() => ws(this, sheet).setColumnOutline(first, last, level));
  }

  /** Set the outline level on `row` (0..7). */
  setRowOutline(sheet: number, row: number, level: number): boolean {
    assertAlive(this);
    if (!this.capabilities.outlines) return false;
    return attempt(() => ws(this, sheet).setRowOutline(row, row, level));
  }

  /** Set or clear the `collapsed` flag on the summary columns `[first, last]`. */
  setColumnCollapsed(sheet: number, first: number, last: number, collapsed: boolean): boolean {
    assertAlive(this);
    if (!this.capabilities.outlineCollapse) return false;
    return attempt(() => ws(this, sheet).setColumnCollapsed(first, last, collapsed));
  }

  /** Set or clear the `collapsed` flag on the summary row `row`. */
  setRowCollapsed(sheet: number, row: number, collapsed: boolean): boolean {
    assertAlive(this);
    if (!this.capabilities.outlineCollapse) return false;
    return attempt(() => ws(this, sheet).setRowCollapsed(row, row, collapsed));
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
    return attempt(() => ws(this, sheet).updateView({ viewMode: engineViewModeOf(mode) }));
  }

  /** Mark `sheet`'s tab as selected or not (`<sheetView tabSelected>`). */
  setSheetTabSelected(sheet: number, selected: boolean): boolean {
    assertAlive(this);
    if (!this.capabilities.sheetTabSelected) return false;
    return attempt(() => ws(this, sheet).updateView({ tabSelected: selected }));
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
    return attempt(() => wb(this).setActiveSheet(sheet));
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
      ranges: r.ranges.map((m) => ({
        sheet,
        r0: m.firstRow,
        c0: m.firstCol,
        r1: m.lastRow,
        c1: m.lastCol,
      })),
    };
  }

  /** Replace the selection saved for `sheet`'s active pane. */
  setSheetSelection(
    sheet: number,
    selection: { activeRow: number; activeCol: number; ranges: readonly Range[] },
  ): boolean {
    assertAlive(this);
    if (!this.capabilities.sheetSelection) return false;
    return attempt(() =>
      ws(this, sheet).setSelection({
        active: { row: selection.activeRow, col: selection.activeCol },
        ranges: selection.ranges.map((r) => ({
          firstRow: r.r0,
          firstCol: r.c0,
          lastRow: r.r1,
          lastCol: r.c1,
        })),
      }),
    );
  }

  /** Set the saved scroll origin (`<sheetView topLeftCell>`). */
  setSheetTopLeft(sheet: number, row: number, col: number): boolean {
    assertAlive(this);
    if (!this.capabilities.sheetScroll) return false;
    return attempt(() => ws(this, sheet).updateView({ topLeft: { row, col } }));
  }

  /** Set the first visible cell of a frozen view's scrolling pane. */
  setSheetPaneTopLeft(sheet: number, row: number, col: number): boolean {
    assertAlive(this);
    if (!this.capabilities.sheetScroll) return false;
    return attempt(() => ws(this, sheet).updateView({ paneTopLeft: { row, col } }));
  }

  setSheetShowGridLines(sheet: number, show: boolean): boolean {
    assertAlive(this);
    if (!this.capabilities.sheetViewFlags) return false;
    return attempt(() => ws(this, sheet).updateView({ showGridLines: show }));
  }

  setSheetShowRowColHeaders(sheet: number, show: boolean): boolean {
    assertAlive(this);
    if (!this.capabilities.sheetViewFlags) return false;
    return attempt(() => ws(this, sheet).updateView({ showRowColumnHeaders: show }));
  }

  setSheetShowZeros(sheet: number, show: boolean): boolean {
    assertAlive(this);
    if (!this.capabilities.sheetViewFlags) return false;
    return attempt(() => ws(this, sheet).updateView({ showZeros: show }));
  }

  setSheetRightToLeft(sheet: number, rightToLeft: boolean): boolean {
    assertAlive(this);
    if (!this.capabilities.sheetViewFlags) return false;
    return attempt(() => ws(this, sheet).updateView({ rightToLeft }));
  }

  setSheetShowFormulas(sheet: number, show: boolean): boolean {
    assertAlive(this);
    if (!this.capabilities.sheetShowFormulas) return false;
    return attempt(() => ws(this, sheet).updateView({ showFormulas: show }));
  }

  setSheetShowOutlineSymbols(sheet: number, show: boolean): boolean {
    assertAlive(this);
    if (!this.capabilities.outlineSymbols) return false;
    return attempt(() => ws(this, sheet).updateView({ showOutlineSymbols: show }));
  }

  /** Split the view at `xTwips` / `yTwips` of screen space, headers included;
   *  `0, 0` removes the split. Replaces any freeze. */
  setSheetSplit(sheet: number, xTwips: number, yTwips: number): boolean {
    assertAlive(this);
    if (!this.capabilities.sheetSplit) return false;
    return attempt(() =>
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
    try {
      return ws(this, sheet).getFormatDefaults().zeroHeight;
    } catch (e) {
      if (!isFormulonError(e)) throw e;
      return false;
    }
  }

  /** Merge `patch` into `sheet`'s `<sheetFormatPr>`; an omitted key keeps its value. */
  patchSheetFormatDefaults(sheet: number, patch: SheetFormatDefaultsPatch): boolean {
    assertAlive(this);
    if (!this.capabilities.sheetFormatDefaults) return false;
    return attempt(() => ws(this, sheet).updateFormatDefaults(patch));
  }

  /** Insert `count` blank rows at `row` on `sheet`. The engine rewrites
   *  cross-workbook formula refs to follow the shift. Returns false on
   *  engines without `insertDeleteRowsCols`. NOT routed through the
   *  per-cell journal — callers wrap this in their own history entry. */
  engineInsertRows(sheet: number, row: number, count: number): boolean {
    assertAlive(this);
    if (!this.capabilities.insertDeleteRowsCols) return false;
    const ok = attempt(() => ws(this, sheet).insertRows(row, row + count - 1));
    if (ok) shiftPivotOwnedAreas(this, sheet, 'row', 'insert', row, count);
    return ok;
  }

  /** Delete `count` rows starting at `row` on `sheet`. Refs that fall
   *  inside the deleted interval collapse to `#REF!`. */
  engineDeleteRows(sheet: number, row: number, count: number): boolean {
    assertAlive(this);
    if (!this.capabilities.insertDeleteRowsCols) return false;
    const ok = attempt(() => ws(this, sheet).deleteRows(row, row + count - 1));
    if (ok) shiftPivotOwnedAreas(this, sheet, 'row', 'delete', row, count);
    return ok;
  }

  /** Insert `count` blank columns at `col` on `sheet`. */
  engineInsertCols(sheet: number, col: number, count: number): boolean {
    assertAlive(this);
    if (!this.capabilities.insertDeleteRowsCols) return false;
    const ok = attempt(() => ws(this, sheet).insertColumns(col, col + count - 1));
    if (ok) shiftPivotOwnedAreas(this, sheet, 'col', 'insert', col, count);
    return ok;
  }

  /** Delete `count` columns starting at `col` on `sheet`. */
  engineDeleteCols(sheet: number, col: number, count: number): boolean {
    assertAlive(this);
    if (!this.capabilities.insertDeleteRowsCols) return false;
    const ok = attempt(() => ws(this, sheet).deleteColumns(col, col + count - 1));
    if (ok) shiftPivotOwnedAreas(this, sheet, 'col', 'delete', col, count);
    return ok;
  }

  /** Insert (`Down` / `Right`) or delete (`Up` / `Left`) the cells of `range`, shifting the band. False when
   *  the engine refuses the shift (a table, multi-cell array, data table or PivotTable in the way, or a value
   *  pushed off the grid); the workbook is then unchanged. Other failures throw. */
  engineShiftCells(range: Range, shift: CellShift): boolean {
    assertAlive(this);
    const target = ws(this, range.sheet);
    const rect = { firstRow: range.r0, firstCol: range.c0, lastRow: range.r1, lastCol: range.c1 };
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
    return attempt(() => {
      ws(this, sheet).addMerge({
        firstRow: range.r0,
        firstCol: range.c0,
        lastRow: range.r1,
        lastCol: range.c1,
      });
    });
  }

  /** Remove every merge on `sheet` overlapping `range` (inclusive). No-op when
   *  nothing overlaps. Returns false on engine failure or capability off. */
  engineRemoveMerge(sheet: number, range: Range): boolean {
    assertAlive(this);
    if (!this.capabilities.merges) return false;
    const rect = {
      firstRow: range.r0,
      firstCol: range.c0,
      lastRow: range.r1,
      lastCol: range.c1,
    };
    const sh = ws(this, sheet);
    // The engine reports NotFound for a range no merge overlaps.
    if (sh.getMergesInRange(rect).length === 0) return true;
    return attempt(() => sh.removeMerge(rect));
  }

  /** Drop every merge on `sheet`. Returns false on engine failure or capability
   *  off. */
  engineClearMerges(sheet: number): boolean {
    assertAlive(this);
    if (!this.capabilities.merges) return false;
    return attempt(() => ws(this, sheet).clearMerges());
  }

  /** Snapshot of every merge on `sheet` as inclusive `Range` records. Empty
   *  array under stub or when `capabilities.merges` is off. */
  getMerges(sheet: number): Range[] {
    assertAlive(this);
    if (!this.capabilities.merges) return [];
    return orFallback([], () => ws(this, sheet).listMerges()).map((m) => ({
      sheet,
      r0: m.firstRow,
      c0: m.firstCol,
      r1: m.lastRow,
      c1: m.lastCol,
    }));
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
