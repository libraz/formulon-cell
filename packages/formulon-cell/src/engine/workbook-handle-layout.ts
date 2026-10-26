import { GeometryMode, type SheetFormatDefaults, type WidthModelResult } from '@libraz/formulon';
import type { WorkbookViewMode } from '../store/types.js';
import { displayCharsForScreenPx, screenPxForDisplayChars } from './column-display-width.js';
import {
  MAC_COLUMN_PADDING_PT,
  macDefaultColumnPt,
  macMdw,
  macPtToRaw,
  macRawToPt,
} from './mac-column-model.js';
import { macDefaultRowPt } from './mac-row-model.js';
import type { EngineCapabilities, Range, Workbook } from './types.js';
import { SheetVisibility } from './types.js';
import type { WorkbookHandle } from './workbook-handle.js';

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

function optionalFlag(value: unknown): boolean | undefined {
  if (typeof value === 'boolean') return value;
  if (typeof value === 'number') return value !== 0;
  return undefined;
}

function usesMacColumnModel(handle: unknown): boolean {
  return internals(handle).geometryPlatform === 'mac';
}

/** `sheet`'s display width model, or null when the engine cannot report it. */
function displayWidthModel(handle: unknown, sheet: number): WidthModelResult | null {
  const getWidthModel = wb(handle).getWidthModel;
  if (typeof getWidthModel !== 'function') return null;
  try {
    const model = getWidthModel.call(wb(handle), sheet, GeometryMode.Display);
    return model.status.ok ? model : null;
  } catch {
    return null;
  }
}

/** `sheet`'s raw Normal font, or null when the engine cannot report it. */
function sheetNormalFont(handle: unknown, sheet: number): { name: string; size: number } | null {
  const model = displayWidthModel(handle, sheet);
  if (!model || !Number.isFinite(model.normalFontSize)) return null;
  return { name: model.normalFontName, size: model.normalFontSize };
}

/** CSS px per platform screen pixel: 1 at 96 DPI, 4/3 where Mac draws 1 pt as 1 px. */
function cssPxPerScreenPx(handle: unknown): number {
  return usesMacColumnModel(handle) ? CSS_PX_PER_POINT : 1;
}

/** `sheet`'s digit width and column padding in platform screen pixels. */
function screenColumnMetrics(handle: unknown, sheet: number): { mdw: number; padding: number } {
  if (usesMacColumnModel(handle)) {
    return {
      mdw: macSheetMdw(handle, sheet) ?? FALLBACK_DIGIT_WIDTH_PX,
      padding: MAC_COLUMN_PADDING_PT,
    };
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
    const result = wb(handle).columnCharsToPt(sheet, GeometryMode.Display, chars);
    if (!result.status.ok || !Number.isFinite(result.value)) return null;
    return result.value * CSS_PX_PER_POINT;
  } catch {
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
    const result = wb(handle).columnPtToChars(
      sheet,
      GeometryMode.Display,
      width * POINTS_PER_CSS_PX,
    );
    if (!result.status.ok || !Number.isFinite(result.value)) return null;
    return result.value;
  } catch {
    return null;
  }
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

/** Map the engine's `<sheetView view>` string; unknown values read as normal. */
function workbookViewOf(viewMode: string | undefined): WorkbookViewMode {
  return viewMode === 'pageBreakPreview' || viewMode === 'pageLayout' ? viewMode : 'normal';
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
    const s = wb(this).setColumnWidth(sheet, first, last, nativeWidth);
    return s.ok;
  }

  /** Persist a row-height override at `row` for `sheet`. See `setColumnWidth`
   *  for the no-op-on-stub rationale. The public height is CSS px; geometry
   *  aware engines store row heights in points. */
  setRowHeight(sheet: number, row: number, height: number): boolean {
    assertAlive(this);
    if (!this.capabilities.colRowSize) return false;
    const nativeHeight = this.capabilities.geometry === true ? height * POINTS_PER_CSS_PX : height;
    if (!Number.isFinite(nativeHeight)) return false;
    const s = wb(this).setRowHeight(sheet, row, nativeHeight);
    return s.ok;
  }

  /** Remove explicit column-width overrides on `[first, last]` so the columns
   *  follow their sheet default. Older engines have no clear operation and
   *  return false. */
  clearColumnWidth(sheet: number, first: number, last: number): boolean {
    assertAlive(this);
    if (this.capabilities.columnWidthClear !== true) return false;
    const clear = wb(this).clearColumnWidth;
    if (typeof clear !== 'function') return false;
    return clear.call(wb(this), sheet, first, last).ok;
  }

  /** Remove an explicit row-height override so the row follows its sheet
   *  default. Older engines have no clear operation and return false. */
  clearRowHeight(sheet: number, row: number): boolean {
    assertAlive(this);
    if (this.capabilities.rowHeightClear !== true) return false;
    const clear = wb(this).clearRowHeight;
    if (typeof clear !== 'function') return false;
    return clear.call(wb(this), sheet, row).ok;
  }

  /** Snapshot of column overrides on `sheet`. Empty array under the stub.
   *  The returned objects own no engine memory — the underlying vector
   *  handle is released before this method returns. */
  getColumnLayouts(sheet: number): LayoutCol[] {
    assertAlive(this);
    if (!this.capabilities.colRowSize) return [];
    const r = wb(this).getSheetColumns(sheet);
    const out: LayoutCol[] = [];
    if (!r.status.ok) return out;
    for (const e of r.columns) {
      const source = e as unknown as Record<string, unknown>;
      const width =
        this.capabilities.geometry === true ? columnCharsToPx(this, sheet, e.width) : e.width;
      if (width === null) return [];
      const hasWidth = optionalFlag(source.hasWidth);
      const entry: LayoutCol = {
        first: e.first,
        last: e.last,
        width,
        hidden: e.hidden !== 0,
        outlineLevel: e.outlineLevel,
      };
      if (hasWidth !== undefined) entry.hasWidth = hasWidth;
      const collapsed = optionalFlag(source.collapsed);
      if (collapsed !== undefined) entry.collapsed = collapsed;
      out.push(entry);
    }
    return out;
  }

  /** Return sheet defaults in the store's CSS-pixel units. Columns follow the
   *  `geometryPlatform` model: Mac derives an unstated default from
   *  `baseColWidth`, the default profile uses the engine's effective width.
   *  Mac rows take the Normal font's height when it is tabulated, unless
   *  `customHeight` pins the stated default; otherwise rows without a stated
   *  default use 102/7 points. Null means the engine
   *  lacks geometry support or a native read/conversion failed. */
  getSheetDefaultLayoutPx(
    sheet: number,
  ): { defaultColWidth: number; defaultRowHeight: number } | null {
    assertAlive(this);
    if (this.capabilities.geometry !== true) return null;
    let defaults: ReturnType<Workbook['getSheetFormatDefaults']>;
    try {
      defaults = wb(this).getSheetFormatDefaults(sheet);
    } catch {
      return null;
    }
    if (!defaults.status.ok) return null;

    let colWidth: number | null;
    if (usesMacColumnModel(this) && !defaults.hasDefaultColWidth) {
      const mdw = macSheetMdw(this, sheet);
      colWidth =
        mdw === null ? null : macDefaultColumnPt(defaults.baseColWidth, mdw) * CSS_PX_PER_POINT;
    } else {
      const chars = defaults.hasDefaultColWidth
        ? defaults.defaultColWidth
        : defaults.effectiveDefaultColWidth;
      colWidth = columnCharsToPx(this, sheet, chars);
    }
    const statedRowPoints = defaults.hasDefaultRowHeight ? defaults.defaultRowHeight : 102 / 7;
    const pinned = defaults.hasDefaultRowHeight && defaults.customHeight;
    const font = usesMacColumnModel(this) && !pinned ? sheetNormalFont(this, sheet) : null;
    const rowPoints = (font && macDefaultRowPt(font.name, font.size)) ?? statedRowPoints;
    if (colWidth === null || !Number.isFinite(rowPoints)) return null;
    const rowHeight = rowPoints * CSS_PX_PER_POINT;
    if (!Number.isFinite(rowHeight)) return null;
    return { defaultColWidth: colWidth, defaultRowHeight: rowHeight };
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

  /** Height a row `heightPx` CSS px tall shows in the Row Height dialog, in points. */
  rowDisplayPt(heightPx: number): number {
    return Math.round(heightPx * POINTS_PER_CSS_PX * 100) / 100;
  }

  /** CSS px of a row the Row Height dialog sets to `pt`, snapped to whole
   *  screen pixels. */
  rowPxForDisplayPt(pt: number): number {
    const perScreenPx = cssPxPerScreenPx(this);
    return Math.round((pt * CSS_PX_PER_POINT) / perScreenPx) * perScreenPx;
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

  /** Set or clear the `collapsed` flag on the summary columns `[first, last]`. */
  setColumnCollapsed(sheet: number, first: number, last: number, collapsed: boolean): boolean {
    assertAlive(this);
    if (!this.capabilities.outlineCollapse) return false;
    return wb(this).setColumnCollapsed(sheet, first, last, collapsed).ok;
  }

  /** Set or clear the `collapsed` flag on the summary row `row`. */
  setRowCollapsed(sheet: number, row: number, collapsed: boolean): boolean {
    assertAlive(this);
    if (!this.capabilities.outlineCollapse) return false;
    return wb(this).setRowCollapsed(sheet, row, collapsed).ok;
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
      viewMode: this.capabilities.sheetViewMode ? workbookViewOf(r.view.viewMode) : undefined,
      topLeftRow: r.view.topLeftRow,
      topLeftCol: r.view.topLeftCol,
      paneTopLeftRow: r.view.paneTopLeftRow,
      paneTopLeftCol: r.view.paneTopLeftCol,
      showFormulas: this.capabilities.sheetShowFormulas ? r.view.showFormulas !== 0 : undefined,
      showOutlineSymbols: this.capabilities.outlineSymbols
        ? r.view.showOutlineSymbols !== 0
        : undefined,
      splitX: this.capabilities.sheetSplit ? r.view.splitX : undefined,
      splitY: this.capabilities.sheetSplit ? r.view.splitY : undefined,
    };
  }

  /** Persist the sheet's view mode (`<sheetView view="...">`). No-op under
   *  engines without the setter. */
  setSheetViewMode(sheet: number, mode: WorkbookViewMode): boolean {
    assertAlive(this);
    if (!this.capabilities.sheetViewMode) return false;
    return wb(this).setSheetViewMode(sheet, mode === 'normal' ? '' : mode).ok;
  }

  /** Mark `sheet`'s tab as selected or not (`<sheetView tabSelected>`). */
  setSheetTabSelected(sheet: number, selected: boolean): boolean {
    assertAlive(this);
    if (!this.capabilities.sheetTabSelected) return false;
    return wb(this).setSheetTabSelected(sheet, selected).ok;
  }

  /** Index of the sheet the workbook opens on, or null when the engine has no
   *  active-sheet record. */
  activeSheet(): number | null {
    assertAlive(this);
    if (!this.capabilities.activeSheet) return null;
    const r = wb(this).activeSheet();
    return r.status.ok ? r.index : null;
  }

  /** Make `sheet` the open-on sheet. Rejected for hidden or out-of-range sheets. */
  setActiveSheet(sheet: number): boolean {
    assertAlive(this);
    if (!this.capabilities.activeSheet) return false;
    return wb(this).setActiveSheet(sheet).ok;
  }

  /** The selection saved for `sheet`'s active pane; null without engine support. */
  getSheetSelection(sheet: number): {
    activeRow: number;
    activeCol: number;
    ranges: Range[];
  } | null {
    assertAlive(this);
    if (!this.capabilities.sheetSelection) return null;
    const r = wb(this).getSheetSelection(sheet);
    if (!r.status.ok) return null;
    return {
      activeRow: r.activeRow,
      activeCol: r.activeCol,
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
    return wb(this).setSheetSelection(sheet, {
      activeRow: selection.activeRow,
      activeCol: selection.activeCol,
      ranges: selection.ranges.map((r) => ({
        firstRow: r.r0,
        firstCol: r.c0,
        lastRow: r.r1,
        lastCol: r.c1,
      })),
    }).ok;
  }

  /** Set the saved scroll origin (`<sheetView topLeftCell>`). */
  setSheetTopLeft(sheet: number, row: number, col: number): boolean {
    assertAlive(this);
    if (!this.capabilities.sheetScroll) return false;
    return wb(this).setSheetTopLeft(sheet, row, col).ok;
  }

  /** Set the first visible cell of a frozen view's scrolling pane. */
  setSheetPaneTopLeft(sheet: number, row: number, col: number): boolean {
    assertAlive(this);
    if (!this.capabilities.sheetScroll) return false;
    return wb(this).setSheetPaneTopLeft(sheet, row, col).ok;
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

  setSheetShowFormulas(sheet: number, show: boolean): boolean {
    assertAlive(this);
    if (!this.capabilities.sheetShowFormulas) return false;
    return wb(this).setSheetShowFormulas(sheet, show).ok;
  }

  setSheetShowOutlineSymbols(sheet: number, show: boolean): boolean {
    assertAlive(this);
    if (!this.capabilities.outlineSymbols) return false;
    return wb(this).setSheetShowOutlineSymbols(sheet, show).ok;
  }

  /** Split the view at `xTwips` / `yTwips` of screen space, headers included;
   *  `0, 0` removes the split. Replaces any freeze. */
  setSheetSplit(sheet: number, xTwips: number, yTwips: number): boolean {
    assertAlive(this);
    if (!this.capabilities.sheetSplit) return false;
    return wb(this).setSheetSplit(sheet, Math.round(xTwips), Math.round(yTwips)).ok;
  }

  /** Whether `sheet` hides every row that does not say otherwise (`zeroHeight`). */
  sheetRowsHiddenByDefault(sheet: number): boolean {
    assertAlive(this);
    if (this.capabilities.geometry !== true) return false;
    try {
      const defaults = wb(this).getSheetFormatDefaults(sheet);
      return defaults.status.ok && defaults.zeroHeight === true;
    } catch {
      return false;
    }
  }

  /** Restate `sheet`'s `<sheetFormatPr>` with `patch` applied. The engine resets
   *  every width key a call leaves out, so the stored set is always passed whole. */
  patchSheetFormatDefaults(sheet: number, patch: Partial<SheetFormatDefaults>): boolean {
    assertAlive(this);
    if (!this.capabilities.sheetFormatDefaults) return false;
    const current = wb(this).getSheetFormatDefaults(sheet);
    if (!current.status.ok) return false;
    const full: SheetFormatDefaults = {
      defaultColWidth: current.defaultColWidth,
      defaultRowHeight: current.defaultRowHeight,
      baseColWidth: current.baseColWidth,
      hasDefaultColWidth: current.hasDefaultColWidth,
      hasDefaultRowHeight: current.hasDefaultRowHeight,
      customHeight: current.customHeight,
      zeroHeight: current.zeroHeight,
      ...patch,
    };
    return wb(this).setSheetFormatDefaults(sheet, full).ok;
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
  getRowLayouts(sheet: number): LayoutRow[] {
    assertAlive(this);
    if (!this.capabilities.colRowSize) return [];
    const r = wb(this).getSheetRowOverrides(sheet);
    const out: LayoutRow[] = [];
    if (!r.status.ok) return out;
    for (const e of r.rows) {
      const source = e as unknown as Record<string, unknown>;
      const height = this.capabilities.geometry === true ? e.height * CSS_PX_PER_POINT : e.height;
      if (!Number.isFinite(height)) return [];
      const hasHeight = optionalFlag(source.hasHeight);
      const customHeight = optionalFlag(source.customHeight);
      const entry: LayoutRow = {
        row: e.row,
        height,
        hidden: e.hidden !== 0,
        outlineLevel: e.outlineLevel,
      };
      if (hasHeight !== undefined) entry.hasHeight = hasHeight;
      if (customHeight !== undefined) entry.customHeight = customHeight;
      const collapsed = optionalFlag(source.collapsed);
      if (collapsed !== undefined) entry.collapsed = collapsed;
      out.push(entry);
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
