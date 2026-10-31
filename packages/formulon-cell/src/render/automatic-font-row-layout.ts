import { associatedPaginationEngine } from '../commands/pagination.js';
import { MAX_COL, MAX_ROW, parseAddrKey } from '../engine/address.js';
import { fontPointsToCssPx } from '../engine/font-units.js';
import { nativeFormatEpoch } from '../engine/native-format-epoch.js';
import type {
  SheetAxisFormatOverride,
  SheetFontAxisSnapshot,
  SheetFontOverride,
} from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { MAX_ROW_HEIGHT_PX } from '../store/layout-limits.js';
import type { CellFormat, LayoutSlice, State } from '../store/store.js';

const projectedLayouts = new WeakSet<object>();
const projectionsByFormats = new WeakMap<object, Map<number, ProjectionCache>>();

interface ProjectionCache {
  readonly rawLayout: LayoutSlice;
  readonly workbook: WorkbookHandle;
  readonly mergesByCell: State['merges']['byCell'];
  readonly mergesByAnchor: State['merges']['byAnchor'];
  readonly sheetFontOverrides: State['format']['sheetFontOverrides'];
  readonly sheetAxisOverrides: State['format']['sheetAxisOverrides'];
  readonly sheetFontOverride: SheetFontOverride | undefined;
  readonly sheetAxisOverride: SheetAxisFormatOverride | undefined;
  readonly normalFontCss: number;
  readonly nativeFormatEpoch: number;
  readonly layout: LayoutSlice;
}

const validFontSize = (value: unknown): value is number =>
  typeof value === 'number' && Number.isFinite(value) && value > 0;

export const isAutomaticFontRowLayoutProjected = (layout: LayoutSlice): boolean =>
  projectedLayouts.has(layout);

export const markAutomaticFontRowLayoutProjected = (layout: LayoutSlice): void => {
  projectedLayouts.add(layout);
};

const lineHeight = (fontSizeCss: number): number | null => {
  const line = Math.round(fontSizeCss * 1.28);
  return Number.isFinite(line) ? line : null;
};

function fontSizeForStyleXf(
  workbook: WorkbookHandle,
  styleXf: number,
  normalFontCss: number,
  cache: Map<number, number | null>,
): number | null {
  if (cache.has(styleXf)) return cache.get(styleXf) ?? null;
  const xf = workbook.getCellXf(styleXf);
  const font = xf ? workbook.getResolvedFontRecord(xf.fontIndex) : null;
  const size = font && validFontSize(font.size) ? fontPointsToCssPx(font.size) : null;
  const result = validFontSize(size) ? size : styleXf === 0 ? normalFontCss : null;
  cache.set(styleXf, result);
  return result;
}

/** A sparse axis snapshot leaves ordinary columns implicit at style XF zero. */
function uniformColumnFontSize(
  snapshot: SheetFontAxisSnapshot,
  workbook: WorkbookHandle,
  normalFontCss: number,
  styleCache: Map<number, number | null>,
): number | null {
  const entries = [...snapshot.columns]
    .filter(
      (entry) =>
        Number.isInteger(entry.first) &&
        Number.isInteger(entry.last) &&
        entry.first <= entry.last &&
        entry.last >= 0 &&
        entry.first <= MAX_COL,
    )
    .map((entry) => ({
      first: Math.max(0, entry.first),
      last: Math.min(MAX_COL, entry.last),
      styleXf: entry.styleXf,
    }))
    .sort((a, b) => a.first - b.first || a.last - b.last);

  const sizes: number[] = [];
  const add = (styleXf: number): boolean => {
    const size = fontSizeForStyleXf(workbook, styleXf, normalFontCss, styleCache);
    if (size === null) return false;
    sizes.push(size);
    return true;
  };

  let cursor = 0;
  for (const entry of entries) {
    if (entry.first > cursor && !add(0)) return null;
    const first = Math.max(cursor, entry.first);
    if (first <= entry.last && !add(entry.styleXf)) return null;
    cursor = Math.max(cursor, entry.last + 1);
    if (cursor > MAX_COL) break;
  }
  if (cursor <= MAX_COL && !add(0)) return null;
  const firstSize = sizes[0];
  return firstSize !== undefined && sizes.every((size) => Math.abs(size - firstSize) < 1e-6)
    ? firstSize
    : null;
}

function getProjectionCache(
  state: State,
  workbook: WorkbookHandle,
  normalFontCss: number,
  currentNativeFormatEpoch: number,
): ProjectionCache | null {
  const sheet = state.data.sheetIndex;
  const formats = state.format.formats;
  const forSheet = projectionsByFormats.get(formats);
  const cached = forSheet?.get(sheet);
  const sheetFontOverride = state.format.sheetFontOverrides?.get(sheet);
  const sheetAxisOverride = state.format.sheetAxisOverrides?.get(sheet);
  if (
    cached &&
    cached.rawLayout === state.layout &&
    cached.workbook === workbook &&
    cached.mergesByCell === state.merges.byCell &&
    cached.mergesByAnchor === state.merges.byAnchor &&
    cached.sheetFontOverrides === state.format.sheetFontOverrides &&
    cached.sheetAxisOverrides === state.format.sheetAxisOverrides &&
    cached.sheetFontOverride === sheetFontOverride &&
    cached.sheetAxisOverride === sheetAxisOverride &&
    cached.normalFontCss === normalFontCss &&
    cached.nativeFormatEpoch === currentNativeFormatEpoch
  ) {
    return cached;
  }
  return null;
}

/**
 * Read-only display geometry for Excel's automatic row-height floor. The
 * workbook's native row records remain authoritative whenever they exist;
 * sparse font formats only affect implicit rows in this projected layout.
 */
export function automaticFontRowLayout(state: State): LayoutSlice {
  if (projectedLayouts.has(state.layout)) return state.layout;

  const workbook = associatedPaginationEngine(state);
  const sheet = state.data.sheetIndex;
  if (!workbook || !Number.isInteger(sheet)) return state.layout;

  const displayFont = workbook.workbookDisplayFont;
  if (!displayFont || !validFontSize(displayFont.size)) return state.layout;
  const normalFontCss = fontPointsToCssPx(displayFont.size);
  const currentNativeFormatEpoch = nativeFormatEpoch(workbook);
  const normalLine = lineHeight(normalFontCss);
  const rawDefaultHeight = state.layout.defaultRowHeight;
  if (normalLine === null || !Number.isFinite(rawDefaultHeight) || rawDefaultHeight <= 0) {
    return state.layout;
  }

  const cached = getProjectionCache(state, workbook, normalFontCss, currentNativeFormatEpoch);
  if (cached) return cached.layout;

  const styleCache = new Map<number, number | null>();
  const axisOverride = state.format.sheetAxisOverrides?.get(sheet);
  const axisSnapshot = axisOverride?.desired ??
    workbook.getSheetFontAxisSnapshot(sheet) ?? { columns: [], rows: [] };
  const patchSize = state.format.sheetFontOverrides?.get(sheet)?.patch.fontSize;
  const columnSize = uniformColumnFontSize(axisSnapshot, workbook, normalFontCss, styleCache);
  const globalFontCss = validFontSize(patchSize) ? patchSize : (columnSize ?? normalFontCss);
  if (lineHeight(globalFontCss) === null) return state.layout;

  const heightFor = (fontSizeCss: number): number | null => {
    const line = lineHeight(fontSizeCss);
    if (line === null) return null;
    return Math.min(MAX_ROW_HEIGHT_PX, rawDefaultHeight + Math.max(0, line - normalLine));
  };

  const defaultRowHeight = heightFor(globalFontCss);
  if (defaultRowHeight === null) return state.layout;

  const rowFonts = new Map<number, number>();
  const canProjectRow = (row: number): boolean =>
    Number.isInteger(row) && row >= 0 && row <= MAX_ROW && !state.layout.rowHeights.has(row);

  for (const entry of axisSnapshot.rows) {
    const size = fontSizeForStyleXf(workbook, entry.styleXf, normalFontCss, styleCache);
    if (size === null || !canProjectRow(entry.row)) continue;
    rowFonts.set(entry.row, Math.max(rowFonts.get(entry.row) ?? size, size));
  }

  const includeCellFont = (row: number, fontSizeCss: number): void => {
    if (!canProjectRow(row) || !validFontSize(fontSizeCss)) {
      return;
    }
    rowFonts.set(row, Math.max(rowFonts.get(row) ?? globalFontCss, fontSizeCss));
  };

  for (const [key, format] of state.format.formats) {
    const addr = parseAddrKey(key);
    if (
      !addr ||
      addr.sheet !== sheet ||
      addr.col < 0 ||
      addr.col > MAX_COL ||
      state.merges.byCell.has(key) ||
      state.merges.byAnchor.has(key)
    ) {
      continue;
    }
    const size = (format as CellFormat | undefined)?.fontSize;
    if (validFontSize(size)) includeCellFont(addr.row, size);
  }

  const rowHeights = new Map(state.layout.rowHeights);
  for (const [row, fontSizeCss] of rowFonts) {
    const height = heightFor(fontSizeCss);
    if (height !== null && height !== defaultRowHeight) rowHeights.set(row, height);
  }

  if (defaultRowHeight === rawDefaultHeight && rowHeights.size === state.layout.rowHeights.size) {
    const forSheet = projectionsByFormats.get(state.format.formats) ?? new Map();
    forSheet.set(sheet, {
      rawLayout: state.layout,
      workbook,
      mergesByCell: state.merges.byCell,
      mergesByAnchor: state.merges.byAnchor,
      sheetFontOverrides: state.format.sheetFontOverrides,
      sheetAxisOverrides: state.format.sheetAxisOverrides,
      sheetFontOverride: state.format.sheetFontOverrides?.get(sheet),
      sheetAxisOverride: state.format.sheetAxisOverrides?.get(sheet),
      normalFontCss,
      nativeFormatEpoch: currentNativeFormatEpoch,
      layout: state.layout,
    });
    projectionsByFormats.set(state.format.formats, forSheet);
    return state.layout;
  }

  const projected: LayoutSlice = {
    ...state.layout,
    defaultRowHeight,
    rowHeights:
      rowHeights.size === state.layout.rowHeights.size ? state.layout.rowHeights : rowHeights,
  };
  projectedLayouts.add(projected);
  const forSheet = projectionsByFormats.get(state.format.formats) ?? new Map();
  forSheet.set(sheet, {
    rawLayout: state.layout,
    workbook,
    mergesByCell: state.merges.byCell,
    mergesByAnchor: state.merges.byAnchor,
    sheetFontOverrides: state.format.sheetFontOverrides,
    sheetAxisOverrides: state.format.sheetAxisOverrides,
    sheetFontOverride: state.format.sheetFontOverrides?.get(sheet),
    sheetAxisOverride: state.format.sheetAxisOverrides?.get(sheet),
    normalFontCss,
    nativeFormatEpoch: currentNativeFormatEpoch,
    layout: projected,
  });
  projectionsByFormats.set(state.format.formats, forSheet);
  return projected;
}
