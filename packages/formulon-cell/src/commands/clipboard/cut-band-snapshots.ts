import { addrKey, parseAddrKey } from '../../engine/address.js';
import { flushFormatToEngine } from '../../engine/cell-format-sync.js';
import type { Addr, Range } from '../../engine/types.js';
import type { WorkbookHandle } from '../../engine/workbook-handle.js';
import { mutators, type SpreadsheetStore, type State } from '../../store/store.js';
import type { History } from '../history.js';
import type { ClipboardSnapshot } from './snapshot.js';

interface SourceDimensionSnapshot {
  colWidths: Map<number, number>;
  rowHeights: Map<number, number>;
  hiddenCols: Set<number>;
  hiddenRows: Set<number>;
  outlineCols: Map<number, number>;
  outlineRows: Map<number, number>;
}

interface AxisVisibilitySnapshot {
  hiddenCols: Set<number>;
  hiddenRows: Set<number>;
  outlineCols: Map<number, number>;
  outlineRows: Map<number, number>;
}

export const emptyAxisVisibility = (): AxisVisibilitySnapshot => ({
  hiddenCols: new Set(),
  hiddenRows: new Set(),
  outlineCols: new Map(),
  outlineRows: new Map(),
});

export function captureStoreAxisVisibility(
  store: SpreadsheetStore,
  range: Range,
  wholeRows: boolean,
): AxisVisibilitySnapshot {
  const layout = store.getState().layout;
  const visibility = emptyAxisVisibility();
  if (wholeRows) {
    for (const row of layout.hiddenRows) {
      if (row >= range.r0 && row <= range.r1) visibility.hiddenRows.add(row);
    }
    for (const [row, level] of layout.outlineRows) {
      if (row >= range.r0 && row <= range.r1) visibility.outlineRows.set(row, level);
    }
  } else {
    for (const col of layout.hiddenCols) {
      if (col >= range.c0 && col <= range.c1) visibility.hiddenCols.add(col);
    }
    for (const [col, level] of layout.outlineCols) {
      if (col >= range.c0 && col <= range.c1) visibility.outlineCols.set(col, level);
    }
  }
  return visibility;
}

export function axisVisibilityFromDimensions(
  dimensions: SourceDimensionSnapshot,
  storeVisibility: AxisVisibilitySnapshot,
): AxisVisibilitySnapshot {
  return {
    hiddenCols: new Set([...dimensions.hiddenCols, ...storeVisibility.hiddenCols]),
    hiddenRows: new Set([...dimensions.hiddenRows, ...storeVisibility.hiddenRows]),
    outlineCols: new Map([...dimensions.outlineCols, ...storeVisibility.outlineCols]),
    outlineRows: new Map([...dimensions.outlineRows, ...storeVisibility.outlineRows]),
  };
}

/**
 * Native row/column delete+insert journals restore cell values and formulas,
 * but an XF assignment can be reset by the inverse native operation after the
 * store's format history has already replayed. Put a repair entry at the
 * beginning of the cut transaction so its undo runs last, after both inverse
 * axis operations have reconstructed the source cells. The same flush also
 * restores source hyperlinks and validation ranges from the format slice.
 */
interface EngineXfSnapshot {
  entries: Map<string, { addr: Addr; xfIndex: number }>;
}

interface EngineXfCleanup {
  wholeRows: boolean;
  ranges: readonly Range[];
  extraIndex?: number;
}

export function captureEngineXfSnapshot(wb: WorkbookHandle, sheet: number): EngineXfSnapshot {
  const entries = new Map<string, { addr: Addr; xfIndex: number }>();
  for (const cell of wb.physicalCells(sheet)) {
    const xfIndex = wb.getCellXfIndex(sheet, cell.addr.row, cell.addr.col);
    if (xfIndex !== null && xfIndex > 0) {
      entries.set(addrKey(cell.addr), { addr: { ...cell.addr }, xfIndex });
    }
  }
  return { entries };
}

function clearEngineXfsNotInStore(
  wb: WorkbookHandle,
  store: SpreadsheetStore,
  sheet: number,
): void {
  const formatKeys = new Set<string>();
  for (const key of store.getState().format.formats.keys()) {
    if (parseAddrKey(key)?.sheet === sheet) formatKeys.add(key);
  }
  for (const cell of wb.physicalCells(sheet)) {
    const key = addrKey(cell.addr);
    if (formatKeys.has(key)) continue;
    const xfIndex = wb.getCellXfIndex(sheet, cell.addr.row, cell.addr.col);
    if (xfIndex !== null && xfIndex > 0) wb.setCellXfIndex(sheet, cell.addr.row, cell.addr.col, 0);
  }
}

function clearEngineXfsInCutAxis(
  wb: WorkbookHandle,
  sheet: number,
  cleanup: EngineXfCleanup,
  original: EngineXfSnapshot,
): void {
  const axisIndices = new Set<number>();
  if (cleanup.extraIndex !== undefined && cleanup.extraIndex >= 0) {
    axisIndices.add(cleanup.extraIndex);
  }
  const perpendicular = new Set<number>();
  for (const range of cleanup.ranges) {
    const first = cleanup.wholeRows ? range.r0 : range.c0;
    const last = cleanup.wholeRows ? range.r1 : range.c1;
    for (let index = first; index <= last; index += 1) axisIndices.add(index);
  }
  for (const cell of wb.physicalCells(sheet)) {
    const axis = cleanup.wholeRows ? cell.addr.row : cell.addr.col;
    const other = cleanup.wholeRows ? cell.addr.col : cell.addr.row;
    if (axisIndices.has(axis)) perpendicular.add(other);
  }
  for (const entry of original.entries.values()) {
    const axis = cleanup.wholeRows ? entry.addr.row : entry.addr.col;
    const other = cleanup.wholeRows ? entry.addr.col : entry.addr.row;
    if (axisIndices.has(axis)) perpendicular.add(other);
  }
  for (const axis of axisIndices) {
    for (const other of perpendicular) {
      const addr = cleanup.wholeRows
        ? { sheet, row: axis, col: other }
        : { sheet, row: other, col: axis };
      const xfIndex = wb.getCellXfIndex(sheet, addr.row, addr.col);
      if (xfIndex !== null && xfIndex > 0) wb.setCellXfIndex(sheet, addr.row, addr.col, 0);
    }
  }
}

function restoreEngineXfSnapshot(
  wb: WorkbookHandle,
  store: SpreadsheetStore,
  sheet: number,
  snapshot: EngineXfSnapshot,
  cleanup: EngineXfCleanup | null,
): void {
  flushFormatToEngine(wb, store, sheet);
  clearEngineXfsNotInStore(wb, store, sheet);
  if (cleanup) clearEngineXfsInCutAxis(wb, sheet, cleanup, snapshot);
  for (const { addr, xfIndex } of snapshot.entries.values()) {
    wb.setCellXfIndex(addr.sheet, addr.row, addr.col, xfIndex);
  }
}

export function recordSourceFormatEngineRepair(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  history: History | null,
  sourceSheet: number,
  snapshot: EngineXfSnapshot,
  cleanup: EngineXfCleanup | null,
): void {
  if (!history) return;
  history.push({
    undo: () => restoreEngineXfSnapshot(wb, store, sourceSheet, snapshot, cleanup),
    redo: () => {
      flushFormatToEngine(wb, store, sourceSheet);
      if (cleanup) clearEngineXfsInCutAxis(wb, sourceSheet, cleanup, snapshot);
    },
  });
}

export function captureSourceDimensionOverrides(
  wb: WorkbookHandle,
  source: Range,
  wholeRows: boolean,
): SourceDimensionSnapshot {
  const colWidths = new Map<number, number>();
  const rowHeights = new Map<number, number>();
  const hiddenCols = new Set<number>();
  const hiddenRows = new Set<number>();
  const outlineCols = new Map<number, number>();
  const outlineRows = new Map<number, number>();
  if (wholeRows) {
    for (const layout of wb.getRowLayouts(source.sheet)) {
      if (layout.row < source.r0 || layout.row > source.r1) continue;
      if (layout.height > 0) rowHeights.set(layout.row, layout.height);
      if (layout.hidden) hiddenRows.add(layout.row);
      if (layout.outlineLevel > 0) outlineRows.set(layout.row, layout.outlineLevel);
    }
  } else {
    for (const layout of wb.getColumnLayouts(source.sheet)) {
      const first = Math.max(source.c0, layout.first);
      const last = Math.min(source.c1, layout.last);
      for (let col = first; col <= last; col += 1) {
        if (layout.width > 0) colWidths.set(col, layout.width);
        if (layout.hidden) hiddenCols.add(col);
        if (layout.outlineLevel > 0) outlineCols.set(col, layout.outlineLevel);
      }
    }
  }
  return { colWidths, rowHeights, hiddenCols, hiddenRows, outlineCols, outlineRows };
}

function applySourceDimensionOverrides(
  wb: WorkbookHandle,
  source: Range,
  wholeRows: boolean,
  desired: SourceDimensionSnapshot,
  defaults: { col: number; row: number },
): void {
  if (wholeRows) {
    const rows = new Set<number>();
    for (const row of desired.rowHeights.keys()) rows.add(row);
    for (const row of desired.hiddenRows) rows.add(row);
    for (const row of desired.outlineRows.keys()) rows.add(row);
    for (const layout of wb.getRowLayouts(source.sheet)) {
      for (
        let row = Math.max(source.r0, layout.row);
        row <= Math.min(source.r1, layout.row);
        row += 1
      ) {
        rows.add(row);
      }
    }
    for (const row of rows) {
      wb.setRowHeight(source.sheet, row, desired.rowHeights.get(row) ?? defaults.row);
      wb.setRowHidden(source.sheet, row, desired.hiddenRows.has(row));
      wb.setRowOutline(source.sheet, row, desired.outlineRows.get(row) ?? 0);
    }
    return;
  }
  const cols = new Set<number>();
  for (const col of desired.colWidths.keys()) cols.add(col);
  for (const col of desired.hiddenCols) cols.add(col);
  for (const col of desired.outlineCols.keys()) cols.add(col);
  for (const layout of wb.getColumnLayouts(source.sheet)) {
    const first = Math.max(source.c0, layout.first);
    const last = Math.min(source.c1, layout.last);
    for (let col = first; col <= last; col += 1) cols.add(col);
  }
  for (const col of cols) {
    wb.setColumnWidth(source.sheet, col, col, desired.colWidths.get(col) ?? defaults.col);
    wb.setColumnHidden(source.sheet, col, col, desired.hiddenCols.has(col));
    wb.setColumnOutline(source.sheet, col, col, desired.outlineCols.get(col) ?? 0);
  }
}

function captureEngineAxisVisibility(
  wb: WorkbookHandle,
  range: Range,
  wholeRows: boolean,
): AxisVisibilitySnapshot {
  const dimensions = captureSourceDimensionOverrides(wb, range, wholeRows);
  return {
    hiddenCols: dimensions.hiddenCols,
    hiddenRows: dimensions.hiddenRows,
    outlineCols: dimensions.outlineCols,
    outlineRows: dimensions.outlineRows,
  };
}

function applyEngineAxisVisibility(
  wb: WorkbookHandle,
  range: Range,
  wholeRows: boolean,
  desired: AxisVisibilitySnapshot,
): void {
  if (wholeRows) {
    const rows = new Set<number>();
    for (const row of desired.hiddenRows) rows.add(row);
    for (const row of desired.outlineRows.keys()) rows.add(row);
    for (const layout of wb.getRowLayouts(range.sheet)) {
      const first = Math.max(range.r0, layout.row);
      const last = Math.min(range.r1, layout.row);
      for (let row = first; row <= last; row += 1) rows.add(row);
    }
    for (const row of rows) {
      wb.setRowHidden(range.sheet, row, desired.hiddenRows.has(row));
      wb.setRowOutline(range.sheet, row, desired.outlineRows.get(row) ?? 0);
    }
    return;
  }
  const cols = new Set<number>();
  for (const col of desired.hiddenCols) cols.add(col);
  for (const col of desired.outlineCols.keys()) cols.add(col);
  for (const layout of wb.getColumnLayouts(range.sheet)) {
    const first = Math.max(range.c0, layout.first);
    const last = Math.min(range.c1, layout.last);
    for (let col = first; col <= last; col += 1) cols.add(col);
  }
  for (const col of cols) {
    wb.setColumnHidden(range.sheet, col, col, desired.hiddenCols.has(col));
    wb.setColumnOutline(range.sheet, col, col, desired.outlineCols.get(col) ?? 0);
  }
}

function applyStoreAxisVisibility(
  store: SpreadsheetStore,
  range: Range,
  wholeRows: boolean,
  desired: AxisVisibilitySnapshot,
): void {
  store.setState((s) => {
    const hiddenRows = new Set(s.layout.hiddenRows);
    const hiddenCols = new Set(s.layout.hiddenCols);
    const outlineRows = new Map(s.layout.outlineRows);
    const outlineCols = new Map(s.layout.outlineCols);
    if (wholeRows) {
      for (const row of [...hiddenRows]) {
        if (row >= range.r0 && row <= range.r1) hiddenRows.delete(row);
      }
      for (const row of [...outlineRows.keys()]) {
        if (row >= range.r0 && row <= range.r1) outlineRows.delete(row);
      }
      for (const row of desired.hiddenRows) hiddenRows.add(row);
      for (const [row, level] of desired.outlineRows) outlineRows.set(row, level);
    } else {
      for (const col of [...hiddenCols]) {
        if (col >= range.c0 && col <= range.c1) hiddenCols.delete(col);
      }
      for (const col of [...outlineCols.keys()]) {
        if (col >= range.c0 && col <= range.c1) outlineCols.delete(col);
      }
      for (const col of desired.hiddenCols) hiddenCols.add(col);
      for (const [col, level] of desired.outlineCols) outlineCols.set(col, level);
    }
    return {
      ...s,
      layout: { ...s.layout, hiddenRows, hiddenCols, outlineRows, outlineCols },
    };
  });
}

function projectAxisVisibility(
  source: AxisVisibilitySnapshot,
  sourceRange: Range,
  targetRange: Range,
  wholeRows: boolean,
): AxisVisibilitySnapshot {
  const projected = emptyAxisVisibility();
  const sourceStart = wholeRows ? sourceRange.r0 : sourceRange.c0;
  const targetStart = wholeRows ? targetRange.r0 : targetRange.c0;
  const count = wholeRows
    ? targetRange.r1 - targetRange.r0 + 1
    : targetRange.c1 - targetRange.c0 + 1;
  for (let offset = 0; offset < count; offset += 1) {
    const sourceIndex = sourceStart + offset;
    const targetIndex = targetStart + offset;
    if (wholeRows) {
      if (source.hiddenRows.has(sourceIndex)) projected.hiddenRows.add(targetIndex);
      const level = source.outlineRows.get(sourceIndex);
      if (level !== undefined) projected.outlineRows.set(targetIndex, level);
    } else {
      if (source.hiddenCols.has(sourceIndex)) projected.hiddenCols.add(targetIndex);
      const level = source.outlineCols.get(sourceIndex);
      if (level !== undefined) projected.outlineCols.set(targetIndex, level);
    }
  }
  return projected;
}

function sameAxisVisibility(a: AxisVisibilitySnapshot, b: AxisVisibilitySnapshot): boolean {
  return (
    a.hiddenCols.size === b.hiddenCols.size &&
    [...a.hiddenCols].every((col) => b.hiddenCols.has(col)) &&
    a.hiddenRows.size === b.hiddenRows.size &&
    [...a.hiddenRows].every((row) => b.hiddenRows.has(row)) &&
    a.outlineCols.size === b.outlineCols.size &&
    [...a.outlineCols].every(([col, level]) => b.outlineCols.get(col) === level) &&
    a.outlineRows.size === b.outlineRows.size &&
    [...a.outlineRows].every(([row, level]) => b.outlineRows.get(row) === level)
  );
}

export function recordAxisVisibilityTransfer(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  history: History | null,
  sourceRange: Range,
  targetRange: Range,
  wholeRows: boolean,
  source: AxisVisibilitySnapshot,
): void {
  const beforeEngine = captureEngineAxisVisibility(wb, targetRange, wholeRows);
  const beforeStore = captureStoreAxisVisibility(store, targetRange, wholeRows);
  const afterEngine = projectAxisVisibility(source, sourceRange, targetRange, wholeRows);
  const afterStore = afterEngine;
  if (
    sameAxisVisibility(beforeEngine, afterEngine) &&
    sameAxisVisibility(beforeStore, afterStore)
  ) {
    return;
  }
  const apply = (engine: AxisVisibilitySnapshot, state: AxisVisibilitySnapshot): void => {
    applyStoreAxisVisibility(store, targetRange, wholeRows, state);
    applyEngineAxisVisibility(wb, targetRange, wholeRows, engine);
  };
  apply(afterEngine, afterStore);
  if (!history || history.isReplaying()) return;
  history.push({
    undo: () => apply(beforeEngine, beforeStore),
    redo: () => apply(afterEngine, afterStore),
  });
}

export function resetCrossSheetSourceDimensions(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  history: History | null,
  snapshot: ClipboardSnapshot,
  wholeRows: boolean,
): void {
  const source = snapshot.logicalRange ?? snapshot.range;
  const before = captureSourceDimensionOverrides(wb, source, wholeRows);
  if (wholeRows) {
    for (const [offset, height] of snapshot.rowHeights ?? []) {
      if (Number.isInteger(offset) && offset >= 0 && Number.isFinite(height)) {
        before.rowHeights.set(source.r0 + offset, height);
      }
    }
  } else {
    for (const [offset, width] of snapshot.colWidths ?? []) {
      if (Number.isInteger(offset) && offset >= 0 && Number.isFinite(width)) {
        before.colWidths.set(source.c0 + offset, width);
      }
    }
  }
  if (
    before.colWidths.size === 0 &&
    before.rowHeights.size === 0 &&
    before.hiddenCols.size === 0 &&
    before.hiddenRows.size === 0 &&
    before.outlineCols.size === 0 &&
    before.outlineRows.size === 0
  ) {
    return;
  }
  const defaults = {
    col: store.getState().layout.defaultColWidth,
    row: store.getState().layout.defaultRowHeight,
  };
  const after: SourceDimensionSnapshot = {
    colWidths: new Map(),
    rowHeights: new Map(),
    hiddenCols: new Set(),
    hiddenRows: new Set(),
    outlineCols: new Map(),
    outlineRows: new Map(),
  };
  const apply = (desired: SourceDimensionSnapshot): void => {
    applySourceDimensionOverrides(wb, source, wholeRows, desired, defaults);
  };
  apply(after);
  if (!history || history.isReplaying()) return;
  history.push({
    undo: () => apply(before),
    redo: () => apply(after),
  });
}

interface SameSheetDimensionRestore {
  engine: SourceDimensionSnapshot;
  storeColWidths: Map<number, number>;
  storeRowHeights: Map<number, number>;
  storeHiddenCols: Set<number>;
  storeHiddenRows: Set<number>;
  storeOutlineCols: Map<number, number>;
  storeOutlineRows: Map<number, number>;
}

export function captureSameSheetDimensionRestore(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  source: Range,
  wholeRows: boolean,
): SameSheetDimensionRestore {
  const engine = captureSourceDimensionOverrides(wb, source, wholeRows);
  const state = store.getState();
  const storeColWidths = new Map<number, number>();
  const storeRowHeights = new Map<number, number>();
  const storeHiddenCols = new Set<number>();
  const storeHiddenRows = new Set<number>();
  const storeOutlineCols = new Map<number, number>();
  const storeOutlineRows = new Map<number, number>();
  if (wholeRows) {
    for (const [row, height] of state.layout.rowHeights) {
      if (row >= source.r0 && row <= source.r1) storeRowHeights.set(row, height);
    }
    for (const row of state.layout.hiddenRows) {
      if (row >= source.r0 && row <= source.r1) storeHiddenRows.add(row);
    }
    for (const [row, level] of state.layout.outlineRows) {
      if (row >= source.r0 && row <= source.r1) storeOutlineRows.set(row, level);
    }
  } else {
    for (const [col, width] of state.layout.colWidths) {
      if (col >= source.c0 && col <= source.c1) storeColWidths.set(col, width);
    }
    for (const col of state.layout.hiddenCols) {
      if (col >= source.c0 && col <= source.c1) storeHiddenCols.add(col);
    }
    for (const [col, level] of state.layout.outlineCols) {
      if (col >= source.c0 && col <= source.c1) storeOutlineCols.set(col, level);
    }
  }
  return {
    engine,
    storeColWidths,
    storeRowHeights,
    storeHiddenCols,
    storeHiddenRows,
    storeOutlineCols,
    storeOutlineRows,
  };
}

function applySameSheetDimensionRestore(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  source: Range,
  wholeRows: boolean,
  snapshot: SameSheetDimensionRestore,
  defaults: { col: number; row: number },
): void {
  applySourceDimensionOverrides(wb, source, wholeRows, snapshot.engine, defaults);
  store.setState((s) => {
    const colWidths = new Map(s.layout.colWidths);
    const rowHeights = new Map(s.layout.rowHeights);
    if (wholeRows) {
      for (const row of [...rowHeights.keys()]) {
        if (row >= source.r0 && row <= source.r1) rowHeights.delete(row);
      }
      for (const [row, height] of snapshot.storeRowHeights) rowHeights.set(row, height);
      const hiddenRows = new Set(s.layout.hiddenRows);
      for (const row of [...hiddenRows]) {
        if (row >= source.r0 && row <= source.r1) hiddenRows.delete(row);
      }
      for (const row of snapshot.storeHiddenRows) hiddenRows.add(row);
      const outlineRows = new Map(s.layout.outlineRows);
      for (const row of [...outlineRows.keys()]) {
        if (row >= source.r0 && row <= source.r1) outlineRows.delete(row);
      }
      for (const [row, level] of snapshot.storeOutlineRows) outlineRows.set(row, level);
      return {
        ...s,
        layout: { ...s.layout, colWidths, rowHeights, hiddenRows, outlineRows },
      };
    } else {
      for (const col of [...colWidths.keys()]) {
        if (col >= source.c0 && col <= source.c1) colWidths.delete(col);
      }
      for (const [col, width] of snapshot.storeColWidths) colWidths.set(col, width);
      const hiddenCols = new Set(s.layout.hiddenCols);
      for (const col of [...hiddenCols]) {
        if (col >= source.c0 && col <= source.c1) hiddenCols.delete(col);
      }
      for (const col of snapshot.storeHiddenCols) hiddenCols.add(col);
      const outlineCols = new Map(s.layout.outlineCols);
      for (const col of [...outlineCols.keys()]) {
        if (col >= source.c0 && col <= source.c1) outlineCols.delete(col);
      }
      for (const [col, level] of snapshot.storeOutlineCols) outlineCols.set(col, level);
      return {
        ...s,
        layout: { ...s.layout, colWidths, rowHeights, hiddenCols, outlineCols },
      };
    }
  });
}

export function recordSameSheetDimensionRestore(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  history: History | null,
  source: Range,
  wholeRows: boolean,
  snapshot: SameSheetDimensionRestore,
): void {
  if (!history) return;
  if (
    snapshot.engine.colWidths.size === 0 &&
    snapshot.engine.rowHeights.size === 0 &&
    snapshot.storeColWidths.size === 0 &&
    snapshot.storeRowHeights.size === 0 &&
    snapshot.storeHiddenCols.size === 0 &&
    snapshot.storeHiddenRows.size === 0 &&
    snapshot.storeOutlineCols.size === 0 &&
    snapshot.storeOutlineRows.size === 0
  ) {
    return;
  }
  const defaults = {
    col: store.getState().layout.defaultColWidth,
    row: store.getState().layout.defaultRowHeight,
  };
  history.push({
    undo: () => applySameSheetDimensionRestore(store, wb, source, wholeRows, snapshot, defaults),
    redo: () => {},
  });
}

export function consumeCutMarquee(store: SpreadsheetStore): void {
  mutators.setCopyRange(store, null);
  mutators.setCopyRanges(store, null);
}

interface CutTransientSnapshot {
  selection: State['selection'];
  copyRange: Range | null;
  copyRanges: Range[] | null | undefined;
  copyMode: State['ui']['copyMode'];
  copyRevision: number | undefined;
}

export function captureCutTransientSnapshot(state: State): CutTransientSnapshot {
  return {
    selection: {
      ...state.selection,
      active: { ...state.selection.active },
      anchor: { ...state.selection.anchor },
      range: { ...state.selection.range },
      extraRanges: state.selection.extraRanges?.map((range) => ({ ...range })),
    },
    copyRange: state.ui.copyRange ? { ...state.ui.copyRange } : null,
    copyRanges:
      state.ui.copyRanges == null
        ? state.ui.copyRanges
        : state.ui.copyRanges.map((range) => ({ ...range })),
    copyMode: state.ui.copyMode,
    copyRevision: state.ui.copyRevision,
  };
}

export function restoreCutTransientSnapshot(
  store: SpreadsheetStore,
  snapshot: CutTransientSnapshot,
): void {
  store.setState((state) => ({
    ...state,
    selection: {
      ...snapshot.selection,
      active: { ...snapshot.selection.active },
      anchor: { ...snapshot.selection.anchor },
      range: { ...snapshot.selection.range },
      extraRanges: snapshot.selection.extraRanges?.map((range) => ({ ...range })),
    },
    ui: {
      ...state.ui,
      copyRange: snapshot.copyRange ? { ...snapshot.copyRange } : null,
      copyRanges:
        snapshot.copyRanges == null
          ? snapshot.copyRanges
          : snapshot.copyRanges.map((range) => ({ ...range })),
      copyMode: snapshot.copyMode,
      copyRevision: snapshot.copyRevision,
    },
  }));
}
