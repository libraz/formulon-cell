import type { WorkbookHandle } from '../engine/workbook-handle.js';
import {
  mutators,
  type SpreadsheetStore,
  type State,
  type StatusAggKey,
  type WorkbookViewMode,
} from '../store/store.js';
import type { History } from './history.js';

/** Toggle worksheet gridlines. */
export function setGridlinesVisible(
  store: SpreadsheetStore,
  visible: boolean,
  wb: WorkbookHandle | null = null,
): void {
  mutators.setShowGridLines(store, visible);
  if (wb && typeof wb.setSheetShowGridLines === 'function') {
    wb.setSheetShowGridLines(store.getState().data.sheetIndex, visible);
  }
}

/** Toggle row/column headings. */
export function setHeadingsVisible(
  store: SpreadsheetStore,
  visible: boolean,
  wb: WorkbookHandle | null = null,
): void {
  mutators.setShowHeaders(store, visible);
  if (wb && typeof wb.setSheetShowRowColHeaders === 'function') {
    wb.setSheetShowRowColHeaders(store.getState().data.sheetIndex, visible);
  }
}

/** Toggle whether numeric zero values are displayed. */
export function setZerosVisible(
  store: SpreadsheetStore,
  visible: boolean,
  wb: WorkbookHandle | null = null,
): void {
  mutators.setShowZeros(store, visible);
  if (wb && typeof wb.setSheetShowZeros === 'function') {
    wb.setSheetShowZeros(store.getState().data.sheetIndex, visible);
  }
}

/** Toggle the sheet's right-to-left direction, mirroring the grid axis. */
export function setSheetRightToLeft(
  store: SpreadsheetStore,
  rightToLeft: boolean,
  wb: WorkbookHandle | null = null,
): void {
  mutators.setRightToLeft(store, rightToLeft);
  if (wb && typeof wb.setSheetRightToLeft === 'function') {
    wb.setSheetRightToLeft(store.getState().data.sheetIndex, rightToLeft);
  }
}

/** Toggle formula text display. */
export function setShowFormulas(store: SpreadsheetStore, visible: boolean): void {
  mutators.setShowFormulas(store, visible);
}

/** Set or clear the Excel-style sheet background image for on-screen display. */
export function setSheetBackgroundImage(
  store: SpreadsheetStore,
  sheet: number,
  url: string | undefined,
  history: History | null = null,
): void {
  recordSheetBackgroundChange(history, store, () => {
    mutators.setSheetBackgroundImage(store, sheet, url);
  });
}

export function clearSheetBackgroundImage(
  store: SpreadsheetStore,
  sheet: number,
  history: History | null = null,
): void {
  recordSheetBackgroundChange(history, store, () => {
    mutators.setSheetBackgroundImage(store, sheet, undefined);
  });
}

/** Toggle R1C1 reference style for visible headers/name-box references. */
export function setR1C1ReferenceStyle(store: SpreadsheetStore, enabled: boolean): void {
  mutators.setR1C1(store, enabled);
}

/** Zoom Page Break Preview pulls back to, so a whole page fits on screen. */
export const PAGE_BREAK_PREVIEW_ZOOM = 0.6;

/**
 * Select the workbook view mode shown by View > Workbook Views.
 *
 * Opening Page Break Preview zooms out far enough to read a whole page and
 * remembers the zoom it replaced; leaving the preview puts that zoom back.
 * Page Layout keeps whatever zoom is in force — its own page framing already
 * tells the user where the paper ends.
 */
export function setWorkbookView(store: SpreadsheetStore, mode: WorkbookViewMode): void {
  const previous = store.getState().ui.workbookView;
  if (previous === mode) return;
  if (mode === 'pageBreakPreview') {
    const zoom = store.getState().viewport.zoom;
    store.setState((s) => ({ ...s, ui: { ...s.ui, zoomBeforePreview: zoom } }));
    mutators.setZoom(store, PAGE_BREAK_PREVIEW_ZOOM);
  } else if (previous === 'pageBreakPreview') {
    const restore = store.getState().ui.zoomBeforePreview;
    if (restore && restore > 0) mutators.setZoom(store, restore);
    store.setState((s) => ({ ...s, ui: { ...s.ui, zoomBeforePreview: null } }));
  }
  mutators.setWorkbookView(store, mode);
}

/** Set zoom as a decimal scale. Values are clamped by the store to 50..400%. */
export function setZoomScale(store: SpreadsheetStore, zoom: number): void {
  mutators.setZoom(store, zoom);
}

/** Set zoom as a percentage. Values are clamped to 50..400%. */
export function setZoomPercent(store: SpreadsheetStore, percent: number): void {
  mutators.setZoom(store, percent / 100);
}

/** Set the per-sheet zoom level. `zoom` is a multiplier (1.0 = 100%) and is
 *  clamped by the store mutator to [0.5, 4]. When `wb` is supplied the
 *  engine receives the equivalent percentage so the value round-trips
 *  through .xlsx. Not journaled — spreadsheets treat zoom as a view setting
 *  outside the undo stack. */
export function setSheetZoom(store: SpreadsheetStore, zoom: number, wb?: WorkbookHandle): void {
  mutators.setZoom(store, zoom);
  if (wb) {
    const sheet = store.getState().data.sheetIndex;
    const pct = Math.round(store.getState().viewport.zoom * 100);
    wb.setSheetZoom(sheet, pct);
  }
}

/** Configure which status-bar aggregates are visible. */
export function setStatusAggregates(store: SpreadsheetStore, keys: StatusAggKey[]): void {
  mutators.setStatusAggs(store, keys);
}

/** Toggle a single status-bar aggregate. */
export function toggleStatusAggregate(store: SpreadsheetStore, key: StatusAggKey): void {
  mutators.toggleStatusAgg(store, key);
}

function captureSheetBackgroundSnapshot(state: State): Map<number, string> {
  return new Map(state.ui.sheetBackgroundImages);
}

function applySheetBackgroundSnapshot(
  store: SpreadsheetStore,
  snap: ReadonlyMap<number, string>,
): void {
  store.setState((s) => ({
    ...s,
    ui: { ...s.ui, sheetBackgroundImages: new Map(snap) },
  }));
}

const sameSheetBackgroundSnapshot = (
  a: ReadonlyMap<number, string>,
  b: ReadonlyMap<number, string>,
): boolean => {
  if (a.size !== b.size) return false;
  for (const [sheet, url] of a) {
    if (b.get(sheet) !== url) return false;
  }
  return true;
};

function recordSheetBackgroundChange(
  history: History | null,
  store: SpreadsheetStore,
  mutate: () => void,
): void {
  if (!history || history.isReplaying()) {
    mutate();
    return;
  }
  const before = captureSheetBackgroundSnapshot(store.getState());
  mutate();
  const after = captureSheetBackgroundSnapshot(store.getState());
  if (sameSheetBackgroundSnapshot(before, after)) return;
  history.push({
    undo: () => applySheetBackgroundSnapshot(store, before),
    redo: () => applySheetBackgroundSnapshot(store, after),
  });
}
