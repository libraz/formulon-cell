import {
  mutators,
  type SpreadsheetStore,
  type State,
  type StatusAggKey,
  type WorkbookViewMode,
} from '../store/store.js';
import type { History } from './history.js';

/** Toggle worksheet gridlines. */
export function setGridlinesVisible(store: SpreadsheetStore, visible: boolean): void {
  mutators.setShowGridLines(store, visible);
}

/** Toggle row/column headings. */
export function setHeadingsVisible(store: SpreadsheetStore, visible: boolean): void {
  mutators.setShowHeaders(store, visible);
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

/** Select the workbook view mode shown by View > Workbook Views. */
export function setWorkbookView(store: SpreadsheetStore, mode: WorkbookViewMode): void {
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
