import { mutators, type SpreadsheetStore, type StatusAggKey } from '../store/store.js';

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

/** Toggle R1C1 reference style for visible headers/name-box references. */
export function setR1C1ReferenceStyle(store: SpreadsheetStore, enabled: boolean): void {
  mutators.setR1C1(store, enabled);
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
