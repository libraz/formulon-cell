import type { WorkbookHandle } from '../engine/workbook-handle.js';

/**
 * The data connections of `wb` that Refresh All runs. The engine exposes none to refresh yet, so
 * a workbook has none: Refresh All then changes nothing, records no undo step and emits no event,
 * as on a desktop workbook without connections.
 */
const dataConnections = (_wb: WorkbookHandle): readonly (() => void)[] => [];

/** Refresh All: refresh every data connection of `wb`. Returns how many it refreshed. */
export function refreshAllConnections(wb: WorkbookHandle): number {
  const connections = dataConnections(wb);
  for (const refresh of connections) refresh();
  return connections.length;
}
