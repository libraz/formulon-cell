import type { TableOverlay } from '../commands/format-as-table.js';

/** Case-insensitive key of a table, slicer or style name. */
export const nameKey = (name: string): string => name.toLocaleLowerCase();

/** The named engine overlays of `tables` by `nameKey`. */
export const engineOverlaysByName = (
  tables: readonly TableOverlay[],
): Map<string, TableOverlay> => {
  const known = new Map<string, TableOverlay>();
  for (const t of tables) {
    if (t.source === 'engine' && t.name) known.set(nameKey(t.name), t);
  }
  return known;
};
