import type { SpreadsheetStore } from '../store/store.js';

/** Per mounted store, whether its chrome is the Mac desktop's, read live from the host as the grid
 *  reads it; the workbook's engine profile is a different switch. */
const macChrome = new WeakMap<SpreadsheetStore, () => boolean>();

/** Register how the mounted instance tells whether its chrome is the Mac desktop's. */
export function registerChromePlatform(store: SpreadsheetStore, isMac: () => boolean): () => void {
  macChrome.set(store, isMac);
  return () => {
    if (macChrome.get(store) === isMac) macChrome.delete(store);
  };
}

/** Whether `store`'s chrome is the Mac desktop's; an unmounted store is not. */
export const isMacChrome = (store: SpreadsheetStore): boolean => macChrome.get(store)?.() ?? false;
