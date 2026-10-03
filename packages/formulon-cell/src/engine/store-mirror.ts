import type { CellFormat, PageSetup, SpreadsheetStore, State } from '../store/store.js';
import { parseAddrKey } from './address.js';
import { syncAutoFilterToEngine } from './auto-filter-sync.js';
import { flushFormatToEngine } from './cell-format-sync.js';
import { syncPageSetupToEngine } from './print-sync.js';
import type { WorkbookHandle } from './workbook-handle.js';

const autoFilterSignature = (state: State): string =>
  JSON.stringify({
    range: state.ui.filterRange,
    criteria: state.ui.filterCriteria.map((entry) => ({
      range: entry.range,
      byCol: entry.byCol,
      hiddenValues: entry.hiddenValues,
      condition: entry.condition,
      color: entry.color,
    })),
  });

const formatSheetFromKey = (key: string): number | null => {
  const addr = parseAddrKey(key);
  return addr && addr.sheet >= 0 ? addr.sheet : null;
};

const changedFormatSheets = (
  previous: ReadonlyMap<string, CellFormat> | null,
  next: ReadonlyMap<string, CellFormat>,
): Set<number> => {
  const changed = new Set<number>();
  if (previous === next) return changed;
  if (previous === null) {
    for (const key of next.keys()) {
      const sheet = formatSheetFromKey(key);
      if (sheet !== null) changed.add(sheet);
    }
    return changed;
  }
  for (const [key, format] of next) {
    if (previous.get(key) === format) continue;
    const sheet = formatSheetFromKey(key);
    if (sheet !== null) changed.add(sheet);
  }
  for (const key of previous.keys()) {
    if (next.has(key)) continue;
    const sheet = formatSheetFromKey(key);
    if (sheet !== null) changed.add(sheet);
  }
  return changed;
};

const sheetsWithCellStyles = (formats: ReadonlyMap<string, CellFormat>): Set<number> => {
  const sheets = new Set<number>();
  for (const [key, format] of formats) {
    if (!format.cellStyle) continue;
    const sheet = formatSheetFromKey(key);
    if (sheet !== null) sheets.add(sheet);
  }
  if (sheets.size === 0) sheets.add(0);
  return sheets;
};

/** Mirrors a bound mount store's formats, AutoFilter and page-setup slices
 *  into the engine as they change. */
export class StoreMirror {
  constructor(private readonly handle: WorkbookHandle) {}

  /** Optional mount store whose AutoFilter and page-setup slices are mirrored
   * into the engine as they change. Keeping this association at the adapter
   * boundary means every React/Vue/vanilla surface gets save-safe filters and
   * print settings without duplicating writeback calls. */
  private boundStore: SpreadsheetStore | null = null;

  private unsubscribeBoundStore: (() => void) | null = null;

  private lastAutoFilterSignature = '';

  /** Identity of the page-setup map the engine already carries. The store
   *  replaces the map on every page-setup write and preserves it otherwise, so
   *  comparing the reference costs nothing on the unrelated notifications that
   *  make up almost all of them. A same-content map with a fresh identity only
   *  costs one redundant flush, and the flush is idempotent. */
  private lastPageSetupMap: ReadonlyMap<number, PageSetup> | null = null;

  /** Format entries are compared by key and value identity. Format mutators
   *  replace a changed entry object, so this catches edits and removals without
   *  deep-comparing every CellFormat on every store notification. */
  private lastFormatEntries: ReadonlyMap<string, CellFormat> | null = null;

  /** Named styles are session metadata; changing the array can affect every
   *  cell that references one, so its identity is tracked independently. */
  private lastCustomCellStyles: State['format']['customCellStyles'] | null = null;

  private storeSyncMuted = 0;

  attach(store: SpreadsheetStore | null): void {
    this.unsubscribeBoundStore?.();
    this.unsubscribeBoundStore = null;
    this.boundStore = store;
    this.rebase();
    if (!store) return;
    this.unsubscribeBoundStore = store.subscribe((state) => {
      const nextFormats = state.format.formats;
      const formatSheets = changedFormatSheets(this.lastFormatEntries, nextFormats);
      const stylesChanged = (state.format.customCellStyles ?? null) !== this.lastCustomCellStyles;
      if (stylesChanged) {
        for (const sheet of sheetsWithCellStyles(nextFormats)) formatSheets.add(sheet);
      }
      // Advance all baselines even while hydration is muted. The state that
      // hydration imports is the new engine baseline, not a pending UI edit.
      this.lastFormatEntries = nextFormats;
      this.lastCustomCellStyles = state.format.customCellStyles ?? null;
      const nextFilter = autoFilterSignature(state);
      const filterChanged = nextFilter !== this.lastAutoFilterSignature;
      if (filterChanged) this.lastAutoFilterSignature = nextFilter;
      const nextPageSetup = state.pageSetup.setupBySheet;
      const pageSetupChanged = nextPageSetup !== this.lastPageSetupMap;
      if (pageSetupChanged) this.lastPageSetupMap = nextPageSetup;
      if (this.storeSyncMuted > 0) return;
      for (const sheet of formatSheets) {
        if (sheet < this.handle.sheetCount) flushFormatToEngine(this.handle, store, sheet);
      }
      if (filterChanged) syncAutoFilterToEngine(this.handle, state, state.data.sheetIndex);
      if (pageSetupChanged) syncPageSetupToEngine(this.handle, store, state.data.sheetIndex);
    });
  }

  detach(expectedStore: SpreadsheetStore): void {
    if (this.boundStore !== expectedStore) return;
    this.attach(null);
  }

  withMuted<T>(fn: () => T): T {
    this.storeSyncMuted += 1;
    try {
      return fn();
    } finally {
      this.storeSyncMuted -= 1;
      this.rebase();
    }
  }

  /** Re-read the bound store's mirrored slices so the next notification is
   *  compared against what the engine already carries. */
  private rebase(): void {
    const state = this.boundStore?.getState();
    this.lastFormatEntries = state?.format.formats ?? null;
    this.lastCustomCellStyles = state?.format.customCellStyles ?? null;
    this.lastAutoFilterSignature = state ? autoFilterSignature(state) : '';
    this.lastPageSetupMap = state?.pageSetup.setupBySheet ?? null;
  }
}
