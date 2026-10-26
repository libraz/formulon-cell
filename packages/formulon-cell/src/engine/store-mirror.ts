import { associatePaginationEngine } from '../commands/pagination.js';
import { hiddenRowsOf } from '../store/row-visibility.js';
import {
  type CellFormat,
  mutators,
  type PageSetup,
  type SpreadsheetStore,
  type State,
} from '../store/store.js';
import type { OutlineSummary } from '../store/types.js';
import { parseAddrKey } from './address.js';
import { syncAutoFilterToEngine, syncHiddenRowsToEngine } from './auto-filter-sync.js';
import { remapActiveAxisFormatOverrides, remapAxisFormatOverrides } from './axis-format-sync.js';
import { bindCalcSettingsStore } from './calc-settings-sync.js';
import {
  captureSyncedFormatKeyMembership,
  flushFormatToEngine,
  remapSyncedFormatKeys,
  restoreSyncedFormatKeyMembership,
  syncNamedStyleCatalogToEngine,
} from './cell-format-sync.js';
import { DrawingObjectMirror } from './drawing-object-mirror.js';
import {
  createFormatHistoryBridge,
  type FormatHistoryBridge,
  registerFormatHistoryBridge,
  unregisterFormatHistoryBridge,
} from './format-history-bridge.js';
import { syncPageSetupToEngine } from './print-sync.js';
import { syncSheetPrToEngine } from './sheet-pr-sync.js';
import { TableMirror } from './table-mirror.js';
import type { SheetAxisFormatOverride, SheetFontOverride } from './types.js';
import { syncViewStateToEngine } from './view-state-sync.js';
import { bindWholeSheetFontStore } from './whole-sheet-font-sync.js';
import type { WorkbookHandle } from './workbook-handle.js';

const autoFilterSignature = (state: State): string =>
  JSON.stringify({
    ranges: state.ui.filterRanges,
    sorts: state.ui.filterSorts,
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
  return sheets;
};

const changedMapKeys = <V>(
  previous: ReadonlyMap<number, V> | undefined,
  next: ReadonlyMap<number, V> | undefined,
  same: (a: V, b: V) => boolean = Object.is,
): Set<number> => {
  const changed = new Set<number>();
  if (previous === next) return changed;
  for (const [key, value] of next ?? []) {
    const before = previous?.get(key);
    if (before === undefined || !same(before, value)) changed.add(key);
  }
  for (const key of previous?.keys() ?? []) if (!next?.has(key)) changed.add(key);
  return changed;
};

const sameOutlineSummary = (a: OutlineSummary, b: OutlineSummary): boolean =>
  a.below === b.below && a.right === b.right;

const sameSheetFontOverride = (
  before: SheetFontOverride | undefined,
  after: SheetFontOverride | undefined,
): boolean => JSON.stringify(before) === JSON.stringify(after);

const changedSheetFontOverrideSheets = (
  previous: ReadonlyMap<number, SheetFontOverride> | null,
  next: ReadonlyMap<number, SheetFontOverride>,
): Set<number> => {
  const changed = new Set<number>();
  if (previous === next) return changed;
  if (previous === null) {
    for (const sheet of next.keys()) changed.add(sheet);
    return changed;
  }
  for (const [sheet, override] of next) {
    if (!sameSheetFontOverride(previous.get(sheet), override)) changed.add(sheet);
  }
  for (const sheet of previous.keys()) {
    if (!next.has(sheet)) changed.add(sheet);
  }
  return changed;
};

const sameAxisFormatOverride = (
  before: SheetAxisFormatOverride | undefined,
  after: SheetAxisFormatOverride | undefined,
): boolean => JSON.stringify(before) === JSON.stringify(after);

const changedAxisFormatOverrideSheets = (
  previous: ReadonlyMap<number, SheetAxisFormatOverride> | null,
  next: ReadonlyMap<number, SheetAxisFormatOverride>,
): Set<number> => {
  const changed = new Set<number>();
  if (previous === next) return changed;
  if (previous === null) {
    for (const sheet of next.keys()) changed.add(sheet);
    return changed;
  }
  for (const [sheet, override] of next) {
    if (!sameAxisFormatOverride(previous.get(sheet), override)) changed.add(sheet);
  }
  for (const sheet of previous.keys()) if (!next.has(sheet)) changed.add(sheet);
  return changed;
};

const sameHiddenRows = (
  before: ReadonlySet<number> | null,
  after: ReadonlySet<number>,
): boolean => {
  if (before === after) return true;
  if (!before || before.size !== after.size) return false;
  for (const row of before) {
    if (!after.has(row)) return false;
  }
  return true;
};

/** Mirrors a bound mount store's formats, tables, AutoFilter and page-setup
 *  slices into the engine as they change. */
export class StoreMirror {
  constructor(private readonly handle: WorkbookHandle) {
    this.tables = new TableMirror(handle);
    this.drawingObjects = new DrawingObjectMirror(handle);
  }

  private readonly tables: TableMirror;

  private readonly drawingObjects: DrawingObjectMirror;

  /** Optional mount store whose AutoFilter and page-setup slices are mirrored
   * into the engine as they change. Keeping this association at the adapter
   * boundary means every React/Vue/vanilla surface gets save-safe filters and
   * print settings without duplicating writeback calls. */
  private boundStore: SpreadsheetStore | null = null;

  /** Bridge owned by this mirror while `boundStore` is attached. The
   * registry removes it conditionally so an older mirror cannot detach a
   * newer mirror that borrowed the same workbook/store pair. */
  private formatHistoryBridge: FormatHistoryBridge | null = null;

  private unsubscribeBoundStore: (() => void) | null = null;

  /** Keep worksheet-indexed store state aligned when callers mutate the
   *  WorkbookHandle directly. Structure commands deliberately do not remap
   *  this map themselves: the handle event is the one source of truth for
   *  both direct and command-mediated mutations. */
  private unsubscribeWorkbook: (() => void) | null = null;

  private unsubscribeCalcSettings: (() => void) | null = null;

  private lastAutoFilterSignature = '';

  /** Hidden rows are part of the UI-owned filter result. Track their content
   * separately because advanced filters, slicers and reapply can change the
   * result while the AutoFilter definition stays byte-for-byte identical. */
  private lastHiddenRows: ReadonlySet<number> | null = null;

  private lastHiddenRowsSheet: number | null = null;

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

  /** Whole-sheet font overrides are a separate sparse map from cell formats.
   *  Track it independently so applying, undoing, or clearing an override
   *  flushes the native column/row styles even when no physical cell changed. */
  private lastSheetFontOverrides: ReadonlyMap<number, SheetFontOverride> | null = null;

  private lastSheetAxisOverrides: ReadonlyMap<number, SheetAxisFormatOverride> | null = null;

  /** Named styles are session metadata; changing the array can affect every
   *  cell that references one, so its identity is tracked independently. */
  private lastCustomCellStyles: State['format']['customCellStyles'] | null = null;

  /** Tab colours and outline summary positions feed `<sheetPr>`; both are
   *  compared per sheet because either map is replaced on every edit. */
  private lastSheetTabColors: ReadonlyMap<number, string> | undefined;

  private lastOutlineSummary: ReadonlyMap<number, OutlineSummary> | undefined;

  private storeSyncMuted = 0;

  attach(store: SpreadsheetStore | null): void {
    if (this.boundStore && this.formatHistoryBridge) {
      unregisterFormatHistoryBridge(this.boundStore, this.formatHistoryBridge);
    }
    this.formatHistoryBridge = null;
    this.unsubscribeBoundStore?.();
    this.unsubscribeBoundStore = null;
    this.unsubscribeWorkbook?.();
    this.unsubscribeWorkbook = null;
    this.unsubscribeCalcSettings?.();
    this.unsubscribeCalcSettings = null;
    if (this.boundStore) associatePaginationEngine(this.boundStore.getState(), null);
    this.boundStore = store;
    this.rebase();
    if (!store) return;
    this.formatHistoryBridge = createFormatHistoryBridge(
      this.handle,
      (sheet) => flushFormatToEngine(this.handle, store, sheet, { strict: true }),
      (keys) => captureSyncedFormatKeyMembership(this.handle, keys),
      (snapshot) => restoreSyncedFormatKeyMembership(this.handle, snapshot),
    );
    registerFormatHistoryBridge(store, this.formatHistoryBridge);
    associatePaginationEngine(store.getState(), this.handle);
    this.unsubscribeCalcSettings = bindCalcSettingsStore(this.handle, store);
    this.unsubscribeWorkbook = bindWholeSheetFontStore(this.handle, store, (mapIndex) => {
      remapSyncedFormatKeys(this.handle, mapIndex);
      remapActiveAxisFormatOverrides(this.handle, mapIndex);
      remapAxisFormatOverrides(store, mapIndex);
    });
    this.unsubscribeBoundStore = store.subscribe((state) => {
      // Tables go first: creating or removing one edits the filter ranges in
      // a nested notification, which mirrors every slice from the newer state.
      if (this.storeSyncMuted > 0) {
        this.tables.rebase(store.getState().tables);
        this.drawingObjects.rebase(store.getState());
      } else if (this.tables.sync(store)) return;
      else if (this.drawingObjects.sync(store)) return;
      associatePaginationEngine(state, this.handle);
      const nextFormats = state.format.formats;
      const formatSheets = changedFormatSheets(this.lastFormatEntries, nextFormats);
      const nextSheetFontOverrides = state.format.sheetFontOverrides ?? new Map();
      for (const sheet of changedSheetFontOverrideSheets(
        this.lastSheetFontOverrides,
        nextSheetFontOverrides,
      )) {
        formatSheets.add(sheet);
      }
      const nextSheetAxisOverrides = state.format.sheetAxisOverrides ?? new Map();
      for (const sheet of changedAxisFormatOverrideSheets(
        this.lastSheetAxisOverrides,
        nextSheetAxisOverrides,
      )) {
        formatSheets.add(sheet);
      }
      const stylesChanged = (state.format.customCellStyles ?? null) !== this.lastCustomCellStyles;
      if (stylesChanged) {
        // Only sheets whose cells reference a style are rewritten; flushing any
        // other sheet would clear validations and hyperlinks it never hydrated.
        for (const sheet of sheetsWithCellStyles(nextFormats)) formatSheets.add(sheet);
      }
      // Advance all baselines even while hydration is muted. The state that
      // hydration imports is the new engine baseline, not a pending UI edit.
      this.lastFormatEntries = nextFormats;
      this.lastSheetFontOverrides = nextSheetFontOverrides;
      this.lastSheetAxisOverrides = nextSheetAxisOverrides;
      this.lastCustomCellStyles = state.format.customCellStyles ?? null;
      const sheetPrSheets = changedMapKeys(this.lastSheetTabColors, state.layout.sheetTabColors);
      for (const sheet of changedMapKeys(
        this.lastOutlineSummary,
        state.layout.outlineSummary,
        sameOutlineSummary,
      )) {
        sheetPrSheets.add(sheet);
      }
      this.lastSheetTabColors = state.layout.sheetTabColors;
      this.lastOutlineSummary = state.layout.outlineSummary;
      const nextFilter = autoFilterSignature(state);
      const filterChanged = nextFilter !== this.lastAutoFilterSignature;
      if (filterChanged) this.lastAutoFilterSignature = nextFilter;
      const hiddenRowsSheetChanged = this.lastHiddenRowsSheet !== state.data.sheetIndex;
      const hiddenRowsChanged =
        !hiddenRowsSheetChanged && !sameHiddenRows(this.lastHiddenRows, state.layout.hiddenRows);
      const previousHiddenRows = this.lastHiddenRows;
      this.lastHiddenRows = state.layout.hiddenRows;
      this.lastHiddenRowsSheet = state.data.sheetIndex;
      const nextPageSetup = state.pageSetup.setupBySheet;
      const pageSetupChanged = nextPageSetup !== this.lastPageSetupMap;
      if (pageSetupChanged) this.lastPageSetupMap = nextPageSetup;
      if (this.storeSyncMuted > 0) return;
      for (const sheet of formatSheets) {
        if (sheet < this.handle.sheetCount) flushFormatToEngine(this.handle, store, sheet);
      }
      if (stylesChanged) syncNamedStyleCatalogToEngine(this.handle, store);
      for (const sheet of sheetPrSheets) {
        if (sheet < this.handle.sheetCount) syncSheetPrToEngine(this.handle, store, sheet);
      }
      if (filterChanged) syncAutoFilterToEngine(this.handle, state, state.data.sheetIndex, store);
      if (hiddenRowsChanged && previousHiddenRows && !filterChanged) {
        syncHiddenRowsToEngine(
          this.handle,
          state.data.sheetIndex,
          previousHiddenRows,
          state.layout.hiddenRows,
          hiddenRowsOf(state.layout),
        );
      }
      if (pageSetupChanged) syncPageSetupToEngine(this.handle, store, state.data.sheetIndex);
    });
  }

  detach(expectedStore: SpreadsheetStore): void {
    if (this.boundStore !== expectedStore) return;
    this.attach(null);
  }

  /** Write the bound store's active sheet, selection and scroll origin to the
   *  engine so a save captures them. Called ahead of each save rather than on
   *  every scroll or selection change. */
  flushViewState(): void {
    if (this.boundStore) syncViewStateToEngine(this.handle, this.boundStore);
  }

  /** Re-read the bound store's active sheet after a recalc no UI command ran. */
  refreshCells(): void {
    const store = this.boundStore;
    if (store) mutators.replaceCells(store, this.handle.cells(store.getState().data.sheetIndex));
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
    this.lastSheetFontOverrides = state?.format.sheetFontOverrides ?? null;
    this.lastSheetAxisOverrides = state?.format.sheetAxisOverrides ?? null;
    this.lastCustomCellStyles = state?.format.customCellStyles ?? null;
    this.lastSheetTabColors = state?.layout.sheetTabColors;
    this.lastOutlineSummary = state?.layout.outlineSummary;
    this.lastAutoFilterSignature = state ? autoFilterSignature(state) : '';
    this.lastHiddenRows = state?.layout.hiddenRows ?? null;
    this.lastHiddenRowsSheet = state?.data.sheetIndex ?? null;
    this.lastPageSetupMap = state?.pageSetup.setupBySheet ?? null;
    this.tables.rebase(state?.tables ?? null);
    this.drawingObjects.rebase(state ?? null);
  }
}
