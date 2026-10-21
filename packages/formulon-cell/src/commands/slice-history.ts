import { parseAddrKey } from '../engine/address.js';
import {
  type FormatHistoryTrackingSnapshot,
  formatHistoryBridgeFor,
  type NativeFormatSnapshot,
} from '../engine/format-history-bridge.js';
import { syncLayoutToEngine } from '../engine/layout-sync.js';
import type { Range, SheetAxisFormatOverride, SheetFontOverride } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type {
  CellFormat,
  ConditionalRule,
  CustomCellStyle,
  LayoutSlice,
  PageMargins,
  PageSetup,
  SessionIllustration,
  SlicerSpec,
  Sparkline,
  SpreadsheetStore,
  State,
} from '../store/store.js';
import type {
  CustomTableStyle,
  PivotTableStyleAssignment,
  TableOverlay,
} from './format-as-table.js';
import type { History } from './history.js';
import type { OperationIntent } from './interaction-policy.js';

/* ---------- Snapshot helpers ---------- */

/** Capture the entire format map. Sufficient for undo since the map is sparse
 *  (only explicitly formatted cells have entries). */
export interface FormatSnapshot {
  formats: Map<string, CellFormat>;
  sheetFontOverrides?: Map<number, SheetFontOverride>;
  sheetAxisOverrides?: Map<number, SheetAxisFormatOverride>;
  customCellStyles: CustomCellStyle[];
}
const cloneSheetFontAxisSnapshot = (
  snapshot: SheetFontOverride['baseline'],
): SheetFontOverride['baseline'] =>
  snapshot
    ? {
        columns: snapshot.columns.map((entry) => ({ ...entry })),
        rows: snapshot.rows.map((entry) => ({ ...entry })),
      }
    : undefined;

const cloneSheetFontOverride = (override: SheetFontOverride): SheetFontOverride => ({
  patch: { ...override.patch },
  ...(override.clearAxes ? { clearAxes: true } : {}),
  ...(override.baseline ? { baseline: cloneSheetFontAxisSnapshot(override.baseline) } : {}),
});

const cloneSheetAxisOverride = (override: SheetAxisFormatOverride): SheetAxisFormatOverride => ({
  baseline: cloneSheetFontAxisSnapshot(override.baseline) ?? { columns: [], rows: [] },
  desired: cloneSheetFontAxisSnapshot(override.desired) ?? { columns: [], rows: [] },
  ...(override.directXfs
    ? {
        directXfs: override.directXfs.map((entry) => ({
          key: entry.key,
          before: { ...entry.before },
          afterXfIndex: entry.afterXfIndex,
        })),
      }
    : {}),
  ...(override.fontPatch ? { fontPatch: { ...override.fontPatch } } : {}),
});

const cloneCustomCellStyle = (style: CustomCellStyle): CustomCellStyle => ({
  ...style,
  format: { ...style.format },
});

export function captureFormatSnapshot(state: State): FormatSnapshot {
  return {
    formats: new Map(state.format.formats),
    sheetFontOverrides: new Map(
      [...(state.format.sheetFontOverrides ?? new Map())].map(([sheet, override]) => [
        sheet,
        cloneSheetFontOverride(override),
      ]),
    ),
    sheetAxisOverrides: new Map(
      [...(state.format.sheetAxisOverrides ?? new Map())].map(([sheet, override]) => [
        sheet,
        cloneSheetAxisOverride(override),
      ]),
    ),
    customCellStyles: (state.format.customCellStyles ?? []).map(cloneCustomCellStyle),
  };
}

export function applyFormatSnapshot(store: SpreadsheetStore, snap: FormatSnapshot): void {
  store.setState((s) => ({
    ...s,
    format: {
      formats: new Map(snap.formats),
      sheetFontOverrides: new Map(
        [...(snap.sheetFontOverrides ?? new Map())].map(([sheet, override]) => [
          sheet,
          cloneSheetFontOverride(override),
        ]),
      ),
      sheetAxisOverrides: new Map(
        [...(snap.sheetAxisOverrides ?? new Map())].map(([sheet, override]) => [
          sheet,
          cloneSheetAxisOverride(override),
        ]),
      ),
      customCellStyles: snap.customCellStyles.map(cloneCustomCellStyle),
    },
  }));
}

const formatSnapshotKey = (snap: FormatSnapshot): string =>
  JSON.stringify({
    formats: [...snap.formats.entries()].sort(([a], [b]) => a.localeCompare(b)),
    sheetFontOverrides: [...(snap.sheetFontOverrides ?? new Map()).entries()].sort(
      ([a], [b]) => a - b,
    ),
    sheetAxisOverrides: [...(snap.sheetAxisOverrides ?? new Map()).entries()].sort(
      ([a], [b]) => a - b,
    ),
    customCellStyles: snap.customCellStyles,
  });

const sameFormatSnapshot = (a: FormatSnapshot, b: FormatSnapshot): boolean =>
  a.formats.size === b.formats.size &&
  (a.sheetFontOverrides?.size ?? 0) === (b.sheetFontOverrides?.size ?? 0) &&
  (a.sheetAxisOverrides?.size ?? 0) === (b.sheetAxisOverrides?.size ?? 0) &&
  a.customCellStyles.length === b.customCellStyles.length &&
  formatSnapshotKey(a) === formatSnapshotKey(b);

const jsonSnapshot = (value: unknown): string => JSON.stringify(value);

const sameFormatValue = (a: CellFormat | undefined, b: CellFormat | undefined): boolean =>
  jsonSnapshot(a) === jsonSnapshot(b);

const changedFormatKeys = (before: FormatSnapshot, after: FormatSnapshot): Set<string> => {
  const changed = new Set<string>();
  for (const key of new Set([...before.formats.keys(), ...after.formats.keys()])) {
    if (!sameFormatValue(before.formats.get(key), after.formats.get(key))) changed.add(key);
  }
  return changed;
};

const changedSheetFontOverrideKeys = (
  before: FormatSnapshot,
  after: FormatSnapshot,
): Set<number> => {
  const changed = new Set<number>();
  const beforeOverrides = before.sheetFontOverrides ?? new Map();
  const afterOverrides = after.sheetFontOverrides ?? new Map();
  for (const sheet of new Set([...beforeOverrides.keys(), ...afterOverrides.keys()])) {
    if (jsonSnapshot(beforeOverrides.get(sheet)) !== jsonSnapshot(afterOverrides.get(sheet))) {
      changed.add(sheet);
    }
  }
  return changed;
};

const changedSheetAxisOverrideKeys = (
  before: FormatSnapshot,
  after: FormatSnapshot,
): Set<number> => {
  const changed = new Set<number>();
  const beforeOverrides = before.sheetAxisOverrides ?? new Map();
  const afterOverrides = after.sheetAxisOverrides ?? new Map();
  for (const sheet of new Set([...beforeOverrides.keys(), ...afterOverrides.keys()])) {
    if (jsonSnapshot(beforeOverrides.get(sheet)) !== jsonSnapshot(afterOverrides.get(sheet))) {
      changed.add(sheet);
    }
  }
  return changed;
};

const directKeysForAxisOverrides = (
  snapshot: FormatSnapshot,
  sheets: Iterable<number>,
): Set<string> => {
  const keys = new Set<string>();
  const overrides = snapshot.sheetAxisOverrides ?? new Map();
  for (const sheet of sheets) {
    for (const entry of overrides.get(sheet)?.directXfs ?? []) keys.add(entry.key);
  }
  return keys;
};

const cloneCellFormat = (format: CellFormat): CellFormat =>
  JSON.parse(JSON.stringify(format)) as CellFormat;

const formatPatch = (
  before: FormatSnapshot,
  keys: Iterable<string>,
): Map<string, CellFormat | undefined> => {
  const patch = new Map<string, CellFormat | undefined>();
  for (const key of keys) {
    const value = before.formats.get(key);
    patch.set(key, value === undefined ? undefined : cloneCellFormat(value));
  }
  return patch;
};

const sheetFontPatch = (
  before: FormatSnapshot,
  sheets: Iterable<number>,
): Map<number, SheetFontOverride | undefined> => {
  const patch = new Map<number, SheetFontOverride | undefined>();
  const overrides = before.sheetFontOverrides ?? new Map();
  for (const sheet of sheets) {
    const value = overrides.get(sheet);
    patch.set(sheet, value === undefined ? undefined : cloneSheetFontOverride(value));
  }
  return patch;
};

const sheetAxisPatch = (
  before: FormatSnapshot,
  sheets: Iterable<number>,
): Map<number, SheetAxisFormatOverride | undefined> => {
  const patch = new Map<number, SheetAxisFormatOverride | undefined>();
  const overrides = before.sheetAxisOverrides ?? new Map();
  for (const sheet of sheets) {
    const value = overrides.get(sheet);
    patch.set(sheet, value === undefined ? undefined : cloneSheetAxisOverride(value));
  }
  return patch;
};

const formatsWithCellStyle = (snapshot: FormatSnapshot): Set<string> => {
  const keys = new Set<string>();
  for (const [key, format] of snapshot.formats) {
    if (format.cellStyle !== undefined) keys.add(key);
  }
  return keys;
};

const sheetsForFormatKeys = (keys: Iterable<string>): Set<number> => {
  const sheets = new Set<number>();
  for (const key of keys) {
    const addr = parseAddrKey(key);
    if (addr && addr.sheet >= 0) sheets.add(addr.sheet);
  }
  return sheets;
};

const sheetsForCellStyles = (before: FormatSnapshot, after: FormatSnapshot): Set<number> => {
  const sheets = sheetsForFormatKeys(
    new Set([...formatsWithCellStyle(before), ...formatsWithCellStyle(after)]),
  );
  if (sheets.size === 0) sheets.add(0);
  return sheets;
};

const applyScopedFormatChange = (
  store: SpreadsheetStore,
  formats: Map<string, CellFormat | undefined>,
  sheetFonts: Map<number, SheetFontOverride | undefined>,
  sheetAxes: Map<number, SheetAxisFormatOverride | undefined>,
  customCellStyles: CustomCellStyle[] | undefined,
): void => {
  store.setState((state) => {
    const nextFormats = new Map(state.format.formats);
    for (const [key, value] of formats) {
      if (value === undefined) nextFormats.delete(key);
      else nextFormats.set(key, cloneCellFormat(value));
    }
    const nextSheetFonts = new Map(state.format.sheetFontOverrides ?? new Map());
    for (const [sheet, value] of sheetFonts) {
      if (value === undefined) nextSheetFonts.delete(sheet);
      else nextSheetFonts.set(sheet, cloneSheetFontOverride(value));
    }
    const nextSheetAxes = new Map(state.format.sheetAxisOverrides ?? new Map());
    for (const [sheet, value] of sheetAxes) {
      if (value === undefined) nextSheetAxes.delete(sheet);
      else nextSheetAxes.set(sheet, cloneSheetAxisOverride(value));
    }
    return {
      ...state,
      format: {
        ...state.format,
        formats: nextFormats,
        sheetFontOverrides: nextSheetFonts,
        sheetAxisOverrides: nextSheetAxes,
        ...(customCellStyles === undefined
          ? {}
          : { customCellStyles: customCellStyles.map(cloneCustomCellStyle) }),
      },
    };
  });
};

const mapSnapshotKey = <T>(snap: Map<string, T>): string =>
  JSON.stringify([...snap.entries()].sort(([a], [b]) => a.localeCompare(b)));

const sameMapSnapshot = <T>(a: Map<string, T>, b: Map<string, T>): boolean =>
  a.size === b.size && mapSnapshotKey(a) === mapSnapshotKey(b);

const sameJsonSnapshot = <T>(a: readonly T[], b: readonly T[]): boolean =>
  a.length === b.length && JSON.stringify(a) === JSON.stringify(b);

/** How one store slice is snapshotted, restored and compared for history. */
interface SliceCodec<T> {
  capture: (state: State) => T;
  apply: (store: SpreadsheetStore, snap: T, origin?: 'undo' | 'redo') => void;
  same: (a: T, b: T) => boolean;
  repeat?: () => void;
  intent?: OperationIntent;
  inverseIntent?: OperationIntent;
}

/** Run `mutate`, capturing the slice before and after, and push one entry
 *  unless the slice is unchanged. Runs `mutate` bare when `history` is null or
 *  replaying. */
function recordSliceChange<T>(
  history: History | null,
  store: SpreadsheetStore,
  mutate: () => void,
  codec: SliceCodec<T>,
): void {
  if (!history || history.isReplaying()) {
    mutate();
    return;
  }
  const before = codec.capture(store.getState());
  mutate();
  const after = codec.capture(store.getState());
  if (codec.same(before, after)) return;
  history.push({
    undo: () => codec.apply(store, before, 'undo'),
    redo: () => codec.apply(store, after, 'redo'),
    repeat: codec.repeat,
    ...(codec.intent ? { intent: codec.intent } : {}),
    ...(codec.inverseIntent ? { inverseIntent: codec.inverseIntent } : {}),
  });
}

const sameTablesSnapshot = (a: TablesSnapshot, b: TablesSnapshot): boolean =>
  JSON.stringify(a) === JSON.stringify(b);

const cloneTableOverlay = (table: TableOverlay): TableOverlay => ({
  ...table,
  range: { ...table.range },
});

export interface TablesSnapshot {
  tables: TableOverlay[];
  customTableStyles: CustomTableStyle[];
  customPivotTableStyles: CustomTableStyle[];
  pivotTableStyles: PivotTableStyleAssignment[];
}

const cloneCustomTableStyle = (style: CustomTableStyle): CustomTableStyle => ({ ...style });
const clonePivotTableStyle = (style: PivotTableStyleAssignment): PivotTableStyleAssignment => ({
  ...style,
});

export function captureTableOverlaysSnapshot(state: State): TablesSnapshot {
  return {
    tables: state.tables.tables.map(cloneTableOverlay),
    customTableStyles: (state.tables.customTableStyles ?? []).map(cloneCustomTableStyle),
    customPivotTableStyles: (state.tables.customPivotTableStyles ?? []).map(cloneCustomTableStyle),
    pivotTableStyles: (state.tables.pivotTableStyles ?? []).map(clonePivotTableStyle),
  };
}

export function applyTableOverlaysSnapshot(
  store: SpreadsheetStore,
  snap: TablesSnapshot | readonly TableOverlay[],
): void {
  const normalized: TablesSnapshot = Array.isArray(snap)
    ? {
        tables: (snap as readonly TableOverlay[]).map(cloneTableOverlay),
        customTableStyles: [],
        customPivotTableStyles: [],
        pivotTableStyles: [],
      }
    : (snap as TablesSnapshot);
  store.setState((s) => ({
    ...s,
    tables: {
      tables: normalized.tables.map(cloneTableOverlay),
      customTableStyles: normalized.customTableStyles.map(cloneCustomTableStyle),
      customPivotTableStyles: normalized.customPivotTableStyles.map(cloneCustomTableStyle),
      pivotTableStyles: normalized.pivotTableStyles.map(clonePivotTableStyle),
    },
  }));
}

export function recordTablesChange(
  history: History | null,
  store: SpreadsheetStore,
  mutate: () => void,
): void {
  recordSliceChange(history, store, mutate, {
    capture: captureTableOverlaysSnapshot,
    apply: applyTableOverlaysSnapshot,
    same: sameTablesSnapshot,
  });
}

const cloneSessionIllustration = (item: SessionIllustration): SessionIllustration => ({ ...item });

export function captureIllustrationsSnapshot(state: State): SessionIllustration[] {
  return state.illustrations.illustrations.map(cloneSessionIllustration);
}

export function applyIllustrationsSnapshot(
  store: SpreadsheetStore,
  snap: readonly SessionIllustration[],
): void {
  store.setState((s) => ({
    ...s,
    illustrations: { illustrations: snap.map(cloneSessionIllustration) },
  }));
}

export function recordIllustrationsChange(
  history: History | null,
  store: SpreadsheetStore,
  mutate: () => void,
): void {
  recordSliceChange(history, store, mutate, {
    capture: captureIllustrationsSnapshot,
    apply: applyIllustrationsSnapshot,
    same: sameJsonSnapshot,
  });
}

export interface LayoutSnapshot {
  colWidths: Map<number, number>;
  rowHeights: Map<number, number>;
  /** Engine-sync input only: freezing is not an undo step, so history neither
   *  compares nor restores these. */
  freezeRows: number;
  freezeCols: number;
  hiddenRows: Set<number>;
  hiddenCols: Set<number>;
  outlineRows: Map<number, number>;
  outlineCols: Map<number, number>;
  outlineRowGutter: number;
  outlineColGutter: number;
  hiddenSheets: Set<number>;
  veryHiddenSheets: Set<number>;
  sheetTabColors: Map<number, string>;
}

export function captureLayoutSnapshot(state: State): LayoutSnapshot {
  return {
    colWidths: new Map(state.layout.colWidths),
    rowHeights: new Map(state.layout.rowHeights),
    freezeRows: state.layout.freezeRows,
    freezeCols: state.layout.freezeCols,
    hiddenRows: new Set(state.layout.hiddenRows),
    hiddenCols: new Set(state.layout.hiddenCols),
    outlineRows: new Map(state.layout.outlineRows),
    outlineCols: new Map(state.layout.outlineCols),
    outlineRowGutter: state.layout.outlineRowGutter,
    outlineColGutter: state.layout.outlineColGutter,
    hiddenSheets: new Set(state.layout.hiddenSheets),
    veryHiddenSheets: new Set(state.layout.veryHiddenSheets),
    sheetTabColors: new Map(state.layout.sheetTabColors),
  };
}

export function applyLayoutSnapshot(store: SpreadsheetStore, snap: LayoutSnapshot): void {
  store.setState((s) => ({
    ...s,
    layout: {
      ...s.layout,
      colWidths: new Map(snap.colWidths),
      rowHeights: new Map(snap.rowHeights),
      hiddenRows: new Set(snap.hiddenRows),
      hiddenCols: new Set(snap.hiddenCols),
      outlineRows: new Map(snap.outlineRows),
      outlineCols: new Map(snap.outlineCols),
      outlineRowGutter: snap.outlineRowGutter,
      outlineColGutter: snap.outlineColGutter,
      hiddenSheets: new Set(snap.hiddenSheets),
      veryHiddenSheets: new Set(snap.veryHiddenSheets),
      sheetTabColors: new Map(snap.sheetTabColors),
    } as LayoutSlice,
  }));
}

const sameNumberMap = (a: Map<number, number>, b: Map<number, number>): boolean =>
  a.size === b.size && [...a].every(([key, value]) => b.get(key) === value);

const sameNumberSet = (a: Set<number>, b: Set<number>): boolean =>
  a.size === b.size && [...a].every((value) => b.has(value));

const sameNumberStringMap = (a: Map<number, string>, b: Map<number, string>): boolean =>
  a.size === b.size && [...a].every(([key, value]) => b.get(key) === value);

const sameLayoutSnapshot = (a: LayoutSnapshot, b: LayoutSnapshot): boolean =>
  sameNumberMap(a.colWidths, b.colWidths) &&
  sameNumberMap(a.rowHeights, b.rowHeights) &&
  sameNumberSet(a.hiddenRows, b.hiddenRows) &&
  sameNumberSet(a.hiddenCols, b.hiddenCols) &&
  sameNumberMap(a.outlineRows, b.outlineRows) &&
  sameNumberMap(a.outlineCols, b.outlineCols) &&
  a.outlineRowGutter === b.outlineRowGutter &&
  a.outlineColGutter === b.outlineColGutter &&
  sameNumberSet(a.hiddenSheets, b.hiddenSheets) &&
  sameNumberSet(a.veryHiddenSheets, b.veryHiddenSheets) &&
  sameNumberStringMap(a.sheetTabColors, b.sheetTabColors);

/** Run `mutate`, capturing the format slice before and after, pushing one
 *  entry. No-op when `history` is null. */
export function recordFormatChange(
  history: History | null,
  store: SpreadsheetStore,
  mutate: () => void,
  opts: { repeat?: () => void } = {},
): void {
  const bridge = history && !history.isReplaying() ? formatHistoryBridgeFor(store) : undefined;
  if (bridge) {
    recordBridgedFormatChange(history as History, store, mutate, opts, bridge);
    return;
  }
  recordSliceChange(history, store, mutate, {
    capture: captureFormatSnapshot,
    apply: applyFormatSnapshot,
    same: sameFormatSnapshot,
    repeat: opts.repeat,
  });
}

/** Record a mounted format mutation while keeping native cell XFs in the same
 * state as the store history entry. The store mirror is muted around the
 * mutation and every replay; strict publication happens explicitly before
 * the exact native XF snapshot is restored. */
function recordBridgedFormatChange(
  history: History,
  store: SpreadsheetStore,
  mutate: () => void,
  opts: { repeat?: () => void },
  bridge: import('../engine/format-history-bridge.js').FormatHistoryBridge,
): void {
  const before = captureFormatSnapshot(store.getState());
  let after: FormatSnapshot | undefined;
  let changedKeys = new Set<string>();
  let changedFontSheets = new Set<number>();
  let changedAxisSheets = new Set<number>();
  let customStylesChanged = false;
  let affectedSheets = new Set<number>();
  const nativeKeys = new Set<string>();
  let beforeNative: NativeFormatSnapshot = new Map();
  let beforeTracking: FormatHistoryTrackingSnapshot | undefined;
  let publicationAttempted = false;

  const flushAndRestore = (
    native: NativeFormatSnapshot,
    tracking: FormatHistoryTrackingSnapshot,
    flush: boolean,
  ): void => {
    const errors: unknown[] = [];
    if (flush) {
      try {
        bridge.flush(affectedSheets);
      } catch (error) {
        errors.push(error);
      }
    }
    // A partial flush may have already published one or more cells. Always
    // try the exact snapshot even when the retry flush failed, otherwise the
    // first successful cell can remain stranded in the engine. Restore the
    // high-water membership after native publication so a later unrelated
    // flush cannot reset a cell that was already restored here.
    let nativeRestored = false;
    try {
      bridge.restore(native);
      nativeRestored = true;
    } catch (error) {
      errors.push(error);
    }
    // Keep newly published membership bits when exact native restoration is
    // incomplete. Those bits describe the stranded native publication and
    // must remain visible to the next cleanup attempt; clear them only after
    // every native XF in the checkpoint was restored successfully.
    if (nativeRestored) {
      try {
        bridge.restoreTracking(tracking);
      } catch (error) {
        errors.push(error);
      }
    }
    if (errors.length === 1) throw errors[0];
    if (errors.length > 1) {
      throw new AggregateError(errors, 'Format history recovery failed', {
        cause: errors[0],
      });
    }
  };

  try {
    bridge.withSyncMuted(mutate);
    after = captureFormatSnapshot(store.getState());
    changedKeys = changedFormatKeys(before, after);
    changedFontSheets = changedSheetFontOverrideKeys(before, after);
    changedAxisSheets = changedSheetAxisOverrideKeys(before, after);
    customStylesChanged =
      jsonSnapshot(before.customCellStyles) !== jsonSnapshot(after.customCellStyles);
    if (
      changedKeys.size === 0 &&
      changedFontSheets.size === 0 &&
      changedAxisSheets.size === 0 &&
      !customStylesChanged
    )
      return;

    affectedSheets = sheetsForFormatKeys(changedKeys);
    for (const sheet of changedFontSheets) affectedSheets.add(sheet);
    for (const sheet of changedAxisSheets) affectedSheets.add(sheet);
    if (customStylesChanged) {
      for (const sheet of sheetsForCellStyles(before, after)) affectedSheets.add(sheet);
      for (const key of formatsWithCellStyle(before)) nativeKeys.add(key);
      for (const key of formatsWithCellStyle(after)) nativeKeys.add(key);
    }
    for (const key of changedKeys) nativeKeys.add(key);
    for (const key of directKeysForAxisOverrides(before, changedAxisSheets)) nativeKeys.add(key);
    for (const key of directKeysForAxisOverrides(after, changedAxisSheets)) nativeKeys.add(key);

    // The mirror stayed muted, so the engine is still at the pre-action XF.
    beforeNative = bridge.capture(nativeKeys);
    beforeTracking = bridge.captureTracking(nativeKeys);
    publicationAttempted = true;
    bridge.flush(affectedSheets);
    const afterNative = bridge.capture(nativeKeys);

    const beforeFormats = formatPatch(before, changedKeys);
    const afterFormats = formatPatch(after, changedKeys);
    const beforeSheetFonts = sheetFontPatch(before, changedFontSheets);
    const afterSheetFonts = sheetFontPatch(after, changedFontSheets);
    const beforeSheetAxes = sheetAxisPatch(before, changedAxisSheets);
    const afterSheetAxes = sheetAxisPatch(after, changedAxisSheets);
    const beforeStyles = customStylesChanged
      ? before.customCellStyles.map(cloneCustomCellStyle)
      : undefined;
    const afterStyles = customStylesChanged
      ? after.customCellStyles.map(cloneCustomCellStyle)
      : undefined;

    const replay = (
      formats: Map<string, CellFormat | undefined>,
      sheetFonts: Map<number, SheetFontOverride | undefined>,
      sheetAxes: Map<number, SheetAxisFormatOverride | undefined>,
      customStyles: CustomCellStyle[] | undefined,
      native: NativeFormatSnapshot,
    ): void => {
      const immediate = captureFormatSnapshot(store.getState());
      const immediateNative = bridge.capture(nativeKeys);
      const immediateTracking = bridge.captureTracking(nativeKeys);
      let replayPublicationAttempted = false;
      try {
        bridge.withSyncMuted(() => {
          applyScopedFormatChange(store, formats, sheetFonts, sheetAxes, customStyles);
          replayPublicationAttempted = true;
          bridge.flush(affectedSheets);
          bridge.restore(native);
        });
      } catch (error) {
        try {
          bridge.withSyncMuted(() => {
            applyFormatSnapshot(store, immediate);
            flushAndRestore(immediateNative, immediateTracking, replayPublicationAttempted);
          });
        } catch (rollbackError) {
          throw new AggregateError(
            [error, rollbackError],
            'Format history replay failed and its rollback failed',
            { cause: error },
          );
        }
        throw error;
      }
    };

    history.push({
      undo: () =>
        replay(beforeFormats, beforeSheetFonts, beforeSheetAxes, beforeStyles, beforeNative),
      redo: () => replay(afterFormats, afterSheetFonts, afterSheetAxes, afterStyles, afterNative),
      repeat: opts.repeat,
    });
  } catch (error) {
    try {
      bridge.withSyncMuted(() => {
        applyFormatSnapshot(store, before);
        if (publicationAttempted && beforeTracking) {
          flushAndRestore(beforeNative, beforeTracking, true);
        }
      });
    } catch (rollbackError) {
      throw new AggregateError(
        [error, rollbackError],
        'Format history change failed and its rollback failed',
        { cause: error },
      );
    }
    throw error;
  }
}

/** Record a format command whose mutation derives its target from the current
 *  store state. It is therefore safe to repeat with F4 after selection moves.
 *  Callers that captured a concrete range (paste/fill) must continue using
 *  `recordFormatChange` directly. */
export function recordRepeatableFormatChange(
  history: History | null,
  store: SpreadsheetStore,
  mutate: () => void,
): void {
  const repeat = (): void => recordRepeatableFormatChange(history, store, mutate);
  recordFormatChange(history, store, mutate, { repeat });
  history?.setRepeat(repeat);
}

/** Record a format command that was handed a concrete range, together with the
 *  repeat that should run in its place when F4 fires. Use this where the
 *  mutation cannot simply be re-run — the caller supplies a `repeat` that
 *  re-targets the current selection itself. */
export function recordFormatChangeWithRepeat(
  history: History | null,
  store: SpreadsheetStore,
  mutate: () => void,
  repeat: () => void,
): void {
  recordFormatChange(history, store, mutate, { repeat });
  history?.setRepeat(repeat);
}

export function recordLayoutChange(
  history: History | null,
  store: SpreadsheetStore,
  mutate: () => void,
): void {
  recordSliceChange(history, store, mutate, {
    capture: captureLayoutSnapshot,
    apply: applyLayoutSnapshot,
    same: sameLayoutSnapshot,
  });
}

const cloneConditionalRule = (rule: ConditionalRule): ConditionalRule => {
  const out = { ...rule, range: { ...rule.range } } as ConditionalRule;
  if ('apply' in out) out.apply = { ...out.apply };
  if ('stops' in out) {
    out.stops = [...out.stops] as [string, string] | [string, string, string];
  }
  return out;
};

export function captureConditionalRulesSnapshot(state: State): ConditionalRule[] {
  return state.conditional.rules.map(cloneConditionalRule);
}

export function applyConditionalRulesSnapshot(
  store: SpreadsheetStore,
  snap: readonly ConditionalRule[],
): void {
  store.setState((s) => ({
    ...s,
    conditional: { rules: snap.map(cloneConditionalRule) },
  }));
}

export function recordConditionalRulesChange(
  history: History | null,
  store: SpreadsheetStore,
  mutate: () => void,
): void {
  recordSliceChange(history, store, mutate, {
    capture: captureConditionalRulesSnapshot,
    apply: applyConditionalRulesSnapshot,
    same: sameJsonSnapshot,
  });
}

/** Engine-aware layout change. Same semantics as `recordLayoutChange` but
 *  the captured before/after pair is also pushed to the workbook engine for
 *  the active sheet, both at apply time and on every undo/redo replay.
 *  Skipped (including the engine sync) when `wb` is null. The per-method
 *  calls inside `syncLayoutToEngine` short-circuit on each capability flag,
 *  so engines that only support a subset still work. */
export function recordLayoutChangeWithEngine(
  history: History | null,
  store: SpreadsheetStore,
  wb: WorkbookHandle | null,
  mutate: () => void,
): void {
  if (!wb) {
    recordLayoutChange(history, store, mutate);
    return;
  }
  const sheet = store.getState().data.sheetIndex;
  const before = captureLayoutSnapshot(store.getState());
  mutate();
  const after = captureLayoutSnapshot(store.getState());
  if (sameLayoutSnapshot(before, after)) return;
  syncLayoutToEngine(wb, store.getState().layout, sheet, before, after);
  if (!history || history.isReplaying()) return;
  history.push({
    undo: () => {
      applyLayoutSnapshot(store, before);
      syncLayoutToEngine(wb, store.getState().layout, sheet, after, before);
    },
    redo: () => {
      applyLayoutSnapshot(store, after);
      syncLayoutToEngine(wb, store.getState().layout, sheet, before, after);
    },
  });
}

export interface MergesSnapshot {
  byAnchor: Map<string, Range>;
  byCell: Map<string, string>;
}

export function captureMergesSnapshot(state: State): MergesSnapshot {
  return {
    byAnchor: new Map(state.merges.byAnchor),
    byCell: new Map(state.merges.byCell),
  };
}

export function applyMergesSnapshot(store: SpreadsheetStore, snap: MergesSnapshot): void {
  store.setState((s) => ({
    ...s,
    merges: { byAnchor: new Map(snap.byAnchor), byCell: new Map(snap.byCell) },
  }));
}

const rangeKey = (range: Range): string =>
  `${range.sheet}:${range.r0}:${range.c0}:${range.r1}:${range.c1}`;

const sameStringMap = (a: Map<string, string>, b: Map<string, string>): boolean =>
  a.size === b.size && [...a].every(([key, value]) => b.get(key) === value);

const sameRangeMap = (a: Map<string, Range>, b: Map<string, Range>): boolean =>
  a.size === b.size &&
  [...a].every(([key, range]) => {
    const other = b.get(key);
    return other !== undefined && rangeKey(range) === rangeKey(other);
  });

const sameMergesSnapshot = (a: MergesSnapshot, b: MergesSnapshot): boolean =>
  sameRangeMap(a.byAnchor, b.byAnchor) && sameStringMap(a.byCell, b.byCell);

export function recordMergesChange(
  history: History | null,
  store: SpreadsheetStore,
  mutate: () => void,
): void {
  recordSliceChange(history, store, mutate, {
    capture: captureMergesSnapshot,
    apply: applyMergesSnapshot,
    same: sameMergesSnapshot,
  });
}

/** Engine-aware merges change. Mirrors the post-mutate state into the workbook
 *  via `clearMerges` + per-anchor `addMerge`. Both apply and undo/redo go
 *  through the same path, so the engine snapshot stays in lockstep with the
 *  store. No-op against the engine when `wb` is null or `capabilities.merges`
 *  is off. */
export function recordMergesChangeWithEngine(
  history: History | null,
  store: SpreadsheetStore,
  wb: WorkbookHandle | null,
  sheet: number,
  mutate: () => void,
): void {
  const sync = (snap: MergesSnapshot): void => {
    if (!wb?.capabilities.merges) return;
    wb.engineClearMerges(sheet);
    for (const r of snap.byAnchor.values()) {
      if (r.sheet !== sheet) continue;
      wb.engineAddMerge(sheet, r);
    }
  };
  const before = captureMergesSnapshot(store.getState());
  mutate();
  const after = captureMergesSnapshot(store.getState());
  if (sameMergesSnapshot(before, after)) return;
  sync(after);
  if (!history || history.isReplaying()) return;
  history.push({
    undo: () => {
      applyMergesSnapshot(store, before);
      sync(before);
    },
    redo: () => {
      applyMergesSnapshot(store, after);
      sync(after);
    },
  });
}

export function captureSparklineSnapshot(state: State): Map<string, Sparkline> {
  return new Map(state.sparkline.sparklines);
}

export function applySparklineSnapshot(
  store: SpreadsheetStore,
  snap: Map<string, Sparkline>,
): void {
  store.setState((s) => ({ ...s, sparkline: { sparklines: new Map(snap) } }));
}

export function recordSparklineChange(
  history: History | null,
  store: SpreadsheetStore,
  mutate: () => void,
): void {
  recordSliceChange(history, store, mutate, {
    capture: captureSparklineSnapshot,
    apply: applySparklineSnapshot,
    same: sameMapSnapshot,
  });
}

export function capturePageSetupSnapshot(state: State): Map<number, PageSetup> {
  const out = new Map<number, PageSetup>();
  for (const [k, v] of state.pageSetup.setupBySheet) {
    out.set(k, {
      ...v,
      margins: { ...v.margins },
      printableBounds: v.printableBounds ? { ...v.printableBounds } : undefined,
      manualPageBreakRows: v.manualPageBreakRows ? [...v.manualPageBreakRows] : undefined,
      manualPageBreakCols: v.manualPageBreakCols ? [...v.manualPageBreakCols] : undefined,
    });
  }
  return out;
}

export function applyPageSetupSnapshot(
  store: SpreadsheetStore,
  snap: Map<number, PageSetup>,
): void {
  const next = new Map<number, PageSetup>();
  for (const [k, v] of snap) {
    next.set(k, {
      ...v,
      margins: { ...v.margins },
      printableBounds: v.printableBounds ? { ...v.printableBounds } : undefined,
      manualPageBreakRows: v.manualPageBreakRows ? [...v.manualPageBreakRows] : undefined,
      manualPageBreakCols: v.manualPageBreakCols ? [...v.manualPageBreakCols] : undefined,
    });
  }
  store.setState((s) => ({ ...s, pageSetup: { setupBySheet: next } }));
}

const sameOptionalArray = (a: readonly number[] | undefined, b: readonly number[] | undefined) => {
  if (!a && !b) return true;
  if (!a || !b || a.length !== b.length) return false;
  return a.every((value, index) => value === b[index]);
};

const sameOptionalMargins = (a: PageMargins | undefined, b: PageMargins | undefined): boolean => {
  if (!a && !b) return true;
  if (!a || !b) return false;
  return a.top === b.top && a.right === b.right && a.bottom === b.bottom && a.left === b.left;
};

const samePageSetup = (a: PageSetup, b: PageSetup): boolean =>
  a.orientation === b.orientation &&
  a.paperSize === b.paperSize &&
  a.paperSizeCode === b.paperSizeCode &&
  a.margins.top === b.margins.top &&
  a.margins.right === b.margins.right &&
  a.margins.bottom === b.margins.bottom &&
  a.margins.left === b.margins.left &&
  sameOptionalMargins(a.printableBounds, b.printableBounds) &&
  a.headerMargin === b.headerMargin &&
  a.footerMargin === b.footerMargin &&
  a.centerHorizontally === b.centerHorizontally &&
  a.centerVertically === b.centerVertically &&
  a.headerLeft === b.headerLeft &&
  a.headerCenter === b.headerCenter &&
  a.headerRight === b.headerRight &&
  a.footerLeft === b.footerLeft &&
  a.footerCenter === b.footerCenter &&
  a.footerRight === b.footerRight &&
  a.firstHeaderLeft === b.firstHeaderLeft &&
  a.firstHeaderCenter === b.firstHeaderCenter &&
  a.firstHeaderRight === b.firstHeaderRight &&
  a.firstFooterLeft === b.firstFooterLeft &&
  a.firstFooterCenter === b.firstFooterCenter &&
  a.firstFooterRight === b.firstFooterRight &&
  a.evenHeaderLeft === b.evenHeaderLeft &&
  a.evenHeaderCenter === b.evenHeaderCenter &&
  a.evenHeaderRight === b.evenHeaderRight &&
  a.evenFooterLeft === b.evenFooterLeft &&
  a.evenFooterCenter === b.evenFooterCenter &&
  a.evenFooterRight === b.evenFooterRight &&
  a.differentOddEvenPages === b.differentOddEvenPages &&
  a.differentFirstPage === b.differentFirstPage &&
  a.scaleHeaderFooterWithDocument === b.scaleHeaderFooterWithDocument &&
  a.alignHeaderFooterWithMargins === b.alignHeaderFooterWithMargins &&
  a.printArea === b.printArea &&
  a.printTitleRows === b.printTitleRows &&
  a.printTitleCols === b.printTitleCols &&
  a.fitWidth === b.fitWidth &&
  a.fitHeight === b.fitHeight &&
  sameOptionalArray(a.manualPageBreakRows, b.manualPageBreakRows) &&
  sameOptionalArray(a.manualPageBreakCols, b.manualPageBreakCols) &&
  a.scale === b.scale &&
  a.printQuality === b.printQuality &&
  a.firstPageNumber === b.firstPageNumber &&
  a.showGridlines === b.showGridlines &&
  a.showHeadings === b.showHeadings &&
  a.blackAndWhite === b.blackAndWhite &&
  a.draftQuality === b.draftQuality &&
  a.comments === b.comments &&
  a.cellErrorsAs === b.cellErrorsAs &&
  a.pageOrder === b.pageOrder;

const samePageSetupSnapshot = (a: Map<number, PageSetup>, b: Map<number, PageSetup>): boolean => {
  if (a.size !== b.size) return false;
  for (const [sheet, setup] of a) {
    const other = b.get(sheet);
    if (!other || !samePageSetup(setup, other)) return false;
  }
  return true;
};

/** Run `mutate`, capturing the page-setup slice before and after, pushing one
 *  entry. No-op when `history` is null. */
export function recordPageSetupChange(
  history: History | null,
  store: SpreadsheetStore,
  mutate: () => void,
): void {
  recordSliceChange(history, store, mutate, {
    capture: capturePageSetupSnapshot,
    apply: applyPageSetupSnapshot,
    same: samePageSetupSnapshot,
  });
}

/** Capture a deep-cloned slicer list for undo replay. Each spec is freshly
 *  cloned so future mutators that mutate the `selected` array can't
 *  retroactively pollute a prior snapshot. */
export function captureSlicersSnapshot(state: State): SlicerSpec[] {
  return state.slicers.slicers.map((sp) => ({ ...sp, selected: [...sp.selected] }));
}

export function applySlicersSnapshot(store: SpreadsheetStore, snap: readonly SlicerSpec[]): void {
  store.setState((s) => ({
    ...s,
    slicers: { slicers: snap.map((sp) => ({ ...sp, selected: [...sp.selected] })) },
  }));
}

/** Run `mutate` and push one history entry capturing the slicer-slice
 *  before/after. Use for any add/remove/update/setSelected call that should
 *  be undoable. */
export function recordSlicersChange(
  history: History | null,
  store: SpreadsheetStore,
  mutate: () => void,
): void {
  recordSliceChange(history, store, mutate, {
    capture: captureSlicersSnapshot,
    apply: applySlicersSnapshot,
    same: sameJsonSnapshot,
  });
}
