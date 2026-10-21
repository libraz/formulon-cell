import type { SpreadsheetStore, State } from '../store/store.js';
import type { SheetViewMemo } from '../store/types.js';
import { addrKey, MAX_COL, MAX_ROW, parseAddrKey } from './address.js';
import { type AxisStructureShift, remapAxisSnapshot } from './axis-structure-remap.js';
import { fontCssPxToPoints } from './font-units.js';
import type {
  CellXf,
  Range,
  SheetFontAxisSnapshot,
  SheetFontOverride,
  SheetFontPatch,
} from './types.js';
import type { WorkbookHandle } from './workbook-handle.js';

type AxisStyle = { readonly first: number; readonly last: number; readonly styleXf: number };

interface ActiveWholeSheetFont {
  readonly baseline: SheetFontAxisSnapshot;
}

export interface ActiveWholeSheetFontSnapshot {
  readonly baseline: SheetFontAxisSnapshot;
}

/** Baselines live with the workbook rather than in a cell-sized map. They are
 * needed when history restores a state that no longer contains the override. */
const activeOverrides = new WeakMap<WorkbookHandle, Map<number, ActiveWholeSheetFont>>();

const cloneAxisSnapshot = (snapshot: SheetFontAxisSnapshot): SheetFontAxisSnapshot => ({
  columns: snapshot.columns.map((entry) => ({ ...entry })),
  rows: snapshot.rows.map((entry) => ({ ...entry })),
});

interface StructureEventBinding {
  readonly stores: Set<SpreadsheetStore>;
  readonly beforeStoreRemap: Map<SpreadsheetStore, (mapIndex: SheetIndexMapper) => void>;
  unsubscribe: () => void;
}

const structureEventBindings = new WeakMap<WorkbookHandle, StructureEventBinding>();

export type SheetIndexMapper = (index: number) => number | null;

const mapIndexForStructureEvent = (
  event: { kind: 'sheet-remove'; index: number } | { kind: 'sheet-move'; from: number; to: number },
): SheetIndexMapper => {
  if (event.kind === 'sheet-remove') {
    return (index) => {
      if (index === event.index) return null;
      return index > event.index ? index - 1 : index;
    };
  }
  return (index) => {
    if (index === event.from) return event.to;
    if (event.from < event.to && index > event.from && index <= event.to) return index - 1;
    if (event.from > event.to && index >= event.to && index < event.from) return index + 1;
    return index;
  };
};

export const sheetIndexMapperForStructureEvent = mapIndexForStructureEvent;

interface RemappedAddressMap<T> {
  readonly map: Map<string, T>;
  readonly changed: boolean;
}

const remapAddressMap = <T>(
  current: ReadonlyMap<string, T>,
  mapIndex: SheetIndexMapper,
): RemappedAddressMap<T> => {
  const next = new Map<string, T>();
  let changed = false;
  for (const [key, value] of current) {
    const addr = parseAddrKey(key);
    if (!addr) {
      next.set(key, value);
      continue;
    }
    const sheet = mapIndex(addr.sheet);
    if (sheet === null) {
      changed = true;
      continue;
    }
    const nextKey = addrKey({ ...addr, sheet });
    if (nextKey !== key) changed = true;
    next.set(nextKey, value);
  }
  return { map: next, changed };
};

/** Each sheet keeps its remembered selection and scroll across a move or remove. */
const remapViewMemos = (
  memos: ReadonlyMap<number, SheetViewMemo> | undefined,
  mapIndex: SheetIndexMapper,
): ReadonlyMap<number, SheetViewMemo> | undefined => {
  if (!memos) return memos;
  let changed = false;
  const next = new Map<number, SheetViewMemo>();
  for (const [index, memo] of memos) {
    const mapped = mapIndex(index);
    if (mapped !== index) changed = true;
    if (mapped === null) continue;
    const sel = memo.selection;
    next.set(
      mapped,
      mapped === index
        ? memo
        : {
            ...memo,
            selection: {
              active: { ...sel.active, sheet: mapped },
              anchor: { ...sel.anchor, sheet: mapped },
              range: { ...sel.range, sheet: mapped },
              ...(sel.extraRanges && {
                extraRanges: sel.extraRanges.map((r) => ({ ...r, sheet: mapped })),
              }),
            },
          },
    );
  }
  return changed ? next : memos;
};

const remapRangeValidations = (
  rules: State['format']['rangeValidations'],
  mapIndex: SheetIndexMapper,
): State['format']['rangeValidations'] => {
  if (!rules || rules.length === 0) return rules;
  let changed = false;
  const next = rules.flatMap((rule) => {
    const ranges = rule.ranges.flatMap((range) => {
      const mapped = mapIndex(range.sheet);
      if (mapped === null) return [];
      if (mapped === range.sheet) return [range];
      changed = true;
      return [{ ...range, sheet: mapped }];
    });
    if (ranges.length !== rule.ranges.length) changed = true;
    return ranges.length > 0 ? [{ ...rule, ranges }] : [];
  });
  return changed ? next : rules;
};

/** Rebuild both format maps from the same pre-event state. The native
 * workbook has already moved the sheets at this point, so publishing the old
 * keys even once would flush their XF records onto the wrong worksheet. */
export const remapStoreSheetStateForStructureEvent = (
  store: SpreadsheetStore,
  mapIndex: SheetIndexMapper,
): void => {
  store.setState((state) => {
    const formats = remapAddressMap(state.format.formats, mapIndex);
    const currentOverrides = state.format.sheetFontOverrides;
    const overrides = currentOverrides ? new Map<number, SheetFontOverride>() : undefined;
    let overridesChanged = false;
    if (currentOverrides && overrides) {
      for (const [index, override] of currentOverrides) {
        const mapped = mapIndex(index);
        if (mapped === null) {
          overridesChanged = true;
          continue;
        }
        if (mapped !== index) overridesChanged = true;
        overrides.set(mapped, override);
      }
    }
    const memos = remapViewMemos(state.ui.sheetViewMemos, mapIndex);
    const rangeValidations = remapRangeValidations(state.format.rangeValidations, mapIndex);
    if (
      !formats.changed &&
      !overridesChanged &&
      memos === state.ui.sheetViewMemos &&
      rangeValidations === state.format.rangeValidations
    ) {
      return state;
    }
    return {
      ...state,
      ui: memos === state.ui.sheetViewMemos ? state.ui : { ...state.ui, sheetViewMemos: memos },
      format: {
        ...state.format,
        formats: formats.changed ? formats.map : state.format.formats,
        sheetFontOverrides: overridesChanged ? overrides : currentOverrides,
        rangeValidations,
      },
    };
  });
};

const ensureStructureEventBinding = (wb: WorkbookHandle): StructureEventBinding | null => {
  const current = structureEventBindings.get(wb);
  if (current) return current;
  if (typeof wb.subscribe !== 'function') return null;
  const stores = new Set<SpreadsheetStore>();
  const beforeStoreRemap = new Map<SpreadsheetStore, (mapIndex: SheetIndexMapper) => void>();
  const binding: StructureEventBinding = {
    stores,
    beforeStoreRemap,
    unsubscribe: () => undefined,
  };
  binding.unsubscribe = wb.subscribe((event) => {
    if (event.kind !== 'sheet-remove' && event.kind !== 'sheet-move') return;
    const mapIndex = mapIndexForStructureEvent(event);
    const remap = (): void => {
      remapActiveWholeSheetFontOverrides(wb, mapIndex);
      for (const store of stores) {
        beforeStoreRemap.get(store)?.(mapIndex);
        remapStoreSheetStateForStructureEvent(store, mapIndex);
      }
    };
    if (typeof wb.withEngineSyncMuted === 'function') wb.withEngineSyncMuted(remap);
    else remap();
  });
  structureEventBindings.set(wb, binding);
  return binding;
};

const activeFor = (wb: WorkbookHandle): Map<number, ActiveWholeSheetFont> => {
  let map = activeOverrides.get(wb);
  if (!map) {
    map = new Map();
    activeOverrides.set(wb, map);
  }
  ensureStructureEventBinding(wb);
  return map;
};

/** Register the store associated with a handle for the same structure event
 *  stream that remaps the workbook-local baseline cache. The event callback
 *  updates both maps before StoreMirror's ordinary store subscriber flushes
 *  native formats, so a move/remove cannot briefly target the old index. */
export const bindWholeSheetFontStore = (
  wb: WorkbookHandle,
  store: SpreadsheetStore,
  beforeStoreRemap?: (mapIndex: SheetIndexMapper) => void,
): (() => void) => {
  const binding = ensureStructureEventBinding(wb);
  if (!binding) return () => undefined;
  binding.stores.add(store);
  if (beforeStoreRemap) binding.beforeStoreRemap.set(store, beforeStoreRemap);
  return () => {
    binding.stores.delete(store);
    binding.beforeStoreRemap.delete(store);
    const active = activeOverrides.get(wb);
    if (binding.stores.size === 0 && (!active || active.size === 0)) {
      binding.unsubscribe();
      structureEventBindings.delete(wb);
    }
  };
};

export const hasWholeSheetFontStoreBinding = (
  wb: WorkbookHandle,
  store: SpreadsheetStore,
): boolean => structureEventBindings.get(wb)?.stores.has(store) ?? false;

/**
 * Sheet style state is addressed by the workbook's current numeric sheet
 * index. Structure commands move and remove those indices after the native
 * workbook has already applied the operation, so keep the workbook-local
 * baseline cache in lockstep with the store's remap.
 */
export const remapActiveWholeSheetFontOverrides = (
  wb: WorkbookHandle,
  mapIndex: (index: number) => number | null,
): void => {
  const active = activeOverrides.get(wb);
  if (!active) return;
  const remapped = new Map<number, ActiveWholeSheetFont>();
  for (const [index, value] of active) {
    const next = mapIndex(index);
    if (next !== null) remapped.set(next, value);
  }
  activeOverrides.set(wb, remapped);
};

/** Row/column structure edits move the native dimension records that a legacy
 * whole-sheet font override captured. Keep its workbook-local immutable
 * baseline in the same coordinate space as the store override. */
export const remapActiveWholeSheetFontOverridesForStructure = (
  wb: WorkbookHandle,
  sheet: number,
  shift: AxisStructureShift,
): void => {
  const active = activeOverrides.get(wb);
  const current = active?.get(sheet);
  if (!active || !current) return;
  active.set(sheet, { baseline: remapAxisSnapshot(current.baseline, shift) });
};

export const captureActiveWholeSheetFontOverrideForStructure = (
  wb: WorkbookHandle,
  sheet: number,
): ActiveWholeSheetFontSnapshot | null => {
  const value = activeOverrides.get(wb)?.get(sheet);
  return value ? { baseline: cloneAxisSnapshot(value.baseline) } : null;
};

export const restoreActiveWholeSheetFontOverrideForStructure = (
  wb: WorkbookHandle,
  sheet: number,
  snapshot: ActiveWholeSheetFontSnapshot | null,
): void => {
  const active = activeOverrides.get(wb);
  if (!active) {
    if (snapshot === null) return;
    activeOverrides.set(wb, new Map([[sheet, { baseline: cloneAxisSnapshot(snapshot.baseline) }]]));
    return;
  }
  if (snapshot === null) active.delete(sheet);
  else active.set(sheet, { baseline: cloneAxisSnapshot(snapshot.baseline) });
  if (active.size === 0) activeOverrides.delete(wb);
};

const sameSheet = (ranges: readonly Range[]): boolean => {
  const sheet = ranges[0]?.sheet;
  return sheet !== undefined && ranges.every((range) => range.sheet === sheet);
};

/** Return true only when the union of the selected rectangles is exactly the
 *  complete worksheet. This accepts a tiled primary/extra selection while
 *  avoiding materialisation of 17 billion cells. */
export const isWholeSheetSelection = (ranges: readonly Range[]): boolean => {
  if (ranges.length === 0 || !sameSheet(ranges)) return false;
  const rowBoundaries = new Set<number>([0, MAX_ROW + 1]);
  for (const range of ranges) {
    if (
      range.r0 < 0 ||
      range.c0 < 0 ||
      range.r1 > MAX_ROW ||
      range.c1 > MAX_COL ||
      range.r0 > range.r1 ||
      range.c0 > range.c1
    )
      return false;
    rowBoundaries.add(range.r0);
    rowBoundaries.add(range.r1 + 1);
  }
  const rows = [...rowBoundaries].sort((a, b) => a - b);
  for (let index = 0; index + 1 < rows.length; index += 1) {
    const start = rows[index] as number;
    const endExclusive = rows[index + 1] as number;
    if (start >= endExclusive) continue;
    const intervals = ranges
      .filter((range) => range.r0 <= start && range.r1 + 1 >= endExclusive)
      .map((range) => ({ first: range.c0, last: range.c1 }))
      .sort((a, b) => a.first - b.first || a.last - b.last);
    let cursor = 0;
    for (const interval of intervals) {
      if (interval.first > cursor) return false;
      cursor = Math.max(cursor, interval.last + 1);
      if (cursor > MAX_COL) break;
    }
    if (cursor <= MAX_COL) return false;
  }
  return true;
};

export const sheetFontPatchFor = (state: State, sheet: number): SheetFontPatch | undefined =>
  state.format.sheetFontOverrides?.get(sheet)?.patch;

const fontPatchedXf = (
  wb: WorkbookHandle,
  baseXfIndex: number,
  patch: SheetFontPatch,
  cache: Map<string, number>,
): number => {
  const cacheKey = `${baseXfIndex}|${patch.fontFamily ?? ''}|${patch.fontSize ?? ''}`;
  const cached = cache.get(cacheKey);
  if (cached !== undefined) return cached;
  const xf = wb.getCellXf(baseXfIndex);
  if (!xf) return -1;
  const baseFont = wb.getFontRecord(xf.fontIndex);
  if (!baseFont) return -1;
  const font = {
    ...baseFont,
    ...(patch.fontFamily === undefined ? {} : { name: patch.fontFamily }),
    ...(patch.fontSize === undefined ? {} : { size: fontCssPxToPoints(patch.fontSize) }),
  };
  // A selected family is a literal font choice. Keeping a theme scheme from
  // the source XF would make the engine resolve the user's family back to a
  // theme font on reload. Preserve the scheme only when the selected family
  // is the workbook's own default family.
  if (patch.fontFamily !== undefined) {
    font.scheme =
      patch.fontFamily === wb.workbookDefaultFont?.name
        ? wb.workbookDefaultFont?.scheme
        : undefined;
  }
  if (!font.name || !Number.isFinite(font.size) || font.size <= 0) return -1;
  const fontIndex = wb.addFontRecord(font);
  if (fontIndex < 0) return -1;
  const xfIndex = wb.addXfRecord({ ...xf, fontIndex });
  if (xfIndex < 0) return -1;
  cache.set(cacheKey, xfIndex);
  return xfIndex;
};

const columnSegments = (baseline: SheetFontAxisSnapshot): AxisStyle[] => {
  const entries = [...baseline.columns]
    .filter(
      (entry) =>
        Number.isInteger(entry.first) &&
        Number.isInteger(entry.last) &&
        entry.first <= entry.last &&
        entry.last >= 0 &&
        entry.first <= MAX_COL,
    )
    .map((entry) => ({
      first: Math.max(0, entry.first),
      last: Math.min(MAX_COL, entry.last),
      styleXf: entry.styleXf,
    }))
    .sort((a, b) => a.first - b.first || a.last - b.last);
  const segments: AxisStyle[] = [];
  let cursor = 0;
  const push = (first: number, last: number, styleXf: number): void => {
    if (first > last) return;
    const previous = segments.at(-1);
    if (previous && previous.last + 1 === first && previous.styleXf === styleXf) {
      segments[segments.length - 1] = { ...previous, last };
    } else segments.push({ first, last, styleXf });
  };
  for (const entry of entries) {
    if (entry.first > cursor) push(cursor, entry.first - 1, 0);
    const first = Math.max(cursor, entry.first);
    if (first <= entry.last) push(first, entry.last, entry.styleXf);
    cursor = Math.max(cursor, entry.last + 1);
    if (cursor > MAX_COL) break;
  }
  if (cursor <= MAX_COL) push(cursor, MAX_COL, 0);
  if (segments.length === 0) push(0, MAX_COL, 0);
  return segments;
};

const restoreBaseline = (
  wb: WorkbookHandle,
  sheet: number,
  baseline: SheetFontAxisSnapshot,
): boolean => {
  if (!wb.clearColumnStyle(sheet, 0, MAX_COL)) return false;
  for (const entry of baseline.columns) {
    if (!wb.setColumnStyle(sheet, entry.first, entry.last, entry.styleXf)) return false;
  }
  for (const entry of baseline.rows) {
    if (!wb.clearRowStyle(sheet, entry.row)) return false;
    if (!wb.setRowStyle(sheet, entry.row, entry.styleXf)) return false;
  }
  return true;
};

const clearedVisualXfForSource = (
  wb: WorkbookHandle,
  sourceXfIndex: number,
  cache: Map<number, number>,
): number => {
  const cached = cache.get(sourceXfIndex);
  if (cached !== undefined) return cached;
  const source = wb.getCellXf(sourceXfIndex);
  const defaultXf = wb.getCellXf(0);
  if (!source || !defaultXf) return -1;
  const cleared: CellXf = { ...defaultXf };
  const target = cleared as unknown as Record<string, unknown>;
  const raw = source as unknown as Record<string, unknown>;
  for (const field of [
    'applyProtection',
    'hasProtection',
    'locked',
    'hidden',
    'quotePrefix',
  ] as const) {
    if (Object.hasOwn(raw, field)) target[field] = raw[field];
    else delete target[field];
  }
  const xfIndex = wb.addXfRecord(cleared);
  if (xfIndex >= 0) cache.set(sourceXfIndex, xfIndex);
  return xfIndex;
};

const applyClearedVisualAxes = (
  wb: WorkbookHandle,
  sheet: number,
  baseline: SheetFontAxisSnapshot,
  patch: SheetFontPatch,
): boolean => {
  // Clearing the complete column span and every row that had an explicit
  // style removes only the axis style record. The native API keeps width,
  // height, hidden and outline metadata on those dimension records.
  if (!wb.clearColumnStyle(sheet, 0, MAX_COL)) return false;
  for (const entry of baseline.rows) {
    if (!wb.clearRowStyle(sheet, entry.row)) return false;
  }
  const sourceCache = new Map<number, number>();
  const fontCache = new Map<string, number>();
  for (const segment of columnSegments(baseline)) {
    const visualXf = clearedVisualXfForSource(wb, segment.styleXf, sourceCache);
    if (visualXf < 0) return false;
    const xfIndex = Object.keys(patch).length
      ? fontPatchedXf(wb, visualXf, patch, fontCache)
      : visualXf;
    if (xfIndex < 0) return false;
    if (xfIndex !== 0 && !wb.setColumnStyle(sheet, segment.first, segment.last, xfIndex)) {
      return false;
    }
  }
  for (const entry of baseline.rows) {
    const visualXf = clearedVisualXfForSource(wb, entry.styleXf, sourceCache);
    if (visualXf < 0) return false;
    const xfIndex = Object.keys(patch).length
      ? fontPatchedXf(wb, visualXf, patch, fontCache)
      : visualXf;
    if (xfIndex < 0) return false;
    if (xfIndex !== 0 && !wb.setRowStyle(sheet, entry.row, xfIndex)) return false;
  }
  return true;
};

const applyBaselineWithFont = (
  wb: WorkbookHandle,
  sheet: number,
  baseline: SheetFontAxisSnapshot,
  patch: SheetFontPatch,
): boolean => {
  const cache = new Map<string, number>();
  for (const segment of columnSegments(baseline)) {
    const xfIndex = fontPatchedXf(wb, segment.styleXf, patch, cache);
    if (xfIndex < 0 || !wb.setColumnStyle(sheet, segment.first, segment.last, xfIndex))
      return false;
  }
  for (const entry of baseline.rows) {
    const xfIndex = fontPatchedXf(wb, entry.styleXf, patch, cache);
    if (xfIndex < 0 || !wb.setRowStyle(sheet, entry.row, xfIndex)) return false;
  }
  return true;
};

/**
 * Synchronise worksheet-level font overrides and return whether the native
 * axis writes succeeded. The caller should treat `false` as an unsupported or
 * rejected mutation; no per-cell fallback is attempted because that would
 * materialise the whole worksheet.
 */
export function syncWholeSheetFontOverride(
  wb: WorkbookHandle,
  store: SpreadsheetStore,
  sheet: number,
): boolean {
  const override = store.getState().format.sheetFontOverrides?.get(sheet);
  const active = activeFor(wb);
  const previous = active.get(sheet);
  if (!override && !previous) return true;
  if (wb.capabilities.sheetFontStyles !== true) return false;
  if (override) {
    const baseline = previous?.baseline ?? override.baseline ?? wb.getSheetFontAxisSnapshot(sheet);
    if (!baseline) return false;
    active.set(sheet, { baseline });
    // Keep the immutable pre-edit axis records in the format slice as well as
    // the workbook-local cache. History captures this state after the store
    // mutation, and a later workbook reload can therefore restore the source
    // column/row styles without relying on the WeakMap.
    if (!override.baseline) {
      store.setState((current) => {
        const currentOverride = current.format.sheetFontOverrides?.get(sheet);
        if (!currentOverride || currentOverride.baseline) return current;
        const overrides = new Map(current.format.sheetFontOverrides);
        overrides.set(sheet, { ...currentOverride, baseline });
        return { ...current, format: { ...current.format, sheetFontOverrides: overrides } };
      });
    }
    if (override.clearAxes) return applyClearedVisualAxes(wb, sheet, baseline, override.patch);
    return applyBaselineWithFont(wb, sheet, baseline, override.patch);
  }
  if (!previous) return true;
  const restored = restoreBaseline(wb, sheet, previous.baseline);
  if (restored) active.delete(sheet);
  return restored;
}

/** Sync every override currently present in the store and restore axis
 * baselines for overrides removed by history. Clear Formats keeps its
 * explicit visual-base marker in the map. */
export function syncWholeSheetFontOverrides(
  wb: WorkbookHandle,
  store: SpreadsheetStore,
  sheets?: Iterable<number>,
): boolean {
  const targets = new Set<number>(
    sheets ?? store.getState().format.sheetFontOverrides?.keys() ?? [],
  );
  const active = activeFor(wb);
  for (const sheet of active.keys()) targets.add(sheet);
  let ok = true;
  for (const sheet of targets) ok = syncWholeSheetFontOverride(wb, store, sheet) && ok;
  return ok;
}

/** Apply the active worksheet override to a direct format only when the
 *  format omitted that font axis. This keeps a later number/fill edit from
 *  creating an XF0 that masks the column style on a blank/future cell. */
export const formatWithSheetFontPatch = <T extends object>(
  format: T,
  patch: SheetFontPatch | undefined,
): T => {
  if (!patch) return format;
  const next = { ...format } as Record<string, unknown>;
  const source = format as Record<string, unknown>;
  if (patch.fontFamily !== undefined && source.fontFamily === undefined) {
    next.fontFamily = patch.fontFamily;
  }
  if (patch.fontSize !== undefined && source.fontSize === undefined) {
    next.fontSize = patch.fontSize;
  }
  return next as T;
};

/** Exposed for tests and engine sync diagnostics. */
export const activeSheetFontBaseline = (
  wb: WorkbookHandle,
  sheet: number,
): SheetFontAxisSnapshot | undefined => activeFor(wb).get(sheet)?.baseline;
