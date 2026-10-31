import type { CellFormat, SpreadsheetStore, State } from '../store/store.js';
import { addrKey, MAX_COL, MAX_ROW, parseAddrKey } from './address.js';
import {
  type AxisStructureShift,
  remapAxisFormatOverride,
  remapAxisSnapshot,
} from './axis-structure-remap.js';
import { patchXfIndexFromBase } from './cell-format-sync.js';
import { formatHistoryBridgeFor } from './format-history-bridge.js';
import type {
  Addr,
  Range,
  SheetAxisDirectXf,
  SheetAxisFormatOverride,
  SheetFontAxisSnapshot,
  SheetFontPatch,
} from './types.js';
import type { WorkbookHandle } from './workbook-handle.js';

export interface AxisFormatBand {
  readonly axis: 'row' | 'column';
  readonly first: number;
  readonly last: number;
}

/** Resolve a patch from the source XF being transformed. Toolbar commands
 * such as decimal/indent bumps use this to preserve differing axis styles
 * while applying one action to each source. Returning an empty object keeps
 * that source unchanged. */
export type AxisFormatPatchResolver = (
  wb: WorkbookHandle,
  baseXfIndex: number,
  addr?: Addr,
) => Partial<CellFormat> | null;

export interface AxisFormatPreparation {
  readonly override: SheetAxisFormatOverride;
  readonly sheet: number;
  readonly directCells: readonly Addr[];
}

const emptySnapshot = (): SheetFontAxisSnapshot => ({ columns: [], rows: [] });

const cloneSnapshot = (snapshot: SheetFontAxisSnapshot): SheetFontAxisSnapshot => ({
  columns: snapshot.columns.map((entry) => ({ ...entry })),
  rows: snapshot.rows.map((entry) => ({ ...entry })),
});

const cloneDirect = (entry: SheetAxisDirectXf): SheetAxisDirectXf => ({
  key: entry.key,
  before: { ...entry.before },
  afterXfIndex: entry.afterXfIndex,
});

const validBand = (band: AxisFormatBand): boolean =>
  Number.isInteger(band.first) &&
  Number.isInteger(band.last) &&
  band.first >= 0 &&
  band.first <= band.last &&
  (band.axis === 'row' ? band.last <= MAX_ROW : band.last <= MAX_COL);

const normalizedBands = (bands: readonly AxisFormatBand[]): AxisFormatBand[] => {
  const sorted = bands.filter(validBand).map((band) => ({ ...band }));
  sorted.sort((a, b) => a.axis.localeCompare(b.axis) || a.first - b.first || a.last - b.last);
  const merged: AxisFormatBand[] = [];
  for (const band of sorted) {
    const previous = merged.at(-1);
    if (previous && previous.axis === band.axis && previous.last + 1 >= band.first) {
      merged[merged.length - 1] = { ...previous, last: Math.max(previous.last, band.last) };
    } else merged.push(band);
  }
  return merged;
};

const intersectsBand = (
  bands: readonly AxisFormatBand[],
  axis: AxisFormatBand['axis'],
  n: number,
): boolean => bands.some((band) => band.axis === axis && n >= band.first && n <= band.last);

const columnSegments = (
  snapshot: SheetFontAxisSnapshot,
): Array<{
  first: number;
  last: number;
  styleXf: number;
  hasStyle: boolean;
}> => {
  const entries = snapshot.columns
    .filter(
      (entry) =>
        Number.isInteger(entry.first) &&
        Number.isInteger(entry.last) &&
        Number.isInteger(entry.styleXf) &&
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
  const result: Array<{ first: number; last: number; styleXf: number; hasStyle: boolean }> = [];
  let cursor = 0;
  const push = (first: number, last: number, styleXf: number, hasStyle: boolean): void => {
    if (first > last) return;
    const previous = result.at(-1);
    if (
      previous &&
      previous.last + 1 === first &&
      previous.styleXf === styleXf &&
      previous.hasStyle === hasStyle
    ) {
      result[result.length - 1] = { ...previous, last };
    } else result.push({ first, last, styleXf, hasStyle });
  };
  for (const entry of entries) {
    if (entry.first > cursor) push(cursor, entry.first - 1, 0, false);
    const first = Math.max(cursor, entry.first);
    if (first <= entry.last) push(first, entry.last, entry.styleXf, true);
    cursor = Math.max(cursor, entry.last + 1);
    if (cursor > MAX_COL) break;
  }
  if (cursor <= MAX_COL) push(cursor, MAX_COL, 0, false);
  if (result.length === 0) push(0, MAX_COL, 0, false);
  return result;
};

const splitColumns = (
  snapshot: SheetFontAxisSnapshot,
  bands: readonly AxisFormatBand[],
  patch: Partial<CellFormat>,
  wb: WorkbookHandle | null,
  resolvePatch?: AxisFormatPatchResolver,
): readonly SheetFontAxisSnapshot['columns'][number][] | null => {
  const out: Array<{ first: number; last: number; styleXf: number }> = [];
  const cache = new Map<string, number>();
  for (const segment of columnSegments(snapshot)) {
    let cursor = segment.first;
    while (cursor <= segment.last) {
      const selected = intersectsBand(bands, 'column', cursor);
      let end = cursor;
      while (end < segment.last && intersectsBand(bands, 'column', end + 1) === selected) end += 1;
      let styleXf = segment.styleXf;
      if (selected && wb) {
        const resolvedPatch = resolvePatch?.(wb, styleXf) ?? patch;
        const cacheKey = `${styleXf}|${JSON.stringify(resolvedPatch)}`;
        styleXf = cache.get(cacheKey) ?? patchXfIndexFromBase(wb, styleXf, resolvedPatch);
        if (styleXf < 0) return null;
        cache.set(cacheKey, styleXf);
      }
      // An implicit default gap is a computation base, not an authored
      // column record. Preserve that distinction from explicit Normal XF0.
      if (selected || segment.hasStyle) {
        const previous = out.at(-1);
        if (previous && previous.last + 1 === cursor && previous.styleXf === styleXf) {
          out[out.length - 1] = { ...previous, last: end };
        } else out.push({ first: cursor, last: end, styleXf });
      }
      cursor = end + 1;
    }
  }
  return out;
};

const rowsForSnapshot = (snapshot: SheetFontAxisSnapshot): Map<number, number> =>
  new Map(
    snapshot.rows
      .filter(
        (entry) =>
          Number.isInteger(entry.row) &&
          Number.isInteger(entry.styleXf) &&
          entry.row >= 0 &&
          entry.row <= MAX_ROW,
      )
      .map((entry) => [entry.row, entry.styleXf]),
  );

interface SplitRowsResult {
  readonly rows: readonly SheetFontAxisSnapshot['rows'][number][];
  readonly baseByRow: ReadonlyMap<number, number>;
}

const splitRows = (
  snapshot: SheetFontAxisSnapshot,
  bands: readonly AxisFormatBand[],
  patch: Partial<CellFormat>,
  wb: WorkbookHandle | null,
  resolvePatch?: AxisFormatPatchResolver,
): SplitRowsResult | null => {
  const rows = rowsForSnapshot(snapshot);
  const selectedRows = new Set<number>();
  let selectedCount = 0;
  for (const band of bands) {
    if (band.axis !== 'row') continue;
    selectedCount += band.last - band.first + 1;
    if (selectedCount > 100_000) return null;
    for (let row = band.first; row <= band.last; row += 1) selectedRows.add(row);
  }
  const cache = new Map<string, number>();
  const baseByRow = new Map<number, number>();
  const segments = columnSegments(snapshot);
  const widthByStyle = new Map<number, number>();
  for (const segment of segments) {
    const width = segment.last - segment.first + 1;
    widthByStyle.set(segment.styleXf, (widthByStyle.get(segment.styleXf) ?? 0) + width);
  }
  const commonColumnBase =
    [...widthByStyle.entries()].sort((left, right) => right[1] - left[1])[0]?.[0] ?? 0;
  for (const row of selectedRows) {
    const base = rows.get(row) ?? commonColumnBase;
    baseByRow.set(row, base);
    let styleXf = base;
    if (wb) {
      const resolvedPatch = resolvePatch?.(wb, base) ?? patch;
      const cacheKey = `${base}|${JSON.stringify(resolvedPatch)}`;
      styleXf = cache.get(cacheKey) ?? patchXfIndexFromBase(wb, base, resolvedPatch);
      if (styleXf < 0) return null;
      cache.set(cacheKey, styleXf);
    }
    rows.set(row, styleXf);
  }
  return {
    rows: [...rows.entries()].sort(([a], [b]) => a - b).map(([row, styleXf]) => ({ row, styleXf })),
    baseByRow,
  };
};

const selectedByBand = (bands: readonly AxisFormatBand[], addr: Addr): boolean =>
  bands.some(
    (band) =>
      (band.axis === 'row' && addr.row >= band.first && addr.row <= band.last) ||
      (band.axis === 'column' && addr.col >= band.first && addr.col <= band.last),
  );

type WorkbookPhysicalCell = { readonly addr: Addr };

/** Return authored native cells without expanding blank worksheet space. The
 * optional adapter method includes blank cells that carry only a direct XF;
 * older handles fall back to their populated-cell iterator. */
const nativeDirectCellsForSelection = (
  wb: WorkbookHandle,
  sheet: number,
  bands: readonly AxisFormatBand[],
): readonly Addr[] => {
  const extended = wb as WorkbookHandle & {
    physicalCellsWithXf?: (sheet: number) => Iterable<WorkbookPhysicalCell>;
  };
  const entries =
    typeof extended.physicalCellsWithXf === 'function'
      ? extended.physicalCellsWithXf(sheet)
      : wb.physicalCells(sheet);
  const seen = new Set<string>();
  const result: Addr[] = [];
  for (const entry of entries) {
    const addr = entry.addr;
    if (addr.sheet !== sheet || !selectedByBand(bands, addr)) continue;
    // A physical cell is a cell-source XF even when its index is 0. The native
    // effective-style resolver gives such cells precedence over row/column
    // styles, so populated/default-XF cells must be kept in the sparse
    // correction set alongside opaque non-zero XFs such as quotePrefix.
    const key = addrKey(addr);
    if (seen.has(key)) continue;
    seen.add(key);
    result.push(addr);
  }
  return result;
};

/** Count every bounded write that this preparation can require before any
 * native XF is allocated.  Row styles are explicit native writes; compact
 * column spans are cheap and are intentionally excluded.  Direct corrections
 * are deduplicated by address so mixed row/column selections do not consume
 * the budget twice. */
const axisBudgetAvailable = (
  sheet: number,
  snapshot: SheetFontAxisSnapshot,
  existing: SheetAxisFormatOverride | undefined,
  bands: readonly AxisFormatBand[],
  readNativeDirectCells: (() => readonly Addr[]) | null,
  exceptions: readonly Addr[],
  baseUsed: number,
  baseKeys: ReadonlySet<string>,
): boolean => {
  if (!Number.isInteger(baseUsed) || baseUsed < 0 || baseUsed > 100_000) return false;
  if (exceptions.some((addr) => addr.sheet !== sheet)) return false;
  let used = baseUsed;
  const countedKeys = new Set(baseKeys);
  const directKeys = new Set<string>();
  const addCost = (key: string): boolean => {
    if (countedKeys.has(key)) return true;
    countedKeys.add(key);
    used += 1;
    return used <= 100_000;
  };
  const addDirect = (key: string): boolean => {
    directKeys.add(key);
    return addCost(key);
  };
  for (const entry of existing?.directXfs ?? []) if (!addDirect(entry.key)) return false;
  for (const addr of exceptions) {
    if (!addDirect(addrKey(addr))) return false;
  }

  const selectedRows = new Set<number>();
  for (const band of bands) {
    if (band.axis !== 'row') continue;
    const count = band.last - band.first + 1;
    if (count > 100_000 || used + count > 100_000) return false;
    used += count;
    for (let row = band.first; row <= band.last; row += 1) selectedRows.add(row);
  }
  if (!readNativeDirectCells) return true;

  const exceptionKeys = new Set(exceptions.map(addrKey));
  for (const addr of readNativeDirectCells()) {
    if (!exceptionKeys.has(addrKey(addr)) && !addDirect(addrKey(addr))) return false;
  }
  const selectedColumns = bands.filter((band) => band.axis === 'column');
  if (selectedColumns.length > 0) {
    // A styled row wins over a selected column at every column in its band.
    // Count that sparse intersection without calling getCellXfIndex or
    // allocating a patched XF. Existing direct entries are deduplicated.
    for (const row of snapshot.rows) {
      for (const band of selectedColumns) {
        for (let col = band.first; col <= band.last; col += 1) {
          const key = addrKey({ sheet, row: row.row, col });
          if (!exceptionKeys.has(key) && !addDirect(key)) return false;
        }
      }
    }
  }

  const rows = rowsForSnapshot(snapshot);
  const segments = columnSegments(snapshot);
  const widthByStyle = new Map<number, number>();
  for (const segment of segments) {
    const width = segment.last - segment.first + 1;
    widthByStyle.set(segment.styleXf, (widthByStyle.get(segment.styleXf) ?? 0) + width);
  }
  const commonColumnBase =
    [...widthByStyle.entries()].sort((left, right) => right[1] - left[1])[0]?.[0] ?? 0;
  for (const row of selectedRows) {
    if (rows.has(row)) continue;
    for (const segment of segments) {
      if (segment.styleXf === commonColumnBase) continue;
      for (let col = segment.first; col <= segment.last; col += 1) {
        const key = addrKey({ sheet, row, col });
        if (!exceptionKeys.has(key) && !addDirect(key)) return false;
      }
    }
  }
  return used <= 100_000;
};

type EffectiveStyle = NonNullable<ReturnType<WorkbookHandle['getEffectiveStyle']>>;

const directSnapshot = (
  wb: WorkbookHandle,
  addr: Addr,
  prefetchedEffective?: EffectiveStyle,
): SheetAxisDirectXf | null => {
  const directXfIndex = wb.getCellXfIndex(addr.sheet, addr.row, addr.col);
  if (directXfIndex === null) return null;
  const effective = prefetchedEffective ?? wb.getEffectiveStyle(addr.sheet, addr.row, addr.col);
  return {
    key: addrKey(addr),
    before: {
      directXfIndex,
      effectiveXfIndex: effective?.xfIndex ?? null,
      source: effective?.source ?? null,
    },
    afterXfIndex: directXfIndex,
  };
};

/**
 * Prepare one compact axis override. The caller applies the returned map in a
 * single store publication. Native XF records are allocated before that
 * publication while StoreMirror is muted, so the subsequent flush is atomic
 * from the store's point of view.
 */
export function prepareAxisFormatOverride(
  state: State,
  store: SpreadsheetStore,
  ranges: readonly Range[],
  bands: readonly AxisFormatBand[],
  patch: Partial<CellFormat>,
): AxisFormatPreparation | null {
  const sheet = ranges[0]?.sheet;
  if (sheet === undefined || ranges.some((range) => range.sheet !== sheet)) return null;
  const normalized = normalizedBands(bands);
  if (normalized.length === 0) return null;
  return prepareAxisFormatOverrideWithWorkbook(
    state,
    store,
    ranges,
    normalized,
    patch,
    formatHistoryBridgeFor(store)?.workbook ?? null,
  );
}

/** Workbook-aware variant used by mounted commands. Exported separately so
 * tests and host adapters can prepare an override without reaching into the
 * format-history bridge registry. */
export function prepareAxisFormatOverrideWithWorkbook(
  state: State,
  _store: SpreadsheetStore,
  ranges: readonly Range[],
  bands: readonly AxisFormatBand[],
  patch: Partial<CellFormat>,
  wb: WorkbookHandle | null,
  exceptions: readonly Addr[] = [],
  resolvePatch?: AxisFormatPatchResolver,
  budgetBase = 0,
  budgetKeys: ReadonlySet<string> = new Set(),
): AxisFormatPreparation | null {
  const sheet = ranges[0]?.sheet;
  if (sheet === undefined || ranges.some((range) => range.sheet !== sheet)) return null;
  const normalized = normalizedBands(bands);
  if (normalized.length === 0) return null;
  const existing = state.format.sheetAxisOverrides?.get(sheet);
  const baseline = existing?.baseline ?? wb?.getSheetFontAxisSnapshot(sheet) ?? emptySnapshot();
  const current = existing?.desired ?? wb?.getSheetFontAxisSnapshot(sheet) ?? baseline;
  let memoizedNativeDirectCells: readonly Addr[] | null = null;
  const readNativeDirectCells: (() => readonly Addr[]) | null = wb
    ? () => {
        if (memoizedNativeDirectCells === null) {
          memoizedNativeDirectCells = nativeDirectCellsForSelection(wb, sheet, normalized);
        }
        return memoizedNativeDirectCells;
      }
    : null;
  if (
    !axisBudgetAvailable(
      sheet,
      current,
      existing,
      normalized,
      readNativeDirectCells,
      exceptions,
      budgetBase,
      budgetKeys,
    )
  )
    return null;
  const columns = splitColumns(current, normalized, patch, wb, resolvePatch);
  const rowPreparation = splitRows(current, normalized, patch, wb, resolvePatch);
  if (!columns || !rowPreparation) return null;
  if (exceptions.length > 0 && !wb) return null;
  const rows = rowPreparation.rows;

  const direct = new Map<string, SheetAxisDirectXf>(
    (existing?.directXfs ?? []).map((entry) => [entry.key, cloneDirect(entry)]),
  );
  if (exceptions.length > 100_000) return null;
  const exceptionKeys = new Set(exceptions.map(addrKey));
  // Resolvers run during the splits above may mutate the workbook, so re-read
  // instead of reusing the budget-check snapshot.
  const nativeDirectCells = wb
    ? resolvePatch
      ? nativeDirectCellsForSelection(wb, sheet, normalized)
      : (readNativeDirectCells?.() ?? [])
    : ([] as readonly Addr[]);
  const selectedRow = (addr: Addr): boolean =>
    normalized.some(
      (band) => band.axis === 'row' && addr.row >= band.first && addr.row <= band.last,
    );
  const selectedColumn = (addr: Addr): boolean =>
    normalized.some(
      (band) => band.axis === 'column' && addr.col >= band.first && addr.col <= band.last,
    );
  const patchCache = new Map<string, number>();
  const directCells: Addr[] = [];
  // Existing sparse store entries are the direct-cell side of a mixed
  // selection. They stay in `format.formats`; collecting them here gives the
  // caller one bounded set for policy checks and history snapshots.
  for (const key of state.format.formats.keys()) {
    const addr = parseAddrKey(key);
    if (!addr || addr.sheet !== sheet || !selectedByBand(normalized, addr)) continue;
    directCells.push(addr);
  }

  // A near-complete axis selection has a bounded set of holes. Native axis
  // styles necessarily cover those holes too, so pin each excluded cell to
  // the effective XF it had before this operation. Keep the original direct
  // snapshot when a prior axis edit already owns the key; this makes the
  // whole sequence undoable without changing the store metadata for a cell
  // that was never selected.
  for (const addr of exceptions) {
    if (addr.sheet !== sheet) return null;
    const before = directSnapshot(wb as WorkbookHandle, addr);
    if (!before) return null;
    const key = before.key;
    const prior = direct.get(key);
    const preserved = before.before.effectiveXfIndex ?? before.before.directXfIndex;
    direct.set(key, {
      key,
      before: prior?.before ?? before.before,
      afterXfIndex: preserved,
    });
  }

  // A later ordinary cell edit may have changed a native direct XF after the
  // previous axis publication. Preserve that current value for corrections
  // outside this selection; replaying the old `afterXfIndex` would silently
  // undo the user's cell edit when the new axis state is flushed. The first
  // operation's `before` entry remains immutable for the history chain.
  if (wb) {
    for (const [key, priorDirect] of [...direct]) {
      const addr = parseAddrKey(key);
      if (
        !addr ||
        addr.sheet !== sheet ||
        exceptionKeys.has(key) ||
        selectedByBand(normalized, addr)
      )
        continue;
      const currentDirect = directSnapshot(wb, addr);
      if (!currentDirect) return null;
      // A user may have cleared the prior derived XF after the axis edit.
      // Native reports direct index 0 while the effective source is again a
      // row/column/default style. Keeping that stale entry would reapply an
      // explicit XF0 during the next unrelated axis flush, masking the
      // inherited source. Remove it from the new desired journal; the
      // previous override still retains its immutable entry for undo.
      if (currentDirect.before.source !== null && currentDirect.before.source !== 0) {
        direct.delete(key);
        continue;
      }
      direct.set(key, { ...priorDirect, afterXfIndex: currentDirect.before.directXfIndex });
    }
  }

  // A row style has native precedence over a column style. For a selected
  // column, patch only the sparse intersections with existing styled rows;
  // this avoids walking MAX_ROW while keeping inherited fills, charset and
  // theme metadata from the effective source XF.
  if (wb && normalized.some((band) => band.axis === 'column')) {
    // Use the current desired rows. A row style may have been introduced by
    // an earlier axis operation after the immutable baseline was captured.
    const styledRows = current.rows;
    let derivedCount = 0;
    for (const rowEntry of styledRows) {
      for (const band of normalized) {
        if (band.axis !== 'column') continue;
        const count = band.last - band.first + 1;
        derivedCount += count;
        if (derivedCount > 100_000) return null;
        for (let col = band.first; col <= band.last; col += 1) {
          const addr = { sheet, row: rowEntry.row, col };
          const effective = wb.getEffectiveStyle(sheet, rowEntry.row, col);
          const key = addrKey(addr);
          if (exceptionKeys.has(key)) continue;
          const priorDirect = direct.get(key);
          // Row styles win over columns. Once an intersection has been
          // materialized by an earlier column operation its effective source
          // is Cell (0), so retain it as a derived target as well.
          if (effective?.source !== 1 && !priorDirect) continue;
          const before = directSnapshot(wb, addr, effective ?? undefined);
          if (!before) return null;
          const baseXf =
            effective && effective.source !== 0 ? effective.xfIndex : before.before.directXfIndex;
          const resolvedPatch = resolvePatch?.(wb, baseXf, addr) ?? patch;
          const cacheKey = `${baseXf}|${JSON.stringify(resolvedPatch)}`;
          const afterXfIndex =
            patchCache.get(cacheKey) ?? patchXfIndexFromBase(wb, baseXf, resolvedPatch);
          if (afterXfIndex < 0) return null;
          patchCache.set(cacheKey, afterXfIndex);
          direct.set(key, {
            key,
            before: priorDirect?.before ?? before.before,
            afterXfIndex,
          });
        }
      }
    }
  }

  // A native row style takes precedence over column styles. For an unstyled
  // row, retain the first column segment as its compact base and materialize
  // only the differing column segments as sparse direct cells. This keeps a
  // uniform full-width column style at one row XF while preserving fills,
  // fonts and number formats in a small number of column segments.
  if (wb && normalized.some((band) => band.axis === 'row')) {
    const segments = columnSegments(current);
    let derivedCount = direct.size;
    const selectedRows = [...rowPreparation.baseByRow.keys()];
    for (const row of selectedRows) {
      const rowBase = rowPreparation.baseByRow.get(row) ?? 0;
      const explicitRow = rowsForSnapshot(current).has(row);
      if (explicitRow) continue;
      for (const segment of segments) {
        if (!explicitRow && segment.styleXf === rowBase) continue;
        const count = segment.last - segment.first + 1;
        derivedCount += count;
        if (derivedCount > 100_000) return null;
        for (let col = segment.first; col <= segment.last; col += 1) {
          const addr = { sheet, row, col };
          const key = addrKey(addr);
          if (exceptionKeys.has(key)) continue;
          const priorDirect = direct.get(key);
          const effective = wb.getEffectiveStyle(sheet, row, col);
          if (effective?.source !== 2 && effective?.source !== 3 && !priorDirect) continue;
          const before = directSnapshot(wb, addr, effective ?? undefined);
          if (!before) return null;
          const baseXf = effective?.xfIndex ?? before.before.directXfIndex;
          const resolvedPatch = resolvePatch?.(wb, baseXf, addr) ?? patch;
          const cacheKey = `${baseXf}|${JSON.stringify(resolvedPatch)}`;
          const afterXfIndex =
            patchCache.get(cacheKey) ?? patchXfIndexFromBase(wb, baseXf, resolvedPatch);
          if (afterXfIndex < 0) return null;
          patchCache.set(cacheKey, afterXfIndex);
          direct.set(key, {
            key,
            before: priorDirect?.before ?? before.before,
            afterXfIndex,
          });
        }
      }
    }
    // Existing derived intersections can have a Cell effective source even
    // when their column segment now matches the compact row base. They still
    // need the current patch, with their original undo baseline retained.
    for (const [key, priorDirect] of [...direct]) {
      const addr = parseAddrKey(key);
      if (
        !addr ||
        addr.sheet !== sheet ||
        exceptionKeys.has(key) ||
        !rowPreparation.baseByRow.has(addr.row)
      )
        continue;
      const effective = wb.getEffectiveStyle(sheet, addr.row, addr.col);
      const before = directSnapshot(wb, addr, effective ?? undefined);
      if (!effective || !before) return null;
      const resolvedPatch = resolvePatch?.(wb, effective.xfIndex, addr) ?? patch;
      const cacheKey = `${effective.xfIndex}|${JSON.stringify(resolvedPatch)}`;
      const afterXfIndex =
        patchCache.get(cacheKey) ?? patchXfIndexFromBase(wb, effective.xfIndex, resolvedPatch);
      if (afterXfIndex < 0) return null;
      patchCache.set(cacheKey, afterXfIndex);
      direct.set(key, { key, before: priorDirect.before, afterXfIndex });
    }
  }

  // Existing exception corrections become selected targets when a later
  // operation includes their row or column. Rebase their post-edit XF on the
  // current effective source while retaining the first operation's undo
  // baseline. Without this pass a subsequent axis patch would leave a stale
  // correction from the earlier hole selection.
  if (wb) {
    for (const [key, priorDirect] of [...direct]) {
      const addr = parseAddrKey(key);
      if (
        !addr ||
        addr.sheet !== sheet ||
        exceptionKeys.has(key) ||
        !selectedByBand(normalized, addr)
      )
        continue;
      const current = directSnapshot(wb, addr);
      if (!current) return null;
      const effectiveXf = current.before.effectiveXfIndex ?? current.before.directXfIndex;
      const resolvedPatch = resolvePatch?.(wb, effectiveXf, addr) ?? patch;
      const cacheKey = `${effectiveXf}|${JSON.stringify(resolvedPatch)}`;
      const afterXfIndex =
        patchCache.get(cacheKey) ?? patchXfIndexFromBase(wb, effectiveXf, resolvedPatch);
      if (afterXfIndex < 0) return null;
      patchCache.set(cacheKey, afterXfIndex);
      direct.set(key, { key, before: priorDirect.before, afterXfIndex });
    }

    // Physical cells with an explicit XF are sparse even when their value is
    // ordinary text/number data. They must participate in an axis edit: XF0
    // is still a direct cell record for native purposes, and carries opaque
    // metadata such as quotePrefix that the CellFormat store cannot model.
    // A row source supplies the raw base at a column intersection, so a
    // column edit preserves its unrelated fields while applying the requested
    // column patch. Selected rows likewise receive the requested patch.
    for (const addr of nativeDirectCells) {
      const key = addrKey(addr);
      if (exceptionKeys.has(key)) continue;
      const before = directSnapshot(wb, addr);
      if (!before) return null;
      const priorDirect = direct.get(key);
      if (!selectedRow(addr) && !selectedColumn(addr)) continue;
      let effectiveXfIndex = before.before.effectiveXfIndex;
      let effectiveSource = before.before.source;
      if (effectiveXfIndex === null) {
        const effective = wb.getEffectiveStyle(addr.sheet, addr.row, addr.col);
        effectiveXfIndex = effective?.xfIndex ?? null;
        effectiveSource = effective?.source ?? null;
      }
      const baseXf =
        effectiveXfIndex !== null && effectiveSource !== 0
          ? effectiveXfIndex
          : before.before.directXfIndex;
      const resolvedPatch = resolvePatch?.(wb, baseXf, addr) ?? patch;
      const cacheKey = `${baseXf}|${JSON.stringify(resolvedPatch)}`;
      const afterXfIndex =
        patchCache.get(cacheKey) ?? patchXfIndexFromBase(wb, baseXf, resolvedPatch);
      if (afterXfIndex < 0) return null;
      patchCache.set(cacheKey, afterXfIndex);
      direct.set(key, {
        key,
        before: priorDirect?.before ?? before.before,
        afterXfIndex,
      });
    }
  }
  if (direct.size > 100_000) return null;

  const fontPatch: SheetFontPatch = {
    ...(patch.fontFamily !== undefined ? { fontFamily: patch.fontFamily } : {}),
    ...(patch.fontSize !== undefined ? { fontSize: patch.fontSize } : {}),
  };

  return {
    sheet,
    directCells,
    override: {
      baseline: cloneSnapshot(baseline),
      desired: { columns, rows },
      ...(direct.size > 0 ? { directXfs: [...direct.values()] } : {}),
      ...(Object.keys(fontPatch).length > 0 ? { fontPatch } : {}),
    },
  };
}

interface ActiveAxisState {
  readonly baseline: SheetFontAxisSnapshot;
  readonly desired: SheetFontAxisSnapshot;
  readonly directXfs: readonly SheetAxisDirectXf[];
}

export interface ActiveAxisFormatSnapshot {
  readonly baseline: SheetFontAxisSnapshot;
  readonly desired: SheetFontAxisSnapshot;
  readonly directXfs: readonly SheetAxisDirectXf[];
}

const activeOverrides = new WeakMap<WorkbookHandle, Map<number, ActiveAxisState>>();

const cloneActiveAxisState = (value: ActiveAxisState): ActiveAxisFormatSnapshot => ({
  baseline: cloneSnapshot(value.baseline),
  desired: cloneSnapshot(value.desired),
  directXfs: value.directXfs.map(cloneDirect),
});

const activeFor = (wb: WorkbookHandle): Map<number, ActiveAxisState> => {
  let active = activeOverrides.get(wb);
  if (!active) {
    active = new Map();
    activeOverrides.set(wb, active);
  }
  return active;
};

const restoreAxes = (
  wb: WorkbookHandle,
  sheet: number,
  snapshot: SheetFontAxisSnapshot,
): boolean => {
  if (!wb.clearColumnStyle(sheet, 0, MAX_COL)) return false;
  for (const entry of snapshot.columns) {
    if (!wb.setColumnStyle(sheet, entry.first, entry.last, entry.styleXf)) return false;
  }
  const current = wb.getSheetFontAxisSnapshot(sheet);
  const rows = new Set<number>([
    ...snapshot.rows.map((entry) => entry.row),
    ...(current?.rows ?? []).map((entry) => entry.row),
  ]);
  for (const row of rows) if (!wb.clearRowStyle(sheet, row)) return false;
  for (const entry of snapshot.rows)
    if (!wb.setRowStyle(sheet, entry.row, entry.styleXf)) return false;
  return true;
};

/** Restore only the compact native row/column style records. The layout
 * attributes on the same native dimension records remain untouched. Structure
 * history uses this checkpoint after the native inverse because deleting a
 * styled row/column does not make the core recreate its style on insertion. */
export const restoreNativeAxisSnapshot = (
  wb: WorkbookHandle,
  sheet: number,
  snapshot: SheetFontAxisSnapshot,
): boolean => {
  const before = wb.getSheetFontAxisSnapshot(sheet);
  const restored = restoreAxes(wb, sheet, snapshot);
  if (restored) return true;
  // Native setters may fail after clearing one or more records. Best effort
  // restoration keeps a failed structure-history callback from leaving a
  // partially shifted axis. The caller still receives `false` and can report
  // any rollback failure as a strict mutation error.
  if (before !== null) restoreAxes(wb, sheet, before);
  return false;
};

const applyAxes = (wb: WorkbookHandle, sheet: number, snapshot: SheetFontAxisSnapshot): boolean => {
  if (!wb.clearColumnStyle(sheet, 0, MAX_COL)) return false;
  for (const entry of snapshot.columns) {
    if (!wb.setColumnStyle(sheet, entry.first, entry.last, entry.styleXf)) return false;
  }
  const current = wb.getSheetFontAxisSnapshot(sheet);
  const rows = new Set<number>([
    ...(current?.rows ?? []).map((entry) => entry.row),
    ...snapshot.rows.map((entry) => entry.row),
  ]);
  for (const row of rows) if (!wb.clearRowStyle(sheet, row)) return false;
  for (const entry of snapshot.rows)
    if (!wb.setRowStyle(sheet, entry.row, entry.styleXf)) return false;
  return true;
};

const applyDirect = (wb: WorkbookHandle, entries: readonly SheetAxisDirectXf[]): boolean => {
  for (const entry of entries) {
    const addr = parseAddrKey(entry.key);
    if (!addr || !wb.setCellXfIndex(addr.sheet, addr.row, addr.col, entry.afterXfIndex))
      return false;
  }
  return true;
};

const restoreDirect = (wb: WorkbookHandle, entries: readonly SheetAxisDirectXf[]): boolean => {
  const clearCellXfIndex = (
    wb as WorkbookHandle & {
      clearCellXfIndex?: (sheet: number, row: number, col: number) => boolean;
    }
  ).clearCellXfIndex;
  for (const entry of entries) {
    const addr = parseAddrKey(entry.key);
    if (!addr) continue;
    const inherited = entry.before.source !== null && entry.before.source !== 0;
    const xf = inherited ? entry.before.effectiveXfIndex : entry.before.directXfIndex;
    if (inherited && wb.capabilities.cellFormattingClear === true) {
      if (
        typeof clearCellXfIndex !== 'function' ||
        !clearCellXfIndex.call(wb, addr.sheet, addr.row, addr.col)
      )
        return false;
      // The native clear API removes the override for blank cells.  A
      // non-empty cell can still be treated as a cell-owned XF by the core's
      // resolver, even after clear reports success.  Verify the requested
      // inherited source and retain the visible/effective format through the
      // documented fallback in that case.  A reported clear failure above is
      // deliberately not converted into this fallback: callers must be able
      // to roll back the failed transaction atomically.
      const effective = wb.getEffectiveStyle(addr.sheet, addr.row, addr.col);
      if (
        effective?.source === entry.before.source &&
        (entry.before.effectiveXfIndex === null ||
          effective.xfIndex === entry.before.effectiveXfIndex)
      )
        continue;
      if (xf === null || !wb.setCellXfIndex(addr.sheet, addr.row, addr.col, xf)) return false;
      continue;
    }
    if (xf !== null && !wb.setCellXfIndex(addr.sheet, addr.row, addr.col, xf)) return false;
  }
  return true;
};

const currentDirectSnapshot = (
  wb: WorkbookHandle,
  entries: readonly SheetAxisDirectXf[],
): readonly SheetAxisDirectXf[] =>
  entries.flatMap((entry) => {
    const addr = parseAddrKey(entry.key);
    if (!addr) return [];
    const current = directSnapshot(wb, addr);
    return current
      ? [{ ...entry, before: current.before, afterXfIndex: current.before.directXfIndex }]
      : [];
  });

/** Flush the compact axis state and restore the previous state when a map
 * entry disappears during undo or Clear Formats. */
export function syncAxisFormatOverrides(
  wb: WorkbookHandle,
  store: SpreadsheetStore,
  sheets?: Iterable<number>,
): boolean {
  const active = activeFor(wb);
  const configured = store.getState().format.sheetAxisOverrides;
  // Keep legacy flushes a no-op on hosts that predate axis style APIs. An
  // actual axis override still fails strictly so callers never believe it was
  // published when the native surface is unavailable.
  if (wb.capabilities.sheetFontStyles !== true) {
    return (configured?.size ?? 0) === 0 && active.size === 0;
  }
  const targets = new Set<number>(
    sheets ?? store.getState().format.sheetAxisOverrides?.keys() ?? [],
  );
  for (const sheet of active.keys()) targets.add(sheet);
  let ok = true;
  for (const sheet of targets) {
    const override = store.getState().format.sheetAxisOverrides?.get(sheet);
    const previous = active.get(sheet);
    if (override) {
      if (
        previous &&
        previous.desired === override.desired &&
        previous.directXfs === override.directXfs
      ) {
        continue;
      }
      const baseline = previous?.baseline ?? override.baseline;
      const beforeAxes = wb.getSheetFontAxisSnapshot(sheet);
      const beforeDirect = currentDirectSnapshot(wb, override.directXfs ?? []);
      const applied =
        beforeAxes !== null &&
        applyAxes(wb, sheet, override.desired) &&
        applyDirect(wb, override.directXfs ?? []);
      if (!applied) {
        const axesRollback = beforeAxes !== null && restoreAxes(wb, sheet, beforeAxes);
        const directRollback = restoreDirect(wb, beforeDirect);
        ok = false;
        if (!axesRollback || !directRollback) ok = false;
        continue;
      }
      active.set(sheet, {
        baseline,
        desired: override.desired,
        directXfs: override.directXfs ?? [],
      });
      ok = applied && ok;
      continue;
    }
    if (!previous) continue;
    const beforeAxes = wb.getSheetFontAxisSnapshot(sheet);
    const beforeDirect = currentDirectSnapshot(wb, previous.directXfs);
    const axesRestored = beforeAxes !== null && restoreAxes(wb, sheet, previous.baseline);
    const directRestored = restoreDirect(wb, previous.directXfs);
    const restored = axesRestored && directRestored;
    if (!restored && beforeAxes !== null) {
      restoreAxes(wb, sheet, beforeAxes);
      restoreDirect(wb, beforeDirect);
    }
    ok = restored && ok;
    if (restored) active.delete(sheet);
  }
  return ok;
}

export const remapActiveAxisFormatOverrides = (
  wb: WorkbookHandle,
  mapIndex: (index: number) => number | null,
): void => {
  const active = activeOverrides.get(wb);
  if (!active) return;
  const next = new Map<number, ActiveAxisState>();
  for (const [sheet, value] of active) {
    const mapped = mapIndex(sheet);
    if (mapped === null) continue;
    const directXfs = value.directXfs.flatMap((entry) => {
      const addr = parseAddrKey(entry.key);
      if (!addr) return [entry];
      const mappedSheet = mapIndex(addr.sheet);
      return mappedSheet === null
        ? []
        : [{ ...entry, key: addrKey({ ...addr, sheet: mappedSheet }) }];
    });
    next.set(mapped, { ...value, directXfs });
  }
  activeOverrides.set(wb, next);
};

/** Keep the workbook-local active axis cache aligned with a row/column
 * structure edit. Native insertion/deletion has already moved its records;
 * remapping the cache prevents a later unrelated format flush from replaying
 * the old coordinates. */
export const remapActiveAxisFormatOverridesForStructure = (
  wb: WorkbookHandle,
  sheet: number,
  shift: AxisStructureShift,
): void => {
  const active = activeOverrides.get(wb);
  const current = active?.get(sheet);
  if (!active || !current) return;
  const remapped = remapAxisFormatOverride(
    {
      baseline: current.baseline,
      desired: current.desired,
      ...(current.directXfs.length > 0 ? { directXfs: current.directXfs } : {}),
    },
    shift,
  );
  active.set(sheet, {
    baseline: remapped.baseline,
    desired: remapped.desired,
    directXfs: remapped.directXfs ?? [],
  });
};

export const captureActiveAxisFormatOverrideForStructure = (
  wb: WorkbookHandle,
  sheet: number,
): ActiveAxisFormatSnapshot | null => {
  const value = activeOverrides.get(wb)?.get(sheet);
  return value ? cloneActiveAxisState(value) : null;
};

export const restoreActiveAxisFormatOverrideForStructure = (
  wb: WorkbookHandle,
  sheet: number,
  snapshot: ActiveAxisFormatSnapshot | null,
): void => {
  const active = activeOverrides.get(wb);
  if (!active) {
    if (snapshot === null) return;
    activeOverrides.set(wb, new Map([[sheet, cloneActiveAxisState(snapshot)]]));
    return;
  }
  if (snapshot === null) active.delete(sheet);
  else active.set(sheet, cloneActiveAxisState(snapshot));
  if (active.size === 0) activeOverrides.delete(wb);
};

/** Expose the same pure transformation used by the active cache for callers
 * that need to inspect a before/after state without mutating the workbook. */
export const remapAxisSnapshotForStructure = remapAxisSnapshot;

export const remapAxisFormatOverrides = (
  store: SpreadsheetStore,
  mapIndex: (index: number) => number | null,
): void => {
  store.setState((state) => {
    const current = state.format.sheetAxisOverrides;
    if (!current) return state;
    const next = new Map<number, SheetAxisFormatOverride>();
    let changed = false;
    for (const [sheet, override] of current) {
      const mapped = mapIndex(sheet);
      if (mapped === null) {
        changed = true;
        continue;
      }
      changed ||= mapped !== sheet;
      const directXfs = override.directXfs?.flatMap((entry) => {
        const addr = parseAddrKey(entry.key);
        if (!addr) return [entry];
        const mappedSheet = mapIndex(addr.sheet);
        if (mappedSheet === null) {
          changed = true;
          return [];
        }
        const key = addrKey({ ...addr, sheet: mappedSheet });
        changed ||= key !== entry.key;
        return [{ ...entry, key }];
      });
      next.set(mapped, {
        ...override,
        ...(directXfs ? { directXfs } : {}),
      });
    }
    return changed ? { ...state, format: { ...state.format, sheetAxisOverrides: next } } : state;
  });
};

export const axisFormatOverrideFor = (
  state: State,
  sheet: number,
): SheetAxisFormatOverride | undefined => state.format.sheetAxisOverrides?.get(sheet);
