import {
  CellShift,
  type TableColumnPatch,
  type TableStyleOptionsPatch,
  TableTotalsRowFunction,
} from '@libraz/formulon';
import { mirrorBandFormatsDown } from '../commands/cell-shift.js';
import { withFilterRange } from '../commands/filter-ranges.js';
import {
  defaultObjectNames,
  overlayStyleName,
  type TableColumnTotals,
  type TableOverlay,
  tableColumns,
  totalsCellAt,
} from '../commands/format-as-table.js';
import { sameRange, sameRangeList } from '../store/selection-geometry.js';
import {
  mutators,
  type SpreadsheetStore,
  type State,
  type ValueFilterCriteria,
} from '../store/store.js';
import type { TablesSlice } from '../store/types.js';
import { formatA1Range, MAX_ROW } from './address.js';
import { forgetTableAutoFilterImport } from './auto-filter-sync.js';
import { parseRangeRef } from './range-resolver.js';
import { TableStyleMirror } from './table-style-sync.js';
import {
  sameTotals,
  totalsFormulaText,
  totalsFromCell,
  totalsFromDetails,
  totalsFunctionFromEngine,
  totalsFunctionToEngine,
} from './table-sync.js';
import { defaultTotals } from './table-totals-defaults.js';
import type { Addr, Range, TablePatch } from './types.js';
import type { WorkbookHandle } from './workbook-handle.js';

/**
 * Reconciles the store's engine-backed table overlays (those carrying a
 * `name`) with the engine's worksheet tables, so every table edit, removal and
 * undo reaches the saved file. Overlays are matched to engine tables by name;
 * indexes are re-resolved for every write because a removal shifts them.
 *
 * A table is written with its style, show-flags and totals row; a rename goes
 * through `updateTable` so structured references follow, and a column-count
 * change re-creates the table and re-applies its totals.
 *
 * A table owns an AutoFilter, so creating, moving or removing one also edits
 * `ui.filterRanges` / `ui.filterCriteria`. That follow-up write is a nested
 * store notification, which mirrors the filter slices from the newer state.
 */

/** Filter-slice edits implied by the table writes of one notification. */
interface FilterEdits {
  /** Ranges of tables whose filter range and criteria leave the store. */
  dropped: Range[];
  /** Filter ranges of tables created or moved; criteria carry over from `from`. */
  added: { range: Range; from: Range | null }[];
  /** Totals models read back from the engine after a write, by overlay id. */
  totals: Map<string, TableColumnTotals[]>;
  /** Overlays the engine refused, by id: restored to their previous shape, or
   *  removed (null) when they had none, so screen and file agree. */
  rejected: Map<string, TableOverlay | null>;
  /** Ranges the engine gave a table when it added or removed its totals row, by overlay id. */
  ranges: Map<string, Range>;
}

const nameKey = (name: string): string => name.toLocaleLowerCase();

const namedOverlays = (tables: readonly TableOverlay[]): Map<string, TableOverlay> => {
  const out = new Map<string, TableOverlay>();
  for (const table of tables) if (table.name) out.set(nameKey(table.name), table);
  return out;
};

/** True when `filter` is the AutoFilter of a table spanning `table`. */
const isFilterOfTable = (table: Range, filter: Range): boolean =>
  table.sheet === filter.sheet &&
  table.r0 === filter.r0 &&
  table.c0 === filter.c0 &&
  table.c1 === filter.c1 &&
  filter.r1 <= table.r1;

const engineTable = (wb: WorkbookHandle, name: string) => {
  const index = wb.tableIndexByName(name);
  return index === null ? null : (wb.getTables()[index] ?? null);
};

const tableFilterRange = (wb: WorkbookHandle, index: number, sheet: number): Range | null => {
  const filter = wb.capabilities.typedTableAutoFilter ? wb.getTableAutoFilter(index) : null;
  if (!filter) return null;
  return {
    sheet,
    r0: filter.range.firstRow,
    c0: filter.range.firstCol,
    r1: filter.range.lastRow,
    c1: filter.range.lastCol,
  };
};

/** Remove the table's AutoFilter, or give it back one over the header and
 *  body, so the header filter buttons follow `table.filterButton`. Row
 *  visibility is left as it is. */
const syncTableFilterButton = (wb: WorkbookHandle, index: number, table: TableOverlay): void => {
  if (table.filterButton === false) {
    wb.removeTableAutoFilter(index);
    return;
  }
  if (wb.getTableAutoFilter(index)) return;
  const { r0, c0, r1, c1 } = table.range;
  wb.setTableAutoFilter(index, {
    range: { firstRow: r0, firstCol: c0, lastRow: table.showTotal ? r1 - 1 : r1, lastCol: c1 },
    columns: [],
    sort: null,
  });
};

const createEngineTable = (wb: WorkbookHandle, table: TableOverlay & { name: string }): number => {
  const styleName = overlayStyleName(table);
  const index = wb.createTable(table.range.sheet, {
    range: formatA1Range(table.range, { collapse: false }),
    name: table.name,
    columns: tableColumns(wb, table.range, table.showHeader, defaultObjectNames(wb).column),
    ...(styleName ? { styleName } : {}),
    showRowStripes: table.banded,
    showFirstColumn: !!table.firstCol,
    showLastColumn: !!table.lastCol,
    showColumnStripes: !!table.colStripes,
    headerRow: table.showHeader,
    totalsRow: table.showTotal,
  });
  // An empty style name keeps the style element and its flags, unnamed.
  if (index >= 0 && !styleName) wb.updateTableStyleOptions(index, { styleName: null });
  return index;
};

/** Label the totals row of a new total row carries in its first column. */
const TOTALS_LABEL = 'Total';

/** Defaults for a new total row that has no totals yet; a table that already has totals is left alone. */
const applyDefaultTotals = (
  wb: WorkbookHandle,
  index: number,
  table: TableOverlay,
  label = TOTALS_LABEL,
): void => {
  const details = wb.getTableStyleDetails(index);
  if (!details || details.columns.length === 0) return;
  if (
    details.columns.some(
      (c) => c.totalsRowFunction !== TableTotalsRowFunction.None || c.totalsRowLabel,
    )
  ) {
    return;
  }
  writeTotals(
    wb,
    index,
    defaultTotals(wb, table.range, table.showHeader, details.columns.length, label),
  );
};

/** Put `label` on the first column when the re-created table lost every label. */
const keepLabel = (wb: WorkbookHandle, index: number, label: string): void => {
  const columns = wb.getTableStyleDetails(index)?.columns ?? [];
  if (
    columns.some((c) => c.totalsRowLabel) ||
    columns[0]?.totalsRowFunction !== TableTotalsRowFunction.None
  ) {
    return;
  }
  wb.updateTable(index, { columns: [{ totalsRowLabel: label }] });
};

/** Write the columns of `desired` the engine holds differently. True when any
 *  column was written. */
const writeTotals = (
  wb: WorkbookHandle,
  index: number,
  desired: readonly TableColumnTotals[] | undefined,
): boolean => {
  const columns = wb.getTableStyleDetails(index)?.columns ?? [];
  let changed = false;
  const patches = columns.map((c, i): TableColumnPatch => {
    const want = desired?.[i];
    if (!want) return {};
    const same =
      want.function === totalsFunctionFromEngine(c.totalsRowFunction) &&
      (want.label ?? '') === (c.totalsRowLabel ?? '') &&
      (want.function !== 'custom' ||
        (want.formula ?? '') === totalsFormulaText(c.totalsRowFormula));
    if (same) return {};
    changed = true;
    return {
      totalsRowFunction: totalsFunctionToEngine(want.function),
      totalsRowLabel: want.label ? want.label : null,
      ...(want.function === 'custom' ? { totalsRowFormula: want.formula ?? '' } : {}),
    };
  });
  return changed && wb.updateTable(index, { columns: patches });
};

/** Totals settings of the table at `index` by lower-cased column name. */
const captureTotals = (wb: WorkbookHandle, index: number) => {
  const out = new Map<string, TableColumnPatch>();
  for (const c of wb.getTableStyleDetails(index)?.columns ?? []) {
    if (c.totalsRowFunction === TableTotalsRowFunction.None && !c.totalsRowLabel) continue;
    out.set(nameKey(c.name), {
      totalsRowFunction: c.totalsRowFunction,
      totalsRowLabel: c.totalsRowLabel || null,
      ...(c.totalsRowFormula ? { totalsRowFormula: c.totalsRowFormula } : {}),
    });
  }
  return out;
};

/** Re-apply captured totals to the columns of a re-created table that kept their name. */
const restoreTotals = (
  wb: WorkbookHandle,
  index: number,
  totals: ReturnType<typeof captureTotals>,
): void => {
  const columns = wb.getTableStyleDetails(index)?.columns ?? [];
  const patches = columns.map((c): TableColumnPatch => totals.get(nameKey(c.name)) ?? {});
  if (patches.some((patch) => Object.keys(patch).length > 0))
    wb.updateTable(index, { columns: patches });
};

const removeEngineTable = (wb: WorkbookHandle, index: number): boolean => {
  if (!wb.removeTable(index)) return false;
  forgetTableAutoFilterImport(wb, index);
  return true;
};

/** The writes that turn engine table `before` into `after`: a rename, a
 *  range / header / totals patch and a style patch. */
interface TableUpdate {
  name?: string;
  patch: TablePatch & { range?: string };
  style: { -readonly [K in keyof TableStyleOptionsPatch]: TableStyleOptionsPatch[K] };
}

const tableUpdate = (before: TableOverlay, after: TableOverlay): TableUpdate => {
  const update: TableUpdate = { patch: {}, style: {} };
  const { patch, style } = update;
  if (!sameRange(before.range, after.range)) {
    patch.range = formatA1Range(after.range, { collapse: false });
  }
  if (before.name !== after.name && after.name) update.name = after.name;
  if (overlayStyleName(before) !== overlayStyleName(after)) {
    style.styleName = overlayStyleName(after) || null;
  }
  if (before.banded !== after.banded) style.showRowStripes = after.banded;
  if (!!before.firstCol !== !!after.firstCol) style.showFirstColumn = !!after.firstCol;
  if (!!before.lastCol !== !!after.lastCol) style.showLastColumn = !!after.lastCol;
  if (!!before.colStripes !== !!after.colStripes) style.showColumnStripes = !!after.colStripes;
  if (before.showHeader !== after.showHeader) patch.headerRow = after.showHeader;
  if (before.showTotal !== after.showTotal) patch.totalsRow = after.showTotal;
  return update;
};

const isEmptyUpdate = (update: TableUpdate): boolean =>
  update.name === undefined &&
  Object.keys(update.patch).length === 0 &&
  Object.keys(update.style).length === 0;

/** Apply `update` to the table at `index`; false when the engine refuses a part. */
const applyTableUpdate = (wb: WorkbookHandle, index: number, update: TableUpdate): boolean =>
  (update.name === undefined || wb.renameTable(index, update.name)) &&
  (Object.keys(update.patch).length === 0 || wb.updateTable(index, update.patch)) &&
  (Object.keys(update.style).length === 0 || wb.updateTableStyleOptions(index, update.style));

/** Apply `edits` to the filter slices; returns `state` itself when nothing changes. */
function applyFilterEdits(state: State, edits: FilterEdits): State {
  const dropped = (range: Range): boolean =>
    edits.dropped.some((table) => isFilterOfTable(table, range));
  const carried: ValueFilterCriteria[] = [];
  for (const { range, from } of edits.added) {
    if (!from) continue;
    for (const entry of state.ui.filterCriteria) {
      if (!sameRange(entry.range, from) || entry.byCol < range.c0 || entry.byCol > range.c1) {
        continue;
      }
      carried.push({ ...entry, range: { ...range } });
    }
  }
  const tables =
    edits.rejected.size > 0 || edits.totals.size > 0 || edits.ranges.size > 0
      ? state.tables.tables.flatMap((table) => {
          if (edits.rejected.has(table.id)) {
            const restored = edits.rejected.get(table.id);
            return restored ? [restored] : [];
          }
          const totals = edits.totals.get(table.id);
          const range = edits.ranges.get(table.id);
          return [{ ...table, ...(totals ? { totals } : {}), ...(range ? { range } : {}) }];
        })
      : state.tables.tables;
  let next: State = {
    ...state,
    tables: { ...state.tables, tables },
    ui: { ...state.ui, filterRanges: state.ui.filterRanges.filter((range) => !dropped(range)) },
  };
  for (const { range } of edits.added) {
    next = { ...next, ui: { ...next.ui, filterRanges: withFilterRange(next, range) } };
  }
  const kept = next.ui.filterRanges;
  const criteria = state.ui.filterCriteria.filter(
    (entry) =>
      !dropped(entry.range) &&
      !edits.added.some(({ from }) => from && sameRange(from, entry.range)) &&
      kept.some((range) => sameRange(range, entry.range)),
  );
  const changed =
    tables !== state.tables.tables ||
    !sameRangeList(state.ui.filterRanges, kept) ||
    criteria.length !== state.ui.filterCriteria.length ||
    carried.length > 0;
  if (!changed) return state;
  return { ...next, ui: { ...next.ui, filterCriteria: [...criteria, ...carried] } };
}

export class TableMirror {
  constructor(private readonly handle: WorkbookHandle) {
    this.styles = new TableStyleMirror(handle);
    // Every cell write funnels into a value event, so an edit of a totals cell is read back
    // here whether it came from the editor, paste, fill, find/replace or an undo replay.
    handle.subscribe((event) => {
      if (event.kind === 'value') this.adoptTotalsCell(event.addr);
    });
  }

  /** The store the last sync ran against; totals-cell edits are applied to it. */
  private store: SpreadsheetStore | null = null;
  private applying = false;

  /** Record what a totals-row cell now holds as its column's totals setting. */
  private adoptTotalsCell(addr: Addr): void {
    const store = this.store;
    if (!store || this.applying) return;
    const hit = totalsCellAt(store.getState(), addr.sheet, addr.row, addr.col);
    if (!hit) return;
    const totals = [...(hit.table.totals ?? [])];
    if (hit.column >= totals.length) return;
    const next = totalsFromCell(this.handle, addr);
    if (sameTotals([next], [totals[hit.column] as TableColumnTotals])) return;
    totals[hit.column] = next;
    mutators.upsertTableOverlay(store, { ...hit.table, totals });
  }

  private readonly styles: TableStyleMirror;

  /** Tables slice the engine already carries. */
  private last: TablesSlice | null = null;

  /** Adopt `slice` as the engine baseline without writing it. */
  rebase(slice: TablesSlice | null): void {
    this.last = slice;
  }

  /** Write the store's table changes to the engine. Returns true when the
   *  filter follow-up notified the store, which then mirrored every slice. */
  sync(store: SpreadsheetStore): boolean {
    // Read the store, not a listener argument: another listener's nested
    // write can deliver an older state after a newer one.
    this.store = store;
    const slice = store.getState().tables;
    const previousSlice = this.last;
    if (slice === previousSlice) return false;
    this.last = slice;
    this.applying = true;
    try {
      // One recalc for all per-cell header/totals edits of the pass.
      return this.handle.withBatchedRecalc(() => this.apply(store, previousSlice, slice));
    } catch (error) {
      // Nothing is known to have reached the engine; retry from the old baseline.
      this.last = previousSlice;
      throw error;
    } finally {
      this.applying = false;
    }
  }

  private apply(store: SpreadsheetStore, previousSlice: TablesSlice | null, slice: TablesSlice) {
    const wb = this.handle;
    // Styles first: a table may name a custom style the same change creates.
    if (TableStyleMirror.changed(previousSlice, slice)) {
      const rejected = this.styles.sync(previousSlice, slice);
      if (rejected.length > 0) this.dropStyles(store, new Set(rejected));
    }
    const tables = slice.tables;
    const previous = previousSlice?.tables ?? null;
    if (tables === previous) return false;
    if (!wb.capabilities.tableMutate) return false;
    const before = namedOverlays(previous ?? []);
    const after = namedOverlays(tables);
    const edits: FilterEdits = {
      dropped: [],
      added: [],
      totals: new Map(),
      rejected: new Map(),
      ranges: new Map(),
    };
    // An overlay keeping its id under a new name is a rename, not a removal.
    const beforeById = new Map<string, TableOverlay>();
    for (const table of before.values()) beforeById.set(table.id, table);
    const renamedFrom = new Map<string, TableOverlay>();
    for (const [key, table] of after) {
      const old = beforeById.get(table.id);
      if (old && nameKey(old.name as string) !== key) renamedFrom.set(key, old);
    }
    const renamedAway = new Set([...renamedFrom.values()].map((t) => nameKey(t.name as string)));
    let cellsChanged = false;
    // Removals first, so a table re-created over the same cells cannot overlap.
    for (const [key, table] of before) {
      if (after.has(key) || renamedAway.has(key) || !table.name) continue;
      const existing = engineTable(wb, table.name);
      if (existing) removeEngineTable(wb, existing.index);
      edits.dropped.push(table.range);
    }
    for (const [key, table] of after) {
      const name = table.name as string;
      const renamedOld = renamedFrom.get(key) ?? null;
      const old = renamedOld ?? before.get(key) ?? null;
      const existing = engineTable(wb, (old?.name ?? name) as string);
      if (existing && !old) continue;
      if (renamedOld && engineTable(wb, name)) {
        edits.rejected.set(table.id, renamedOld);
        continue;
      }
      let index = existing?.index ?? -1;
      let from: Range | null = null;
      if (existing && old) {
        const update = tableUpdate(old, table);
        const { patch } = update;
        const totalsChanged = !sameTotals(old.totals, table.totals);
        const filterToggled = (old.filterButton !== false) !== (table.filterButton !== false);
        if (isEmptyUpdate(update) && !totalsChanged && !filterToggled) continue;
        from = tableFilterRangeIn(store.getState(), old.range);
        const reshaped =
          old.range.sheet !== table.range.sheet ||
          old.range.c1 - old.range.c0 !== table.range.c1 - table.range.c0;
        if (!reshaped) {
          // A flip the engine already holds (a row delete took the totals row) only moves the range.
          const details = wb.getTableStyleDetails(existing.index);
          if (details && patch.headerRow === details.headerRowCount > 0) delete patch.headerRow;
          if (details && patch.totalsRow === details.totalsRowCount > 0) delete patch.totalsRow;
          const flip = flipShiftRow(old.range, patch);
          // The engine shapes a flipped table itself; the overlay adopts its ref.
          if (flip !== null) delete patch.range;
          const cells = flip === null ? null : cellsFrom(wb, old.range, flip);
          if (!isEmptyUpdate(update) && !applyTableUpdate(wb, existing.index, update)) {
            edits.rejected.set(table.id, old);
            continue;
          }
          if (patch.headerRow === false && !closeHeaderGap(wb, existing.index, old, table)) {
            wb.updateTable(existing.index, { headerRow: true });
            edits.rejected.set(table.id, old);
            continue;
          }
          if (flip !== null && cells) {
            this.adoptFlippedRange(store, edits, table, existing.index, old.range, flip, cells);
            cellsChanged = true;
          }
          if (totalsChanged && writeTotals(wb, existing.index, table.totals) && table.showTotal) {
            cellsChanged = true;
          }
          this.adoptTotals(edits, table, existing.index);
          // Rewritten structured references hold no value until the next recalc.
          if (update.name !== undefined) cellsChanged = true;
          // Turning the header on gives the table its AutoFilter back.
          if (filterToggled || (patch.headerRow === true && table.filterButton === false)) {
            syncTableFilterButton(wb, existing.index, table);
          }
          const filterMoved =
            patch.range !== undefined ||
            patch.headerRow !== undefined ||
            patch.totalsRow !== undefined ||
            filterToggled;
          if (!filterMoved) continue;
        } else {
          // A new column count or sheet needs a new table; when the engine
          // refuses it, the old table comes back.
          const totals = captureTotals(wb, existing.index);
          removeEngineTable(wb, existing.index);
          index = createEngineTable(wb, { ...table, name });
          if (index < 0) {
            const oldName = old.name as string;
            const restored = createEngineTable(wb, { ...old, name: oldName }) >= 0;
            if (restored) restoreTotals(wb, engineTable(wb, oldName)?.index ?? -1, totals);
            if (!restored) edits.dropped.push(old.range);
            edits.rejected.set(table.id, restored ? old : null);
            continue;
          }
          // Keep the label the table already carried instead of defaulting it.
          const label = old.totals?.find((t) => t.label)?.label;
          if (totals.size > 0) {
            restoreTotals(wb, index, totals);
            if (label) keepLabel(wb, index, label);
          } else if (table.showTotal) applyDefaultTotals(wb, index, table, label);
          this.adoptTotals(edits, table, index);
          cellsChanged = true;
        }
      } else {
        index = createEngineTable(wb, { ...table, name });
        if (index >= 0) {
          writeTotals(wb, index, table.totals);
          if (table.showTotal) {
            applyDefaultTotals(wb, index, table);
            cellsChanged = true;
          }
          this.adoptTotals(edits, table, index);
        }
      }
      if (old) edits.dropped.push(old.range);
      if (index < 0) {
        edits.rejected.set(table.id, null);
        continue;
      }
      // A table created or re-created without its filter buttons drops the AutoFilter.
      if (table.filterButton === false) wb.removeTableAutoFilter(index);
      const range = tableFilterRange(wb, index, table.range.sheet);
      if (range) edits.added.push({ range, from });
    }
    let wrote = false;
    if (
      edits.dropped.length > 0 ||
      edits.added.length > 0 ||
      edits.totals.size > 0 ||
      edits.rejected.size > 0 ||
      edits.ranges.size > 0
    ) {
      store.setState((state) => {
        const next = applyFilterEdits(state, edits);
        wrote = next !== state;
        // Restored or rejected overlays already match the engine.
        this.last = next.tables;
        return next;
      });
    }
    // After the filter follow-up, whose engine writes this recalc then covers.
    if (cellsChanged) this.refreshCells(store);
    return wrote;
  }

  /** Remove custom styles the engine refused, and the PivotTable assignments naming them. */
  private dropStyles(store: SpreadsheetStore, ids: ReadonlySet<string>): void {
    store.setState((state) => {
      const t = state.tables;
      const next: TablesSlice = {
        ...t,
        customTableStyles: (t.customTableStyles ?? []).filter((s) => !ids.has(s.id)),
        customPivotTableStyles: (t.customPivotTableStyles ?? []).filter((s) => !ids.has(s.id)),
        pivotTableStyles: (t.pivotTableStyles ?? []).filter((a) => !ids.has(a.styleId)),
      };
      this.last = next;
      return { ...state, tables: next };
    });
  }

  /** Take the range the engine gave the table at `index` for a header or totals flip, and move
   *  the store's formats from row `from` by the rows the ref grew or shrank at its bottom, when
   *  the engine shifted those cells (`cells`, captured from `from` before the flip). */
  private adoptFlippedRange(
    store: SpreadsheetStore,
    edits: FilterEdits,
    table: TableOverlay,
    index: number,
    before: Range,
    from: number,
    cells: ReadonlyMap<string, string>,
  ): void {
    const range = engineTableRange(this.handle, index, table.range.sheet);
    if (!range) return;
    if (!sameRange(range, table.range)) edits.ranges.set(table.id, range);
    // The shift also moves other tables in the same columns below.
    for (const other of store.getState().tables.tables) {
      if (other.id === table.id || !other.name || other.range.sheet !== table.range.sheet) continue;
      const otherIndex = this.handle.tableIndexByName(other.name);
      const moved =
        otherIndex === null ? null : engineTableRange(this.handle, otherIndex, other.range.sheet);
      if (moved && !sameRange(moved, other.range)) edits.ranges.set(other.id, moved);
    }
    const delta = range.r1 - before.r1;
    if (delta === 0) return;
    const band: Range = {
      sheet: before.sheet,
      r0: from + (delta < 0 ? -1 : 0),
      c0: before.c0,
      r1: MAX_ROW,
      c1: before.c1,
    };
    if (
      cells.size > 0 &&
      !sameShiftedCells(cells, cellsFrom(this.handle, range, from + delta), delta)
    ) {
      return;
    }
    mirrorBandFormatsDown(store, this.handle, band, delta);
  }

  /** Queue the engine's totals of the table at `index` for the overlay when its model differs. */
  private adoptTotals(edits: FilterEdits, table: TableOverlay, index: number): void {
    const details = this.handle.getTableStyleDetails(index);
    if (!details) return;
    const totals = totalsFromDetails(details);
    if (!sameTotals(totals, table.totals)) edits.totals.set(table.id, totals);
  }

  /** Pull the active sheet's cells after the engine wrote totals-row cells. */
  private refreshCells(store: SpreadsheetStore): void {
    const wb = this.handle;
    wb.recalcAuto();
    wb.withEngineSyncMuted(() =>
      mutators.replaceCells(store, wb.cells(store.getState().data.sheetIndex)),
    );
  }
}

/** The engine's current range of the table at workbook-wide `index`. */
function engineTableRange(wb: WorkbookHandle, index: number, sheet: number): Range | null {
  const table = wb.getTables()[index];
  const parsed = table ? parseRangeRef(table.ref ?? '') : null;
  return parsed ? { sheet, r0: parsed.r0, c0: parsed.c0, r1: parsed.r1, c1: parsed.c1 } : null;
}

/** Value signatures of the populated cells from row `from` down across the columns of `table`. */
function cellsFrom(wb: WorkbookHandle, table: Range, from: number): Map<string, string> {
  const out = new Map<string, string>();
  for (const cell of wb.physicalCells(table.sheet)) {
    const { row, col } = cell.addr;
    if (row < from || col < table.c0 || col > table.c1) continue;
    out.set(`${row}:${col}`, JSON.stringify([cell.value, cell.formula]));
  }
  return out;
}

/** The first row whose cells a header or totals flip in `patch` can move: the data rows for a
 *  header (down when it is turned on over an occupied row, back up when an undo closes the gap),
 *  the rows below the table for a totals row. Null when `patch` flips neither. */
function flipShiftRow(table: Range, patch: TablePatch): number | null {
  if (patch.headerRow === true) return table.r0;
  if (patch.headerRow === false) return table.r0 + 1;
  if (patch.totalsRow !== undefined) return table.r1 + 1;
  return null;
}

/**
 * After the header of `old` was turned off, move the table back up when `table` asks for the
 * range just above the engine's: undoing a header that shifted the table down leaves the
 * emptied header row as a gap above the data. False when the engine refuses the shift.
 */
function closeHeaderGap(
  wb: WorkbookHandle,
  index: number,
  old: TableOverlay,
  table: TableOverlay,
): boolean {
  const range = engineTableRange(wb, index, table.range.sheet);
  const want = table.range;
  if (
    !range ||
    sameRange(want, old.range) ||
    range.r0 !== want.r0 + 1 ||
    range.r1 !== want.r1 + 1
  ) {
    return true;
  }
  return wb.engineShiftCells({ ...want, r1: want.r0 }, CellShift.Up);
}

/** Whether every cell of `before` reappears `delta` rows away in `after`. */
function sameShiftedCells(
  before: ReadonlyMap<string, string>,
  after: ReadonlyMap<string, string>,
  delta: number,
): boolean {
  for (const [key, signature] of before) {
    const [row = 0, col = 0] = key.split(':').map(Number);
    if (after.get(`${row + delta}:${col}`) !== signature) return false;
  }
  return true;
}

/** The store filter range belonging to the table spanning `table`, if any. */
function tableFilterRangeIn(state: State, table: Range): Range | null {
  return state.ui.filterRanges.find((range) => isFilterOfTable(table, range)) ?? null;
}
