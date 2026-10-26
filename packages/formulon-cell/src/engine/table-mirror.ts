import { withFilterRange } from '../commands/filter-ranges.js';
import {
  overlayStyleName,
  type TableColumnTotals,
  type TableOverlay,
  tableColumns,
  totalsCellAt,
} from '../commands/format-as-table.js';
import { sameRange } from '../store/selection-geometry.js';
import {
  mutators,
  type SpreadsheetStore,
  type State,
  type ValueFilterCriteria,
} from '../store/store.js';
import type { TablesSlice } from '../store/types.js';
import { formatA1Range } from './address.js';
import { forgetTableAutoFilterImport } from './auto-filter-sync.js';
import { TableStyleMirror } from './table-style-sync.js';
import { sameTotals, totalsFromCell, totalsFromDetails } from './table-sync.js';
import { defaultTotals } from './table-totals-defaults.js';
import type { Addr, Range, TableUpdateInput } from './types.js';
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

const engineTable = (wb: WorkbookHandle, name: string) =>
  wb.getTables().find((table) => nameKey(table.name) === nameKey(name)) ?? null;

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

const createEngineTable = (
  wb: WorkbookHandle,
  table: TableOverlay & { name: string },
  displayName = table.name,
): number =>
  wb.createTable({
    sheetIndex: table.range.sheet,
    ref: formatA1Range(table.range, { collapse: false }),
    name: table.name,
    displayName,
    columns: tableColumns(wb, table.range, table.showHeader),
    styleName: overlayStyleName(table),
    showRowStripes: table.banded,
    showFirstColumn: !!table.firstCol,
    showLastColumn: !!table.lastCol,
    showColumnStripes: !!table.colStripes,
    headerRow: table.showHeader,
    totalsRow: table.showTotal,
  });

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
  if (details.columns.some((c) => c.totalsRowFunction !== 'none' || c.totalsRowLabel)) return;
  writeTotals(
    wb,
    index,
    defaultTotals(wb, table.range, table.showHeader, details.columns.length, label),
  );
};

/** Put `label` on the first column when the re-created table lost every label. */
const keepLabel = (wb: WorkbookHandle, index: number, label: string): void => {
  const columns = wb.getTableStyleDetails(index)?.columns ?? [];
  if (columns.some((c) => c.totalsRowLabel) || columns[0]?.totalsRowFunction !== 'none') return;
  wb.setTableColumnTotals(index, 0, { label });
};

/** Write the columns of `desired` the engine holds differently. True when any
 *  column was written. */
const writeTotals = (
  wb: WorkbookHandle,
  index: number,
  desired: readonly TableColumnTotals[] | undefined,
): boolean => {
  const columns = wb.getTableStyleDetails(index)?.columns ?? [];
  let wrote = false;
  columns.forEach((c, i) => {
    const want = desired?.[i];
    if (!want) return;
    const same =
      want.function === c.totalsRowFunction &&
      (want.label ?? '') === c.totalsRowLabel &&
      (want.function !== 'custom' || (want.formula ?? '') === c.totalsRowFormula);
    if (same) return;
    const ok = wb.setTableColumnTotals(index, i, {
      function: want.function,
      label: want.label ?? '',
      ...(want.function === 'custom' ? { formula: want.formula ?? '' } : {}),
    });
    wrote = wrote || ok;
  });
  return wrote;
};

/** Blank the cells of `row` across the columns of `range`. */
const clearRow = (wb: WorkbookHandle, range: Range, row: number): void => {
  for (let col = range.c0; col <= range.c1; col += 1) wb.setBlank({ sheet: range.sheet, row, col });
};

/** Write the column names of the table at `index` into the header row of `range`. */
const writeHeaderNames = (wb: WorkbookHandle, index: number, range: Range): void => {
  const columns = wb.getTableStyleDetails(index)?.columns ?? [];
  columns.forEach((c, i) => {
    wb.setText({ sheet: range.sheet, row: range.r0, col: range.c0 + i }, c.name);
  });
};

/** Totals settings of the table at `index` by lower-cased column name. */
const captureTotals = (wb: WorkbookHandle, index: number) => {
  const out = new Map<string, { function: string; label: string; formula: string }>();
  for (const c of wb.getTableStyleDetails(index)?.columns ?? []) {
    if (c.totalsRowFunction === 'none' && !c.totalsRowLabel) continue;
    out.set(nameKey(c.name), {
      function: c.totalsRowFunction,
      label: c.totalsRowLabel,
      formula: c.totalsRowFormula,
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
  columns.forEach((c, i) => {
    const saved = totals.get(nameKey(c.name));
    if (!saved) return;
    wb.setTableColumnTotals(index, i, {
      function: saved.function as (typeof columns)[number]['totalsRowFunction'],
      label: saved.label,
      ...(saved.formula ? { formula: saved.formula } : {}),
    });
  });
};

const removeEngineTable = (wb: WorkbookHandle, index: number): boolean => {
  if (!wb.removeTable(index)) return false;
  forgetTableAutoFilterImport(wb, index);
  return true;
};

const tableUpdate = (before: TableOverlay, after: TableOverlay): TableUpdateInput => {
  const patch: TableUpdateInput = {};
  if (!sameRange(before.range, after.range)) {
    patch.ref = formatA1Range(after.range, { collapse: false });
  }
  if (before.name !== after.name) patch.name = after.name;
  if (overlayStyleName(before) !== overlayStyleName(after)) {
    patch.styleName = overlayStyleName(after);
  }
  if (before.banded !== after.banded) patch.showRowStripes = after.banded;
  if (!!before.firstCol !== !!after.firstCol) patch.showFirstColumn = !!after.firstCol;
  if (!!before.lastCol !== !!after.lastCol) patch.showLastColumn = !!after.lastCol;
  if (!!before.colStripes !== !!after.colStripes) patch.showColumnStripes = !!after.colStripes;
  if (before.showHeader !== after.showHeader) patch.headerRow = after.showHeader;
  if (before.showTotal !== after.showTotal) patch.totalsRow = after.showTotal;
  return patch;
};

const sameRangeList = (a: readonly Range[], b: readonly Range[]): boolean =>
  a.length === b.length && a.every((range, i) => sameRange(range, b[i] as Range));

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
    edits.rejected.size > 0 || edits.totals.size > 0
      ? state.tables.tables.flatMap((table) => {
          if (edits.rejected.has(table.id)) {
            const restored = edits.rejected.get(table.id);
            return restored ? [restored] : [];
          }
          const totals = edits.totals.get(table.id);
          return [totals ? { ...table, totals } : table];
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
      return this.apply(store, previousSlice, slice);
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
    const edits: FilterEdits = { dropped: [], added: [], totals: new Map(), rejected: new Map() };
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
        const patch = tableUpdate(old, table);
        const totalsChanged = !sameTotals(old.totals, table.totals);
        if (Object.keys(patch).length === 0 && !totalsChanged) continue;
        from = tableFilterRangeIn(store.getState(), old.range);
        const reshaped =
          old.range.sheet !== table.range.sheet ||
          old.range.c1 - old.range.c0 !== table.range.c1 - table.range.c0;
        if (!reshaped) {
          if (Object.keys(patch).length > 0 && !wb.updateTable(existing.index, patch)) {
            edits.rejected.set(table.id, old);
            continue;
          }
          if (patch.totalsRow === false && old.showTotal) clearRow(wb, old.range, old.range.r1);
          if (patch.headerRow === false && old.showHeader) clearRow(wb, old.range, old.range.r0);
          if (patch.headerRow === true) writeHeaderNames(wb, existing.index, table.range);
          if (patch.headerRow !== undefined) cellsChanged = true;
          const wroteTotals = totalsChanged && writeTotals(wb, existing.index, table.totals);
          if (patch.totalsRow === true) {
            applyDefaultTotals(wb, existing.index, table);
            cellsChanged = true;
          } else if (patch.totalsRow === false || (wroteTotals && table.showTotal)) {
            cellsChanged = true;
          }
          this.adoptTotals(edits, table, existing.index);
          // Rewritten structured references hold no value until the next recalc.
          if (patch.name !== undefined) cellsChanged = true;
          const filterMoved =
            patch.ref !== undefined ||
            patch.headerRow !== undefined ||
            patch.totalsRow !== undefined;
          if (!filterMoved) continue;
        } else {
          // A new column count or sheet needs a new table; when the engine
          // refuses it, the old table comes back.
          const totals = captureTotals(wb, existing.index);
          removeEngineTable(wb, existing.index);
          const displayName = renamedOld ? name : existing.displayName;
          index = createEngineTable(wb, { ...table, name }, displayName);
          if (index < 0) {
            const oldName = old.name as string;
            const restored =
              createEngineTable(wb, { ...old, name: oldName }, existing.displayName) >= 0;
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
      const range = tableFilterRange(wb, index, table.range.sheet);
      if (range) edits.added.push({ range, from });
    }
    if (cellsChanged) this.refreshCells(store);
    if (
      edits.dropped.length === 0 &&
      edits.added.length === 0 &&
      edits.totals.size === 0 &&
      edits.rejected.size === 0
    ) {
      return false;
    }
    let wrote = false;
    store.setState((state) => {
      const next = applyFilterEdits(state, edits);
      wrote = next !== state;
      // Restored or rejected overlays already match the engine.
      this.last = next.tables;
      return next;
    });
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

/** The store filter range belonging to the table spanning `table`, if any. */
function tableFilterRangeIn(state: State, table: Range): Range | null {
  return state.ui.filterRanges.find((range) => isFilterOfTable(table, range)) ?? null;
}
