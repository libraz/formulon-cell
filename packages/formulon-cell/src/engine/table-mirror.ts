import { withFilterRange } from '../commands/filter-ranges.js';
import { type TableOverlay, tableColumns, tableStyleName } from '../commands/format-as-table.js';
import { sameRange } from '../store/selection-geometry.js';
import type { SpreadsheetStore, State, ValueFilterCriteria } from '../store/store.js';
import { formatA1Range } from './address.js';
import { forgetTableAutoFilterImport } from './auto-filter-sync.js';
import type { Range, TableUpdateInput } from './types.js';
import type { WorkbookHandle } from './workbook-handle.js';

/**
 * Reconciles the store's engine-backed table overlays (those carrying a
 * `name`) with the engine's worksheet tables, so every table edit, removal and
 * undo reaches the saved file. Overlays are matched to engine tables by name;
 * indexes are re-resolved for every write because a removal shifts them.
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
    styleName: tableStyleName(table.style),
    headerRow: table.showHeader,
    totalsRow: table.showTotal,
  });

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
  if (before.style !== after.style) patch.styleName = tableStyleName(after.style);
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
    edits.rejected.size > 0
      ? state.tables.tables.flatMap((table) => {
          if (!edits.rejected.has(table.id)) return [table];
          const restored = edits.rejected.get(table.id);
          return restored ? [restored] : [];
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
  constructor(private readonly handle: WorkbookHandle) {}

  /** Overlays the engine already carries. */
  private last: readonly TableOverlay[] | null = null;

  /** Adopt `tables` as the engine baseline without writing it. */
  rebase(tables: readonly TableOverlay[] | null): void {
    this.last = tables;
  }

  /** Write the store's table changes to the engine. Returns true when the
   *  filter follow-up notified the store, which then mirrored every slice. */
  sync(store: SpreadsheetStore): boolean {
    // Read the store, not a listener argument: another listener's nested
    // write can deliver an older state after a newer one.
    const tables = store.getState().tables.tables;
    const previous = this.last;
    if (tables === previous) return false;
    this.last = tables;
    const wb = this.handle;
    if (!wb.capabilities.tableMutate) return false;
    const before = namedOverlays(previous ?? []);
    const after = namedOverlays(tables);
    const edits: FilterEdits = { dropped: [], added: [], rejected: new Map() };
    // Removals first, so a table re-created over the same cells cannot overlap.
    for (const [key, table] of before) {
      if (after.has(key) || !table.name) continue;
      const existing = engineTable(wb, table.name);
      if (existing) removeEngineTable(wb, existing.index);
      edits.dropped.push(table.range);
    }
    for (const [key, table] of after) {
      const name = table.name as string;
      const old = before.get(key) ?? null;
      const existing = engineTable(wb, name);
      if (existing && !old) continue;
      let index = existing?.index ?? -1;
      let from: Range | null = null;
      if (existing && old) {
        const patch = tableUpdate(old, table);
        if (Object.keys(patch).length === 0) continue;
        from = tableFilterRangeIn(store.getState(), old.range);
        const reshaped =
          old.range.sheet !== table.range.sheet ||
          old.range.c1 - old.range.c0 !== table.range.c1 - table.range.c0;
        if (!reshaped) {
          if (!wb.updateTable(existing.index, patch)) {
            edits.rejected.set(table.id, old);
            continue;
          }
          const filterMoved =
            patch.ref !== undefined ||
            patch.headerRow !== undefined ||
            patch.totalsRow !== undefined;
          if (!filterMoved) continue;
        } else {
          // A new column count or sheet needs a new table under the same name;
          // when the engine refuses it, the old table comes back.
          removeEngineTable(wb, existing.index);
          index = createEngineTable(wb, { ...table, name }, existing.displayName);
          if (index < 0) {
            const restored = createEngineTable(wb, { ...old, name }, existing.displayName) >= 0;
            if (!restored) edits.dropped.push(old.range);
            edits.rejected.set(table.id, restored ? old : null);
            continue;
          }
        }
      } else {
        index = createEngineTable(wb, { ...table, name });
      }
      if (old) edits.dropped.push(old.range);
      if (index < 0) {
        edits.rejected.set(table.id, null);
        continue;
      }
      const range = tableFilterRange(wb, index, table.range.sheet);
      if (range) edits.added.push({ range, from });
    }
    if (edits.dropped.length === 0 && edits.added.length === 0 && edits.rejected.size === 0) {
      return false;
    }
    let wrote = false;
    store.setState((state) => {
      const next = applyFilterEdits(state, edits);
      wrote = next !== state;
      // Restored or rejected overlays already match the engine.
      this.last = next.tables.tables;
      return next;
    });
    return wrote;
  }
}

/** The store filter range belonging to the table spanning `table`, if any. */
function tableFilterRangeIn(state: State, table: Range): Range | null {
  return state.ui.filterRanges.find((range) => isFilterOfTable(table, range)) ?? null;
}
