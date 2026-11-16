import type {
  AutoFilter,
  AutoFilterEval,
  AutoFilterInput,
  Table,
  TableStyleOptionsPatch,
} from '@libraz/formulon';
import { formatA1Range, rangeOfRect } from './address.js';
import { parseRangeRef as parseTableRef } from './range-resolver.js';
import { attemptWrite, orFallback } from './result.js';
import type {
  Addr,
  CellValue,
  EngineCapabilities,
  TableInput,
  TablePatch,
  Workbook,
} from './types.js';
import { formatCell } from './value.js';
import type { WorkbookHandle } from './workbook-handle.js';
import { assertAlive, wb, ws } from './workbook-handle-internals.js';

type WorkbookHandleCtor = { prototype: WorkbookHandle };

/** Every table of the workbook: sheet order, then each sheet's table order. A
 *  table's position here is the workbook-wide index the handle methods take. */
export function workbookTables(wb: Workbook): Table[] {
  const out: Table[] = [];
  for (let sheet = 0; sheet < wb.sheetCount; sheet += 1) out.push(...wb.sheet(sheet).listTables());
  return out;
}

/** The table at workbook-wide `index` (see {@link workbookTables}), or null. */
export function workbookTableAt(wb: Workbook, index: number): Table | null {
  if (index < 0) return null;
  return workbookTables(wb)[index] ?? null;
}

declare module './workbook-handle.js' {
  interface WorkbookHandle extends WorkbookHandleTablesMethods {}
}

function tableAt(handle: unknown, index: number): Table | null {
  return workbookTableAt(wb(handle), index);
}

/** Derive column display names from the header row of a table's `ref`.
 *  Returns labels in source order; cells that read blank fall back to the
 *  Spreadsheet-style `Column1` / `Column2` placeholder so the structured-ref
 *  autocomplete still has something to insert. */
function tableColumnNames(
  handle: WorkbookHandleTablesMethods,
  sheet: number,
  ref: string,
): string[] {
  const parsed = parseTableRef(ref);
  if (!parsed) return [];
  const out: string[] = [];
  for (let col = parsed.c0; col <= parsed.c1; col += 1) {
    const v = handle.getValue({ sheet, row: parsed.r0, col });
    const text = formatCell(v);
    out.push(text || `Column${out.length + 1}`);
  }
  return out;
}

export abstract class WorkbookHandleTablesMethods {
  declare readonly capabilities: EngineCapabilities;
  abstract getValue(addr: Addr): CellValue;

  /** Returns the complete worksheet `<autoFilter>` fragment (`''` when the
   * sheet has none), or null when the read fails. */
  getSheetAutoFilterXml(sheet: number): string | null {
    assertAlive(this);
    if (!this.capabilities.autoFilter) return null;
    return orFallback(null, () => ws(this, sheet).ooxml.getAutoFilterXml() ?? '');
  }

  /** Replaces the worksheet `<autoFilter>` fragment. Empty XML removes the
   * definition. Returns false for failed writes. */
  setSheetAutoFilterXml(sheet: number, xml: string): boolean {
    assertAlive(this);
    if (!this.capabilities.autoFilter) return false;
    return attemptWrite(() => ws(this, sheet).ooxml.setAutoFilterXml(xml.length > 0 ? xml : null));
  }

  /** Read the typed worksheet AutoFilter. `null` means that the sheet has no
   * definition; `null` is also returned when the read itself failed. */
  getAutoFilter(sheet: number): AutoFilter | null {
    assertAlive(this);
    if (!this.capabilities.typedAutoFilter) return null;
    return orFallback(null, () => ws(this, sheet).getAutoFilter());
  }

  /** Replace the typed worksheet AutoFilter definition. This does not change
   * row visibility; callers that want Excel's Reapply behaviour should call
   * {@link applyAutoFilter} after this method. */
  setAutoFilter(sheet: number, filter: AutoFilterInput): boolean {
    assertAlive(this);
    if (!this.capabilities.typedAutoFilter) return false;
    return attemptWrite(() => ws(this, sheet).setAutoFilter(filter));
  }

  /** Remove a worksheet AutoFilter definition. Row visibility is unchanged by
   * the engine, so a caller clearing the UI filter should clear it first. */
  removeAutoFilter(sheet: number): boolean {
    assertAlive(this);
    if (!this.capabilities.typedAutoFilter) return false;
    return attemptWrite(() => ws(this, sheet).setAutoFilter(null));
  }

  /** Apply the worksheet AutoFilter and update native row visibility. */
  applyAutoFilter(sheet: number): boolean {
    assertAlive(this);
    if (!this.capabilities.typedAutoFilter) return false;
    return attemptWrite(() => ws(this, sheet).applyAutoFilter());
  }

  /** Remove all worksheet AutoFilter criteria and show every body row. */
  clearAutoFilter(sheet: number): boolean {
    assertAlive(this);
    if (!this.capabilities.typedAutoFilter) return false;
    return attemptWrite(() => ws(this, sheet).clearAutoFilterCriteria());
  }

  /** Evaluate the worksheet AutoFilter without changing native visibility. */
  evaluateAutoFilter(sheet: number): AutoFilterEval | null {
    assertAlive(this);
    if (!this.capabilities.typedAutoFilter) return null;
    return orFallback(null, () => ws(this, sheet).evaluateAutoFilter());
  }

  /** Read a table's own typed AutoFilter by workbook-wide table index (the
   * `index` of {@link getTables}). `null` means the table has none. */
  getTableAutoFilter(table: number): AutoFilter | null {
    assertAlive(this);
    if (!this.capabilities.typedTableAutoFilter) return null;
    const t = tableAt(this, table);
    return t ? orFallback(null, () => t.getAutoFilter()) : null;
  }

  /** Replace a table's AutoFilter. Row visibility is unchanged; follow with
   * {@link applyTableAutoFilter} to hide the excluded rows. */
  setTableAutoFilter(table: number, filter: AutoFilterInput): boolean {
    assertAlive(this);
    if (!this.capabilities.typedTableAutoFilter) return false;
    const t = tableAt(this, table);
    return t ? attemptWrite(() => t.setAutoFilter(filter)) : false;
  }

  /** Remove a table's AutoFilter (its header buttons). Row visibility is unchanged. */
  removeTableAutoFilter(table: number): boolean {
    assertAlive(this);
    if (!this.capabilities.typedTableAutoFilter) return false;
    const t = tableAt(this, table);
    return t ? attemptWrite(() => t.setAutoFilter(null)) : false;
  }

  /** Apply a table's AutoFilter and update native row visibility. */
  applyTableAutoFilter(table: number): boolean {
    assertAlive(this);
    if (!this.capabilities.typedTableAutoFilter) return false;
    const t = tableAt(this, table);
    return t ? attemptWrite(() => t.applyAutoFilter()) : false;
  }

  /** Drop every criterion of a table's AutoFilter and show all its body rows. */
  clearTableAutoFilter(table: number): boolean {
    assertAlive(this);
    if (!this.capabilities.typedTableAutoFilter) return false;
    const t = tableAt(this, table);
    return t ? attemptWrite(() => t.clearAutoFilterCriteria()) : false;
  }

  /** Evaluate a table's AutoFilter without changing native visibility. */
  evaluateTableAutoFilter(table: number): AutoFilterEval | null {
    assertAlive(this);
    if (!this.capabilities.typedTableAutoFilter) return null;
    const t = tableAt(this, table);
    return t ? orFallback(null, () => t.evaluateAutoFilter()) : null;
  }

  /** Creates an OOXML worksheet table on `sheet` and returns its workbook-wide
   * index, or -1 when the engine refuses it. */
  createTable(sheet: number, input: TableInput): number {
    assertAlive(this);
    if (!this.capabilities.tableMutate) return -1;
    const created = orFallback(null, () => ws(this, sheet).addTable(input));
    return created ? workbookTables(wb(this)).indexOf(created) : -1;
  }

  /** Partially update the table at `index`; omitted fields keep their value.
   * A new `range` must keep the table's column count. */
  updateTable(index: number, patch: TablePatch): boolean {
    assertAlive(this);
    if (!this.capabilities.tableMutate) return false;
    const t = tableAt(this, index);
    return t ? attemptWrite(() => t.update(patch)) : false;
  }

  /** Insert data rows `first`..`last` (data-row positions) into the table at `index`; the
   *  table's columns move down and calculated columns fill the new rows. */
  insertTableRows(index: number, first: number, last: number): boolean {
    assertAlive(this);
    if (!this.capabilities.tableMutate) return false;
    const t = tableAt(this, index);
    return t ? attemptWrite(() => t.insertRows(first, last)) : false;
  }

  /** Insert columns `first`..`last` (table-column positions) into the table at `index`; the
   *  table's rows move right and each new column is named from its left neighbour. */
  insertTableColumns(index: number, first: number, last: number): boolean {
    assertAlive(this);
    if (!this.capabilities.tableMutate) return false;
    const t = tableAt(this, index);
    return t ? attemptWrite(() => t.insertColumns(first, last)) : false;
  }

  /** Rename the table at `index`; structured references follow the new name. */
  renameTable(index: number, name: string): boolean {
    assertAlive(this);
    if (!this.capabilities.tableMutate) return false;
    const t = tableAt(this, index);
    return t ? attemptWrite(() => t.rename(name)) : false;
  }

  /** Update the style name and show-flags of the table at `index`; omitted
   * keys keep their value and a `null` style name removes the name. */
  updateTableStyleOptions(index: number, patch: TableStyleOptionsPatch): boolean {
    assertAlive(this);
    if (!this.capabilities.tableMutate) return false;
    const t = tableAt(this, index);
    return t ? attemptWrite(() => t.updateStyleOptions(patch)) : false;
  }

  /** Remove the table at `index`; later tables shift down by one. */
  removeTable(index: number): boolean {
    assertAlive(this);
    if (!this.capabilities.tableMutate) return false;
    const t = tableAt(this, index);
    return t ? attemptWrite(() => t.sheet.removeTable(t)) : false;
  }

  /** Workbook-wide index of the table named `name` under the engine's name
   * identity rule, or null when no table has it. */
  tableIndexByName(name: string): number | null {
    assertAlive(this);
    if (!this.capabilities.tableMutate) return null;
    const w = wb(this);
    const table = orFallback(null, () => w.table(name));
    if (!table) return null;
    const index = workbookTables(w).indexOf(table);
    return index < 0 ? null : index;
  }

  /** Snapshot of every spreadsheet Table on the workbook, in index order.
   *  Empty array on the stub. */
  getTables(): {
    /** Workbook-wide table index, as taken by the table AutoFilter methods. */
    index: number;
    name: string;
    displayName: string;
    ref: string;
    sheetIndex: number;
    columns: string[];
  }[] {
    assertAlive(this);
    if (!this.capabilities.tableMutate) return [];
    return (orFallback(null, () => workbookTables(wb(this))) ?? []).map((table, index) => {
      const { name, range } = table.getProperties();
      const ref = formatA1Range(rangeOfRect(range.sheet, range.rect), { collapse: false });
      return {
        index,
        name,
        displayName: name,
        ref,
        sheetIndex: range.sheet,
        columns: tableColumnNames(this, range.sheet, ref),
      };
    });
  }
}

export function installTablesMethods(target: WorkbookHandleCtor): void {
  for (const key of Object.getOwnPropertyNames(WorkbookHandleTablesMethods.prototype)) {
    if (key === 'constructor') continue;
    const descriptor = Object.getOwnPropertyDescriptor(WorkbookHandleTablesMethods.prototype, key);
    if (!descriptor) continue;
    Object.defineProperty(target.prototype, key, descriptor);
  }
}
