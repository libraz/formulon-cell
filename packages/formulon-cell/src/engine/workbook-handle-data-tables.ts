import type { DataTable } from '@libraz/formulon';
import type { Addr, EngineCapabilities, Range, Workbook } from './types.js';
import type { WorkbookHandle } from './workbook-handle.js';

type WorkbookHandleCtor = { prototype: WorkbookHandle };
type WorkbookHandleInternals = {
  wb: Workbook;
  assertAlive(): void;
  emitRecalc(): void;
};

declare module './workbook-handle.js' {
  interface WorkbookHandle extends WorkbookHandleDataTablesMethods {}
}

/** A What-If data table. `range` is the result block (the `TABLE()` cells);
 *  the formula row sits above it and the substitution values left of it. */
export interface DataTableInfo {
  range: Range;
  rowInput: Addr | null;
  colInput: Addr | null;
  /** An input cell was deleted, so the table evaluates to `#REF!`. */
  inputDeleted: boolean;
}

/** Argument of `addDataTable`: the result block and at least one input cell. */
export interface DataTableDefinition {
  range: Range;
  rowInput: Addr | null;
  colInput: Addr | null;
}

const TABLE_FORMULA = /^=?\s*TABLE\(/i;

/** Whether `formula` is the array formula the engine writes into a data table's result cells. */
export const isDataTableFormula = (formula: string | null | undefined): boolean =>
  !!formula && TABLE_FORMULA.test(formula);

/** Whether `formula` calls TABLE() anywhere outside a string literal. Only
 *  the Data Table command may create that call. */
export const formulaCallsTable = (formula: string | null | undefined): boolean => {
  if (!formula) return false;
  const code = formula.replace(/"(?:[^"]|"")*"/g, '""');
  return /(?:^|[^A-Za-z0-9_.])TABLE\s*\(/i.test(code);
};

function internals(handle: unknown): WorkbookHandleInternals {
  return handle as WorkbookHandleInternals;
}

function toInfo(sheet: number, table: DataTable): DataTableInfo {
  const input = (i: DataTable['rowInput']): Addr | null =>
    i ? { sheet, row: i.row, col: i.col } : null;
  return {
    range: { sheet, r0: table.firstRow, c0: table.firstCol, r1: table.lastRow, c1: table.lastCol },
    rowInput: input(table.rowInput),
    colInput: input(table.colInput),
    inputDeleted: !!table.rowInput?.deleted || !!table.colInput?.deleted,
  };
}

/** A new table reads 0 until computed; Manual leaves that to Calculate Now. */
function computeUnlessManual(handle: unknown): void {
  const self = internals(handle);
  if ((handle as WorkbookHandle).calcMode() !== 1) {
    const s = self.wb.recalcDataTables();
    if (!s.ok) throw new Error(`recalcDataTables: ${s.message}`);
  }
  self.emitRecalc();
}

export abstract class WorkbookHandleDataTablesMethods {
  declare readonly capabilities: EngineCapabilities;

  /** The data table whose result block holds `addr`, or `null`. */
  dataTableAt(addr: Addr): DataTableInfo | null {
    const self = internals(this);
    self.assertAlive();
    if (!this.capabilities.dataTables) return null;
    const r = self.wb.getDataTable(addr.sheet, addr.row, addr.col);
    return r.status.ok && r.dataTable ? toInfo(addr.sheet, r.dataTable) : null;
  }

  /** Data tables whose result block intersects `range`. Every result cell
   *  carries the `TABLE()` formula, so the populated cells find them. */
  dataTablesInRange(range: Range): DataTableInfo[] {
    internals(this).assertAlive();
    if (!this.capabilities.dataTables) return [];
    const found = new Map<string, DataTableInfo>();
    const handle = this as unknown as WorkbookHandle;
    const inFound = ({ row, col }: Addr): boolean => {
      for (const { range: r } of found.values()) {
        if (row >= r.r0 && row <= r.r1 && col >= r.c0 && col <= r.c1) return true;
      }
      return false;
    };
    for (const cell of handle.cellsInRange(range)) {
      if (!isDataTableFormula(cell.formula) || inFound(cell.addr)) continue;
      const table = this.dataTableAt(cell.addr);
      if (!table) continue;
      const { r0, c0 } = table.range;
      found.set(`${r0}:${c0}`, table);
    }
    return [...found.values()];
  }

  /** Every data table on `sheet`. */
  dataTables(sheet: number): DataTableInfo[] {
    return this.dataTablesInRange({ sheet, r0: 0, c0: 0, r1: 1048575, c1: 16383 });
  }

  /** Creates a data table and computes it unless calculation is Manual.
   *  Returns `false` when the engine rejects the definition: no input, an
   *  input inside the block, no formula row/column, or an overlap. */
  addDataTable(def: DataTableDefinition): boolean {
    const self = internals(this);
    self.assertAlive();
    if (!this.capabilities.dataTables) return false;
    const { range } = def;
    const s = self.wb.addDataTable(range.sheet, {
      firstRow: range.r0,
      lastRow: range.r1,
      firstCol: range.c0,
      lastCol: range.c1,
      rowInput: def.rowInput ? { row: def.rowInput.row, col: def.rowInput.col } : null,
      colInput: def.colInput ? { row: def.colInput.row, col: def.colInput.col } : null,
    });
    if (!s.ok) return false;
    computeUnlessManual(this);
    return true;
  }

  /** Removes the data table covering `addr` and blanks its result cells. */
  removeDataTable(addr: Addr): boolean {
    const self = internals(this);
    self.assertAlive();
    if (!this.capabilities.dataTables) return false;
    if (!self.wb.removeDataTable(addr.sheet, addr.row, addr.col).ok) return false;
    self.emitRecalc();
    return true;
  }
}

export function installDataTablesMethods(target: WorkbookHandleCtor): void {
  for (const key of Object.getOwnPropertyNames(WorkbookHandleDataTablesMethods.prototype)) {
    if (key === 'constructor') continue;
    const descriptor = Object.getOwnPropertyDescriptor(
      WorkbookHandleDataTablesMethods.prototype,
      key,
    );
    if (!descriptor) continue;
    Object.defineProperty(target.prototype, key, descriptor);
  }
}
