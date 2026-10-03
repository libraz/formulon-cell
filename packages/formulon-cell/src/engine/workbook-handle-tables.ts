import { parseRangeRef as parseTableRef } from './range-resolver.js';
import type { Addr, CellValue, EngineCapabilities, TableInput, Workbook } from './types.js';
import { formatCell } from './value.js';
import type { WorkbookHandle } from './workbook-handle.js';

type WorkbookHandleCtor = { prototype: WorkbookHandle };
type WorkbookHandleInternals = {
  wb: Workbook;
  assertAlive(): void;
};

type TableAuthoringWorkbook = Workbook & {
  createTable?: (input: TableInput) => { status: { ok: boolean }; index: number };
  updateTable?: (
    index: number,
    input: Pick<TableInput, 'ref' | 'styleName' | 'headerRow' | 'totalsRow'>,
  ) => { ok: boolean };
  removeTable?: (index: number) => { ok: boolean };
};

type AutoFilterWorkbook = Workbook & {
  getSheetAutoFilterXml?: (sheet: number) => { status: { ok: boolean }; xml: string };
  setSheetAutoFilterXml?: (sheet: number, xml: string) => { ok: boolean };
};

declare module './workbook-handle.js' {
  interface WorkbookHandle extends WorkbookHandleTablesMethods {}
}

function internals(handle: unknown): WorkbookHandleInternals {
  return handle as WorkbookHandleInternals;
}

function assertAlive(handle: unknown): void {
  internals(handle).assertAlive();
}

function wb(handle: unknown): Workbook {
  return internals(handle).wb;
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

  /** Returns the complete worksheet `<autoFilter>` fragment, or null when
   * the engine does not expose the optional raw-definition seam. */
  getSheetAutoFilterXml(sheet: number): string | null {
    assertAlive(this);
    if (!this.capabilities.autoFilter) return null;
    const r = (wb(this) as AutoFilterWorkbook).getSheetAutoFilterXml?.(sheet);
    return r?.status.ok ? r.xml : null;
  }

  /** Replaces the worksheet `<autoFilter>` fragment. Empty XML removes the
   * definition. Returns false for old bundles and failed writes. */
  setSheetAutoFilterXml(sheet: number, xml: string): boolean {
    assertAlive(this);
    if (!this.capabilities.autoFilter) return false;
    return (wb(this) as AutoFilterWorkbook).setSheetAutoFilterXml?.(sheet, xml)?.ok === true;
  }

  /** Creates an OOXML worksheet table and returns its index, or -1 when the
   * loaded engine predates table authoring support. */
  createTable(input: TableInput): number {
    assertAlive(this);
    if (!this.capabilities.tableMutate) return -1;
    const result = (wb(this) as TableAuthoringWorkbook).createTable?.(input);
    return result?.status.ok ? result.index : -1;
  }

  updateTable(
    index: number,
    input: Pick<TableInput, 'ref' | 'styleName' | 'headerRow' | 'totalsRow'>,
  ): boolean {
    assertAlive(this);
    return (
      this.capabilities.tableMutate === true &&
      (wb(this) as TableAuthoringWorkbook).updateTable?.(index, input).ok === true
    );
  }

  removeTable(index: number): boolean {
    assertAlive(this);
    return (
      this.capabilities.tableMutate === true &&
      (wb(this) as TableAuthoringWorkbook).removeTable?.(index).ok === true
    );
  }

  /** Snapshot of every spreadsheet Table on the workbook. Read-only in the engine —
   *  we surface it as a badge count + listing for the status bar. Empty array
   *  on the stub. */
  getTables(): {
    name: string;
    displayName: string;
    ref: string;
    sheetIndex: number;
    columns: string[];
  }[] {
    assertAlive(this);
    if (!wb(this).tableCount) return [];
    const count = wb(this).tableCount();
    if (!count.status.ok) return [];
    const n = count.value;
    const out: {
      name: string;
      displayName: string;
      ref: string;
      sheetIndex: number;
      columns: string[];
    }[] = [];
    for (let i = 0; i < n; i += 1) {
      const e = wb(this).tableAt(i);
      if (!e.status.ok || !e.name || !e.displayName || !e.ref || e.sheetIndex === undefined)
        continue;
      out.push({
        name: e.name,
        displayName: e.displayName,
        ref: e.ref,
        sheetIndex: e.sheetIndex,
        columns: tableColumnNames(this, e.sheetIndex, e.ref),
      });
    }
    return out;
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
