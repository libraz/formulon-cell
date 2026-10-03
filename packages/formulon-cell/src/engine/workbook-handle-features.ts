import type {
  Addr,
  CellValue,
  EngineCapabilities,
  FormulonModule,
  FunctionMetadataProvider,
  TableInput,
  Workbook,
} from './types.js';
import type { WorkbookHandle } from './workbook-handle.js';

type WorkbookHandleCtor = { prototype: WorkbookHandle };
type WorkbookHandleInternals = {
  wb: Workbook;
  module: FormulonModule;
  capabilities: EngineCapabilities;
  functionMetadataProvider: FunctionMetadataProvider | null;
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
  interface WorkbookHandle extends WorkbookHandleFeatureMethods {}
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

export abstract class WorkbookHandleFeatureMethods {
  declare readonly capabilities: EngineCapabilities;
  declare readonly sheetCount: number;
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
}

export function installWorkbookFeatureMethods(target: WorkbookHandleCtor): void {
  for (const key of Object.getOwnPropertyNames(WorkbookHandleFeatureMethods.prototype)) {
    if (key === 'constructor') continue;
    const descriptor = Object.getOwnPropertyDescriptor(WorkbookHandleFeatureMethods.prototype, key);
    if (!descriptor) continue;
    Object.defineProperty(target.prototype, key, descriptor);
  }
}
