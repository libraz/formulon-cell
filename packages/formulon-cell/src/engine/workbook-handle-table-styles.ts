import type {
  PivotStyleInfo,
  PivotStyleInfoResult,
  TableColumnEntry,
  TableColumnTotalsInput,
  TableEntry,
  TableStyle,
  TableStyleDefaultsResult,
  TableStyleResult,
} from '@libraz/formulon';
import type { EngineCapabilities, Workbook } from './types.js';
import type { WorkbookHandle } from './workbook-handle.js';

type WorkbookHandleCtor = { prototype: WorkbookHandle };
type WorkbookHandleInternals = {
  wb: Workbook;
  assertAlive(): void;
};

type TableStyleWorkbook = Workbook & {
  tableAt?: (idx: number) => TableEntry;
  setTableColumnTotals?: (
    idx: number,
    column: number,
    totals: TableColumnTotalsInput,
  ) => { ok: boolean };
  tableStyleCount?: () => { status: { ok: boolean }; value: number };
  getTableStyle?: (index: number) => TableStyleResult;
  setTableStyle?: (style: TableStyle) => { ok: boolean };
  removeTableStyle?: (name: string) => { ok: boolean };
  getTableStyleDefaults?: () => TableStyleDefaultsResult;
  setTableStyleDefaults?: (defaults: { tableStyle?: string; pivotStyle?: string }) => {
    ok: boolean;
  };
  pivotGetStyleInfo?: (sheet: number, pivotIdx: number) => PivotStyleInfoResult;
  pivotSetStyleInfo?: (sheet: number, pivotIdx: number, info: PivotStyleInfo) => { ok: boolean };
};

declare module './workbook-handle.js' {
  interface WorkbookHandle extends WorkbookHandleTableStylesMethods {}
}

/** A table's style, show-flags, row counts and per-column totals as the engine stores them. */
export interface TableStyleDetails {
  styleName: string;
  hasStyleInfo: boolean;
  showFirstColumn: boolean;
  showLastColumn: boolean;
  showRowStripes: boolean;
  showColumnStripes: boolean;
  headerRowCount: number;
  totalsRowCount: number;
  columns: TableColumnEntry[];
}

/** A custom table / PivotTable style as stored in `<tableStyles>`. */
export type CustomEngineTableStyle = Required<Pick<TableStyle, 'name' | 'pivot' | 'table'>> &
  Required<Pick<TableStyle, 'elements'>>;

function internals(handle: unknown): WorkbookHandleInternals {
  return handle as WorkbookHandleInternals;
}

function assertAlive(handle: unknown): void {
  internals(handle).assertAlive();
}

function wb(handle: unknown): TableStyleWorkbook {
  return internals(handle).wb as TableStyleWorkbook;
}

export abstract class WorkbookHandleTableStylesMethods {
  declare readonly capabilities: EngineCapabilities;

  /** Style, show-flags and totals of the table at `index`; null when the
   *  engine predates them or `index` is out of range. */
  getTableStyleDetails(index: number): TableStyleDetails | null {
    assertAlive(this);
    if (!this.capabilities.tableStyleInfo) return null;
    const e = wb(this).tableAt?.(index);
    if (!e?.status.ok) return null;
    return {
      styleName: e.styleName ?? '',
      hasStyleInfo: e.hasStyleInfo === true,
      showFirstColumn: e.showFirstColumn === true,
      showLastColumn: e.showLastColumn === true,
      showRowStripes: e.showRowStripes === true,
      showColumnStripes: e.showColumnStripes === true,
      headerRowCount: e.headerRowCount ?? 0,
      totalsRowCount: e.totalsRowCount ?? 0,
      columns: e.columns ?? [],
    };
  }

  /** Set one table column's totals-row function and label. While the totals
   *  row is shown the engine writes the column's totals cell to match. */
  setTableColumnTotals(index: number, column: number, totals: TableColumnTotalsInput): boolean {
    assertAlive(this);
    if (!this.capabilities.tableStyleInfo) return false;
    return wb(this).setTableColumnTotals?.(index, column, totals).ok === true;
  }

  /** Every custom table / PivotTable style, in file order. */
  getCustomTableStyles(): CustomEngineTableStyle[] {
    assertAlive(this);
    if (!this.capabilities.customTableStyles) return [];
    const count = wb(this).tableStyleCount?.();
    if (!count?.status.ok) return [];
    const out: CustomEngineTableStyle[] = [];
    for (let i = 0; i < count.value; i += 1) {
      const r = wb(this).getTableStyle?.(i);
      if (!r?.status.ok) continue;
      out.push({ name: r.name, pivot: r.pivot, table: r.table, elements: r.elements });
    }
    return out;
  }

  /** Add or replace a custom table style. False when the engine refuses it
   *  (built-in name, unknown differential-format id). */
  setCustomTableStyle(style: TableStyle): boolean {
    assertAlive(this);
    if (!this.capabilities.customTableStyles) return false;
    return wb(this).setTableStyle?.(style).ok === true;
  }

  /** Remove the custom table style `name`. */
  removeCustomTableStyle(name: string): boolean {
    assertAlive(this);
    if (!this.capabilities.customTableStyles) return false;
    return wb(this).removeTableStyle?.(name).ok === true;
  }

  /** The workbook's default table and PivotTable style names. */
  getTableStyleDefaults(): { tableStyle: string; pivotStyle: string } | null {
    assertAlive(this);
    if (!this.capabilities.customTableStyles) return null;
    const r = wb(this).getTableStyleDefaults?.();
    return r?.status.ok ? { tableStyle: r.tableStyle, pivotStyle: r.pivotStyle } : null;
  }

  /** Set the default style names; an empty string removes one. */
  setTableStyleDefaults(defaults: { tableStyle?: string; pivotStyle?: string }): boolean {
    assertAlive(this);
    if (!this.capabilities.customTableStyles) return false;
    return wb(this).setTableStyleDefaults?.(defaults).ok === true;
  }

  /** Style name and toggles of a PivotTable; null when it has no style element. */
  getPivotStyleInfo(sheet: number, pivotIndex: number): Required<PivotStyleInfo> | null {
    assertAlive(this);
    if (!this.capabilities.pivotStyleInfo) return null;
    const r = wb(this).pivotGetStyleInfo?.(sheet, pivotIndex);
    if (!r?.status.ok || !r.present) return null;
    return {
      name: r.name,
      showRowHeaders: r.showRowHeaders,
      showColHeaders: r.showColHeaders,
      showRowStripes: r.showRowStripes,
      showColStripes: r.showColStripes,
      showLastColumn: r.showLastColumn,
    };
  }

  /** Write a PivotTable's style element. */
  setPivotStyleInfo(sheet: number, pivotIndex: number, info: PivotStyleInfo): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotStyleInfo) return false;
    return wb(this).pivotSetStyleInfo?.(sheet, pivotIndex, info).ok === true;
  }
}

export function installTableStylesMethods(target: WorkbookHandleCtor): void {
  for (const key of Object.getOwnPropertyNames(WorkbookHandleTableStylesMethods.prototype)) {
    if (key === 'constructor') continue;
    const descriptor = Object.getOwnPropertyDescriptor(
      WorkbookHandleTableStylesMethods.prototype,
      key,
    );
    if (!descriptor) continue;
    Object.defineProperty(target.prototype, key, descriptor);
  }
}
