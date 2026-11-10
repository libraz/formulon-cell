import type {
  TableColumn,
  TableColumnTotalsInput,
  TableStyle,
  TableStyleInput,
} from '@libraz/formulon';
import { orFallback } from './result.js';
import type { EngineCapabilities, Workbook } from './types.js';
import type { WorkbookHandle } from './workbook-handle.js';
import { workbookTableAt } from './workbook-handle-tables.js';

type WorkbookHandleCtor = { prototype: WorkbookHandle };
type WorkbookHandleInternals = {
  wb: Workbook;
  assertAlive(): void;
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
  columns: TableColumn[];
}

/** A PivotTable's style name and toggles; an omitted key keeps its value. */
export interface PivotStyleInfo {
  name?: string;
  showRowHeaders?: boolean;
  showColHeaders?: boolean;
  showRowStripes?: boolean;
  showColStripes?: boolean;
  showLastColumn?: boolean;
}

/** A custom table / PivotTable style as stored in `<tableStyles>`. */
export type CustomEngineTableStyle = TableStyle;

function internals(handle: unknown): WorkbookHandleInternals {
  return handle as WorkbookHandleInternals;
}

function assertAlive(handle: unknown): void {
  internals(handle).assertAlive();
}

function wb(handle: unknown): Workbook {
  return internals(handle).wb;
}

export abstract class WorkbookHandleTableStylesMethods {
  declare readonly capabilities: EngineCapabilities;

  /** Style, show-flags and totals of the table at workbook-wide `index`;
   *  null when `index` is out of range. */
  getTableStyleDetails(index: number): TableStyleDetails | null {
    assertAlive(this);
    if (!this.capabilities.tableStyleInfo) return null;
    return orFallback(null, (): TableStyleDetails | null => {
      const table = workbookTableAt(wb(this), index);
      if (!table) return null;
      const props = table.getProperties();
      const style = table.getStyleOptions();
      return {
        styleName: style?.styleName ?? '',
        hasStyleInfo: style !== null,
        showFirstColumn: style?.showFirstColumn === true,
        showLastColumn: style?.showLastColumn === true,
        showRowStripes: style?.showRowStripes === true,
        showColumnStripes: style?.showColumnStripes === true,
        headerRowCount: props.headerRow ? 1 : 0,
        totalsRowCount: props.totalsRow ? 1 : 0,
        // The details keep the formula text without the leading `=` a read carries.
        columns: props.columns.map((column) =>
          column.totalsRowFormula?.startsWith('=')
            ? { ...column, totalsRowFormula: column.totalsRowFormula.slice(1) }
            : column,
        ),
      };
    });
  }

  /** Set one table column's totals-row function and label. While the totals
   *  row is shown the engine writes the column's totals cell to match. */
  setTableColumnTotals(index: number, column: number, totals: TableColumnTotalsInput): boolean {
    assertAlive(this);
    if (!this.capabilities.tableStyleInfo) return false;
    return orFallback(false, () => {
      const table = workbookTableAt(wb(this), index);
      if (!table) return false;
      table.setColumnTotals(column, totals);
      return true;
    });
  }

  /** Every custom table / PivotTable style, in file order. */
  getCustomTableStyles(): CustomEngineTableStyle[] {
    assertAlive(this);
    if (!this.capabilities.customTableStyles) return [];
    return orFallback([], () => [...wb(this).listTableStyles()]);
  }

  /** Add or replace a custom table style. False when the engine refuses it
   *  (built-in name, unknown differential-format index). */
  setCustomTableStyle(style: TableStyleInput): boolean {
    assertAlive(this);
    if (!this.capabilities.customTableStyles) return false;
    return orFallback(false, () => {
      wb(this).setTableStyle(style);
      return true;
    });
  }

  /** Remove the custom table style `name`. */
  removeCustomTableStyle(name: string): boolean {
    assertAlive(this);
    if (!this.capabilities.customTableStyles) return false;
    return orFallback(false, () => {
      wb(this).removeTableStyle(name);
      return true;
    });
  }

  /** The workbook's default table and PivotTable style names; `''` when unset. */
  getTableStyleDefaults(): { tableStyle: string; pivotStyle: string } | null {
    assertAlive(this);
    if (!this.capabilities.customTableStyles) return null;
    return orFallback(null, () => {
      const r = wb(this).getTableStyleDefaults();
      return { tableStyle: r.tableStyleName ?? '', pivotStyle: r.pivotStyleName ?? '' };
    });
  }

  /** Set the default style names; an empty string removes one. */
  setTableStyleDefaults(defaults: { tableStyle?: string; pivotStyle?: string }): boolean {
    assertAlive(this);
    if (!this.capabilities.customTableStyles) return false;
    const name = (value: string | undefined): string | null | undefined =>
      value === undefined ? undefined : value === '' ? null : value;
    const tableStyleName = name(defaults.tableStyle);
    const pivotStyleName = name(defaults.pivotStyle);
    return orFallback(false, () => {
      wb(this).updateTableStyleDefaults({
        ...(tableStyleName === undefined ? {} : { tableStyleName }),
        ...(pivotStyleName === undefined ? {} : { pivotStyleName }),
      });
      return true;
    });
  }

  /** Style name and toggles of a PivotTable; null when it has no style element. */
  getPivotStyleInfo(sheet: number, pivotIndex: number): Required<PivotStyleInfo> | null {
    assertAlive(this);
    if (!this.capabilities.pivotStyleInfo) return null;
    const r = orFallback(null, () => wb(this).sheet(sheet).pivot(pivotIndex).getStyleOptions());
    if (!r) return null;
    return {
      name: r.styleName ?? '',
      showRowHeaders: r.showRowHeaders,
      showColHeaders: r.showColumnHeaders,
      showRowStripes: r.showRowStripes,
      showColStripes: r.showColumnStripes,
      showLastColumn: r.showLastColumn,
    };
  }

  /** Write a PivotTable's style element; an omitted `name` writes it without one. */
  setPivotStyleInfo(sheet: number, pivotIndex: number, info: PivotStyleInfo): boolean {
    assertAlive(this);
    if (!this.capabilities.pivotStyleInfo) return false;
    return orFallback(false, () => {
      wb(this)
        .sheet(sheet)
        .pivot(pivotIndex)
        .updateStyleOptions({
          styleName: info.name ? info.name : null,
          ...(info.showRowHeaders === undefined ? {} : { showRowHeaders: info.showRowHeaders }),
          ...(info.showColHeaders === undefined ? {} : { showColumnHeaders: info.showColHeaders }),
          ...(info.showRowStripes === undefined ? {} : { showRowStripes: info.showRowStripes }),
          ...(info.showColStripes === undefined ? {} : { showColumnStripes: info.showColStripes }),
          ...(info.showLastColumn === undefined ? {} : { showLastColumn: info.showLastColumn }),
        });
      return true;
    });
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
