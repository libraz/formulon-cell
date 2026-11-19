import type {
  Addr,
  PivotAggregation,
  PivotAxis,
  PivotCalendar,
  PivotDataFieldSpec,
  PivotDateGrouping,
  PivotFilterSpec,
  PivotReportLayout,
  PivotSortOrder,
  Range,
} from './types.js';
import type { PivotStyleInfo } from './workbook-handle-table-styles.js';

/** One member of a pivot field, paired with the cache index that addresses it.
 *  The blank member's `label` is empty — it has no label of its own, and the
 *  index is the only way to name it. */
export interface PivotFieldItem {
  readonly label: string;
  readonly cacheIndex: number;
}

/** One member of a pivot field's item list, as stored on the report. */
export interface PivotItemDefinition {
  /** The item's own label; empty for an item bound only by cache index. */
  readonly label: string;
  /** Index into the bound cache field's shared items, -1 for a label-only item. */
  readonly cacheIndex: number;
  readonly visible: boolean;
  /** Whether the item's detail (the items under it) shows; false for a collapsed item. */
  readonly showDetail: boolean;
}

/** A pivot field as stored on the report, whichever axis (if any) it sits on. */
export interface PivotFieldDefinition {
  readonly index: number;
  readonly sourceName: string;
  readonly customName: string;
  /** Null for a field that sits on no axis. */
  readonly axis: PivotAxis | null;
  readonly subtotalTop: boolean;
  /** Whether the automatic subtotal shows; false whenever `subtotalFns` is non-empty. */
  readonly defaultSubtotal: boolean;
  /** Custom subtotal functions that replace the automatic one. */
  readonly subtotalFns: readonly PivotAggregation[];
  /** Number format code, empty for none. */
  readonly numberFormat: string;
  readonly cacheFieldIndex: number;
  readonly sortOrder: PivotSortOrder;
  readonly sortByField: string;
  /** A bound of -1 follows the field's data (automatic). */
  readonly dateGroup: {
    /** The finest of `periods`. */
    readonly granularity: PivotDateGrouping;
    /** Every period the field is grouped by, an OR of `PivotDatePeriod` bits. */
    readonly periods: number;
    readonly calendar: PivotCalendar;
    readonly startYear: number;
    readonly endYear: number;
    readonly intervalDays: number;
    readonly startSerial: number;
    readonly endSerial: number;
  } | null;
  /** For a group field, the index of the field it groups; `null` otherwise. */
  readonly groupBaseIndex: number | null;
  /** Numeric grouping; a `null` bound follows the field's data. */
  readonly numericGroup: PivotNumericGroupSpec | null;
  readonly items: readonly PivotItemDefinition[];
  /** Whether every item of the field is expanded. */
  readonly showDetail: boolean;
}

/** Numeric grouping of a pivot field into ranges of `interval`; a `null` bound is automatic. */
export interface PivotNumericGroupSpec {
  readonly interval: number;
  readonly start: number | null;
  readonly end: number | null;
}

/** Table style and toggles of a PivotTable; `name` is empty for a style element without a named style. */
export type PivotStyleDefinition = Required<PivotStyleInfo>;

/** The stored definition of a PivotTable, read back from the engine. */
export interface PivotDefinition {
  readonly name: string;
  readonly cacheId: number;
  readonly grandTotals: { readonly rows: boolean; readonly cols: boolean };
  readonly layout: PivotReportLayout | null;
  /** Pivot field indices of the row / column axes, in display order. */
  readonly rowFieldOrder: readonly number[];
  readonly colFieldOrder: readonly number[];
  readonly fields: readonly PivotFieldDefinition[];
  /** Value fields in display order; `fieldIndex` is a pivot field index. */
  readonly dataFields: readonly PivotDataFieldSpec[];
  /** Listed filters: those added in this session and the loaded caption, value,
   *  top-N and date ones. Other loaded kinds are kept verbatim by the engine and
   *  survive `clearPivotFilters`, but are not listed. */
  readonly filters: readonly PivotFilterSpec[];
  /** Table style, null when the definition has no style element. */
  readonly style: PivotStyleDefinition | null;
}

/** A report held at #SPILL! because cells are in the way of the area it lays
 *  out over. */
export interface PivotSpill {
  pivotIndex: number;
  anchor: Addr;
  /** The area the report needs. */
  needed: Range;
  /** Non-empty cells inside `needed` that hold the report back. */
  blockers: Addr[];
}

/** A report layout cell as `pivotCellInfo` reports it. */
export interface PivotCellInfo {
  pivotIndex: number;
  kind: number;
  fieldName: string;
  fieldIndex: number;
  /** Cache bound to the report, -1 when the engine cannot say. */
  cacheId: number;
  isValue: boolean;
}
