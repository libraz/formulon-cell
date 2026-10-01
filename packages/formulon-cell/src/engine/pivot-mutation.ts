import type { Status, Value, Workbook } from './types.js';

/** The common `{ status, index }` shape returned by mutating helpers. */
export interface IndexResult {
  status: Status;
  index: number;
}

/**
 * The npm/native 0.12 Workbook surface already declares the complete pivot
 * mutation API. Keep this type as a small adapter boundary for the two
 * optional hooks that the cell layer can use when an older host or the test
 * engine exposes them:
 *
 * - `pivotCacheFieldSharedItemValue` reads a shared item by cache index;
 * - `pivotCacheId` finds the cache bound to a projected PivotTable.
 *
 * Every required method comes from the upstream `Workbook` declaration. This
 * deliberately avoids re-declaring copied signatures, which used to turn
 * 0.12 status envelopes back into bare numbers and kept removed aggregation
 * methods alive in the adapter.
 */
export interface PivotMutationWorkbook extends Workbook {
  pivotCacheFieldSharedItemValue?: (
    cacheId: number,
    fieldIdx: number,
    itemIdx: number,
  ) => { status: Status; value: Value };
  pivotCacheId?: (sheet: number, pivotIdx: number) => IndexResult;
}
