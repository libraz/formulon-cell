import type {
  PivotAggregation,
  PivotAxis,
  PivotCalendar,
  PivotDataFieldSpec,
  PivotDateGrouping,
  PivotFieldSpec,
  PivotFilterSpec,
  Status,
  Value,
  Workbook,
} from './types.js';

export interface IndexResult {
  status: Status;
  index: number;
}

export interface PivotMutationWorkbook extends Workbook {
  pivotCacheCount(): number;
  pivotCacheIdAt(index: number): IndexResult;
  pivotCacheCreate(requestedId: number): IndexResult;
  pivotCacheRemove(cacheId: number): Status;
  pivotCacheFieldCount(cacheId: number): number;
  pivotCacheFieldName(cacheId: number, fieldIdx: number): { status: Status; value: string };
  pivotCacheFieldAdd(cacheId: number, name: string): IndexResult;
  pivotCacheFieldClear(cacheId: number): Status;
  pivotCacheFieldSharedItemCount(cacheId: number, fieldIdx: number): number;
  pivotCacheFieldSharedItemValue?(
    cacheId: number,
    fieldIdx: number,
    itemIdx: number,
  ): { status: Status; value: Value };
  pivotCacheFieldAddSharedItemNumber(cacheId: number, fieldIdx: number, value: number): Status;
  pivotCacheFieldAddSharedItemText(cacheId: number, fieldIdx: number, value: string): Status;
  pivotCacheFieldAddSharedItemBool(cacheId: number, fieldIdx: number, value: boolean): Status;
  pivotCacheFieldAddSharedItemBlank(cacheId: number, fieldIdx: number): Status;
  pivotCacheFieldAddSharedItemError?(cacheId: number, fieldIdx: number, code: number): Status;
  pivotCacheFieldClearSharedItems(cacheId: number, fieldIdx: number): Status;
  pivotCacheRecordAdd(cacheId: number): IndexResult;
  pivotCacheRecordClear(cacheId: number): Status;
  pivotCacheRecordSetNumber(
    cacheId: number,
    recordIdx: number,
    fieldIdx: number,
    value: number,
  ): Status;
  pivotCacheRecordSetText(
    cacheId: number,
    recordIdx: number,
    fieldIdx: number,
    value: string,
  ): Status;
  pivotCacheRecordSetBool(
    cacheId: number,
    recordIdx: number,
    fieldIdx: number,
    value: boolean,
  ): Status;
  pivotCacheRecordSetBlank(cacheId: number, recordIdx: number, fieldIdx: number): Status;
  pivotCacheRecordSetError(
    cacheId: number,
    recordIdx: number,
    fieldIdx: number,
    code: number,
  ): Status;
  pivotCreate(sheet: number, name: string, cacheId: number, row: number, col: number): IndexResult;
  pivotCacheId?(sheet: number, pivotIdx: number): IndexResult;
  pivotRemove(sheet: number, pivotIdx: number): Status;
  pivotSetName(sheet: number, pivotIdx: number, name: string): Status;
  pivotSetAnchor(
    sheet: number,
    pivotIdx: number,
    row: number,
    col: number,
    rows: number,
    cols: number,
  ): Status;
  pivotSetGrandTotals(
    sheet: number,
    pivotIdx: number,
    rowsEnabled: boolean,
    colsEnabled: boolean,
  ): Status;
  pivotFieldCount(sheet: number, pivotIdx: number): number;
  pivotFieldAdd(sheet: number, pivotIdx: number, spec: PivotFieldSpec): IndexResult;
  pivotFieldClear(sheet: number, pivotIdx: number): Status;
  pivotFieldSetAxis(sheet: number, pivotIdx: number, fieldIdx: number, axis: PivotAxis): Status;
  pivotFieldSetSort(
    sheet: number,
    pivotIdx: number,
    fieldIdx: number,
    ascending: boolean,
    byField: string,
  ): Status;
  pivotFieldSetSubtotalTop(sheet: number, pivotIdx: number, fieldIdx: number, top: boolean): Status;
  pivotFieldAddAggregation(
    sheet: number,
    pivotIdx: number,
    fieldIdx: number,
    agg: PivotAggregation,
  ): Status;
  pivotFieldClearAggregations(sheet: number, pivotIdx: number, fieldIdx: number): Status;
  pivotFieldAddItem(
    sheet: number,
    pivotIdx: number,
    fieldIdx: number,
    name: string,
    visible: boolean,
  ): Status;
  pivotFieldClearItems(sheet: number, pivotIdx: number, fieldIdx: number): Status;
  pivotFieldSetItemVisible(
    sheet: number,
    pivotIdx: number,
    fieldIdx: number,
    itemIdx: number,
    visible: boolean,
  ): Status;
  pivotFieldAddSubtotalFn(
    sheet: number,
    pivotIdx: number,
    fieldIdx: number,
    agg: PivotAggregation,
  ): Status;
  pivotFieldClearSubtotalFns(sheet: number, pivotIdx: number, fieldIdx: number): Status;
  pivotFieldSetDateGroup(
    sheet: number,
    pivotIdx: number,
    fieldIdx: number,
    granularity: PivotDateGrouping,
    calendar: PivotCalendar,
    startYear: number,
    endYear: number,
  ): Status;
  pivotFieldClearDateGroup(sheet: number, pivotIdx: number, fieldIdx: number): Status;
  pivotFieldSetNumberFormat(
    sheet: number,
    pivotIdx: number,
    fieldIdx: number,
    format: string,
  ): Status;
  pivotSetRowFieldOrder(sheet: number, pivotIdx: number, indices: readonly number[]): Status;
  pivotSetColFieldOrder(sheet: number, pivotIdx: number, indices: readonly number[]): Status;
  pivotDataFieldCount(sheet: number, pivotIdx: number): number;
  pivotDataFieldAdd(sheet: number, pivotIdx: number, spec: PivotDataFieldSpec): IndexResult;
  pivotDataFieldSet(
    sheet: number,
    pivotIdx: number,
    dataFieldIdx: number,
    spec: PivotDataFieldSpec,
  ): Status;
  pivotDataFieldClear(sheet: number, pivotIdx: number): Status;
  pivotFilterCount(sheet: number, pivotIdx: number): number;
  pivotFilterSpec?(
    sheet: number,
    pivotIdx: number,
    filterIdx: number,
  ): {
    status: Status;
    spec: PivotFilterSpec;
  };
  pivotFilterAxis?(
    sheet: number,
    pivotIdx: number,
    filterIdx: number,
  ): {
    status: Status;
    value: number;
  };
  pivotFilterFieldName?(
    sheet: number,
    pivotIdx: number,
    filterIdx: number,
  ): {
    status: Status;
    value: string;
  };
  pivotFilterType?(
    sheet: number,
    pivotIdx: number,
    filterIdx: number,
  ): {
    status: Status;
    value: number;
  };
  pivotFilterValueKind?(
    sheet: number,
    pivotIdx: number,
    filterIdx: number,
  ): {
    status: Status;
    value: number;
  };
  pivotFilterValueInt?(
    sheet: number,
    pivotIdx: number,
    filterIdx: number,
  ): {
    status: Status;
    value: number;
  };
  pivotFilterValueDouble?(
    sheet: number,
    pivotIdx: number,
    filterIdx: number,
  ): {
    status: Status;
    value: number;
  };
  pivotFilterValueText?(
    sheet: number,
    pivotIdx: number,
    filterIdx: number,
  ): {
    status: Status;
    value: string;
  };
  pivotFilterValueHighKind?(
    sheet: number,
    pivotIdx: number,
    filterIdx: number,
  ): {
    status: Status;
    value: number;
  };
  pivotFilterValueHighInt?(
    sheet: number,
    pivotIdx: number,
    filterIdx: number,
  ): {
    status: Status;
    value: number;
  };
  pivotFilterValueHighDouble?(
    sheet: number,
    pivotIdx: number,
    filterIdx: number,
  ): {
    status: Status;
    value: number;
  };
  pivotFilterValueHighText?(
    sheet: number,
    pivotIdx: number,
    filterIdx: number,
  ): {
    status: Status;
    value: string;
  };
  pivotFilterAdd(sheet: number, pivotIdx: number, spec: PivotFilterSpec): Status;
  pivotFilterClear(sheet: number, pivotIdx: number): Status;
  pivotFilterRemoveAt(sheet: number, pivotIdx: number, filterIdx: number): Status;
}
