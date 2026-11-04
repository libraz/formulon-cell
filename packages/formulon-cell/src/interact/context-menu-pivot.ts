import type { History } from '../commands/history.js';
import {
  clearPivotFieldFilter,
  pivotFieldFilterModel,
  setPivotFieldItemsChecked,
} from '../commands/pivot-filter.js';
import { deletePivotTables } from '../commands/pivot-table.js';
import { PIVOT_CELL_KIND } from '../engine/pivot-cell-kind.js';
import { type Addr, PivotAxis, PivotSortOrder } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { Strings } from '../i18n/strings.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';
import {
  type ItemId,
  PIVOT_MENU_AGGREGATION,
  PIVOT_MENU_SHOW_AS,
  type PivotMenuCell,
  type PivotMenuTarget,
} from './context-menu-spec.js';
import type { PivotNumberFormatPicker } from './pivot-field-settings-dialog.js';
import {
  commitPivotChange,
  openPivotFieldSettingsAt,
  setPivotFieldAxes,
  updatePivotDataField,
} from './pivot-field-settings-dialog-apply.js';
import { openPivotFieldFilter } from './pivot-filter-panel.js';
import { openPivotOptions } from './pivot-options-dialog.js';

type PivotCellInfo = NonNullable<ReturnType<WorkbookHandle['pivotCellInfo']>>;

/** The report cell a pivot menu was opened on. */
export interface PivotMenuCellRef {
  addr: Addr;
  info: PivotCellInfo;
}

const cellKind = (info: PivotCellInfo, dataFieldNames: readonly string[]): PivotMenuCell => {
  if (info.isValue) return 'value';
  if (info.kind === PIVOT_CELL_KIND.RowLabel || info.kind === PIVOT_CELL_KIND.ColLabel) {
    return info.fieldIndex >= 0 ? 'item' : 'header';
  }
  if (
    info.kind === PIVOT_CELL_KIND.GrandTotal ||
    info.kind === PIVOT_CELL_KIND.RowSubtotal ||
    info.kind === PIVOT_CELL_KIND.ColSubtotal
  ) {
    return 'total';
  }
  return info.fieldName && dataFieldNames.includes(info.fieldName) ? 'dataHeader' : 'header';
};

/** What the menu shows for a report cell: its kind, names and current settings. */
export function pivotMenuTarget(
  wb: WorkbookHandle,
  ref: PivotMenuCellRef,
  canOpenFieldList: boolean,
  fieldListHidden?: boolean,
): PivotMenuTarget {
  const { addr, info } = ref;
  const definition = wb.getPivotDefinition(addr.sheet, info.pivotIndex);
  const dataFields = definition?.dataFields ?? [];
  const cell = cellKind(
    info,
    dataFields.map((spec) => spec.name ?? ''),
  );
  const field = definition?.fields.find((f) => f.cacheFieldIndex === info.fieldIndex);
  const dataField =
    dataFields.find((spec) => spec.name === info.fieldName) ??
    (cell === 'item' ? undefined : dataFields[0]);
  // The report's own cells carry the item text; the sheet value may not have it yet.
  let value = wb.getValue(addr);
  for (const entry of wb.pivotCells(addr.sheet)) {
    if (entry.addr.row === addr.row && entry.addr.col === addr.col) {
      value = entry.value;
      break;
    }
  }
  return {
    cell,
    fieldName: field ? field.customName || field.sourceName : info.fieldName,
    itemName:
      value.kind === 'text' ? value.value : value.kind === 'number' ? String(value.value) : '',
    dataFieldName: dataField?.name ?? info.fieldName,
    itemOnColumns: field?.axis === PivotAxis.Col,
    ...(dataField ? { aggregation: dataField.aggregation } : {}),
    ...(dataField?.showValuesAs !== undefined ? { showValuesAs: dataField.showValuesAs } : {}),
    subtotalOn: field ? field.defaultSubtotal || field.subtotalFns.length > 0 : false,
    filtered:
      field !== undefined &&
      (pivotFieldFilterModel(wb, addr.sheet, info.pivotIndex, field.index)?.filtered ?? false),
    ...(fieldListHidden === undefined ? {} : { fieldListHidden }),
    canRefresh: info.cacheId >= 0,
    canOpenFieldList,
  };
}

/** The items of the clicked cell's field that the selection covers. */
function selectedItemsOf(ctx: PivotMenuContext, ref: PivotMenuCellRef): Set<string> {
  const { wb } = ctx;
  const out = new Set<string>();
  const selection = ctx.store.getState().selection;
  const ranges = [selection.range, ...(selection.extraRanges ?? [])];
  for (const entry of wb.pivotCells(ref.addr.sheet)) {
    if (entry.pivotIndex !== ref.info.pivotIndex) continue;
    const { row, col } = entry.addr;
    if (!ranges.some((r) => row >= r.r0 && row <= r.r1 && col >= r.c0 && col <= r.c1)) continue;
    const info = wb.pivotCellInfo(entry.addr);
    if (info?.fieldIndex !== ref.info.fieldIndex || info.isValue) continue;
    if (entry.value.kind === 'text') out.add(entry.value.value);
    else if (entry.value.kind === 'number') out.add(String(entry.value.value));
  }
  return out;
}

export interface PivotMenuContext {
  host: HTMLElement;
  store: SpreadsheetStore;
  wb: WorkbookHandle;
  strings: Strings;
  history: History | null;
  pickNumberFormat: PivotNumberFormatPicker;
  /** Repaints the sheet after the report changed. */
  refresh(sheet: number): void;
  /** Where the menu was opened, in viewport coordinates; the filter panel hangs from it. */
  anchor?: { x: number; y: number };
  onOpenPivotFieldList?: (sheetIndex: number, pivotIndex: number) => void;
  isPivotFieldListOpen?: () => boolean;
  onClosePivotFieldList?: () => void;
}

/** Runs a pivot menu entry; false for an id this module does not handle. */
export function runContextMenuPivotItem(
  ctx: PivotMenuContext,
  ref: PivotMenuCellRef,
  id: ItemId,
): boolean {
  const { wb } = ctx;
  const { addr, info } = ref;
  const sheet = addr.sheet;
  const pivot = { sheetIndex: sheet, pivotIndex: info.pivotIndex };
  const commit = { wb, history: ctx.history, onChanged: () => ctx.refresh(sheet) };
  const definition = wb.getPivotDefinition(sheet, info.pivotIndex);
  const dataFieldIndex = (): number => {
    const list = definition?.dataFields ?? [];
    const byName = list.findIndex((spec) => spec.name === info.fieldName);
    return byName >= 0 ? byName : list.length > 0 ? 0 : -1;
  };
  const pivotField = definition?.fields.find((f) => f.cacheFieldIndex === info.fieldIndex);
  const openFieldDialog = (): void => {
    const opened = openPivotFieldSettingsAt(
      { ...commit, host: ctx.host, strings: ctx.strings, pickNumberFormat: ctx.pickNumberFormat },
      addr,
    );
    if (!opened) ctx.onOpenPivotFieldList?.(sheet, info.pivotIndex);
  };

  const aggregation = PIVOT_MENU_AGGREGATION[id];
  if (aggregation !== undefined) {
    const index = dataFieldIndex();
    if (index >= 0) updatePivotDataField(commit, pivot, index, { aggregation }, true);
    return true;
  }
  const showAs = PIVOT_MENU_SHOW_AS[id];
  if (showAs !== undefined) {
    const index = dataFieldIndex();
    if (index >= 0) updatePivotDataField(commit, pivot, index, { showValuesAs: showAs }, false);
    return true;
  }
  switch (id) {
    case 'pivotNumberFormat': {
      const index = dataFieldIndex();
      const current = definition?.dataFields[index]?.numberFormat ?? '';
      if (index < 0) return true;
      ctx.pickNumberFormat(current, (code) => {
        updatePivotDataField(commit, pivot, index, { numberFormat: code }, false);
      });
      return true;
    }
    case 'pivotDelete': {
      const range = wb.pivotBounds(sheet).find((b) => b.pivotIndex === info.pivotIndex)?.range;
      if (!range) return true;
      deletePivotTables({ store: ctx.store, workbook: wb, history: ctx.history }, sheet, [
        { pivotIndex: info.pivotIndex, range },
      ]);
      ctx.refresh(sheet);
      return true;
    }
    case 'pivotSortAsc':
    case 'pivotSortDesc': {
      const order = id === 'pivotSortAsc' ? PivotSortOrder.Ascending : PivotSortOrder.Descending;
      // An item sorts its own field; a value sorts the innermost row field by itself.
      const target =
        pivotField?.index ??
        (info.fieldIndex >= 0 && !info.isValue
          ? info.fieldIndex
          : definition?.rowFieldOrder.at(-1));
      const byField = pivotField ? '' : (definition?.dataFields[dataFieldIndex()]?.name ?? '');
      if (
        target !== undefined &&
        wb.setPivotFieldSort(sheet, info.pivotIndex, target, order, byField)
      ) {
        ctx.refresh(sheet);
      }
      return true;
    }
    case 'pivotSelectTable': {
      const range = wb.pivotBounds(sheet).find((b) => b.pivotIndex === info.pivotIndex)?.range;
      if (range) mutators.setRange(ctx.store, range);
      return true;
    }
    case 'pivotRemoveField': {
      const index = pivotField
        ? pivotField.index
        : definition?.dataFields[dataFieldIndex()]?.fieldIndex;
      if (index !== undefined)
        setPivotFieldAxes(commit, pivot, (f) => (f === index ? null : undefined));
      return true;
    }
    case 'pivotRemoveValues': {
      setPivotFieldAxes(commit, pivot, (_f, axis) => (axis === PivotAxis.Value ? null : undefined));
      return true;
    }
    case 'pivotMoveFieldToColumns':
    case 'pivotMoveFieldToRows': {
      const index = pivotField?.index;
      const axis = id === 'pivotMoveFieldToColumns' ? PivotAxis.Col : PivotAxis.Row;
      if (index !== undefined)
        setPivotFieldAxes(commit, pivot, (f) => (f === index ? axis : undefined));
      return true;
    }
    case 'pivotRemoveGrandTotal': {
      const totals = wb.getPivotTableGrandTotals(sheet, info.pivotIndex);
      if (!totals) return true;
      // The total row under the report is the column grand total.
      const write = (cols: boolean): void => {
        wb.setPivotTableGrandTotals(sheet, info.pivotIndex, totals.rows, cols);
        ctx.refresh(sheet);
      };
      write(false);
      const history = ctx.history && !ctx.history.isReplaying() ? ctx.history : null;
      history?.push({ undo: () => write(totals.cols), redo: () => write(false) });
      return true;
    }
    case 'pivotFieldSettings':
    case 'pivotSummarizeMore':
    case 'pivotShowAsMore':
      openFieldDialog();
      return true;
    case 'pivotOptions':
      void openPivotOptions(
        {
          host: ctx.host,
          strings: ctx.strings,
          wb,
          history: ctx.history,
          onChanged: () => ctx.refresh(sheet),
        },
        sheet,
        info.pivotIndex,
      );
      return true;
    case 'pivotShowFieldList':
      // Hide Field List: a shown list closes, a hidden one opens again.
      if (ctx.isPivotFieldListOpen?.() && ctx.onClosePivotFieldList) ctx.onClosePivotFieldList();
      else ctx.onOpenPivotFieldList?.(sheet, info.pivotIndex);
      return true;
    case 'pivotFilterTop10':
    case 'pivotFilterLabel':
    case 'pivotFilterValue':
      openPivotFieldFilter(
        { ...commit, host: ctx.host, store: ctx.store, strings: ctx.strings },
        addr,
        ctx.anchor,
      );
      return true;
    case 'pivotFilterClear': {
      const index = pivotField?.index;
      if (index !== undefined) {
        commitPivotChange(commit, sheet, info.pivotIndex, () =>
          clearPivotFieldFilter(wb, sheet, info.pivotIndex, index),
        );
      }
      return true;
    }
    case 'pivotFilterKeepSelected':
    case 'pivotFilterHideSelected': {
      const index = pivotField?.index;
      const model =
        index === undefined ? null : pivotFieldFilterModel(wb, sheet, info.pivotIndex, index);
      if (index === undefined || !model) return true;
      const selected = selectedItemsOf(ctx, ref);
      // The model names every item, so the selection matches by name.
      const keep = id === 'pivotFilterKeepSelected';
      const items = model.items.map((entry) => {
        const hit = selected.has(entry.value);
        return { ...entry, checked: keep ? hit : entry.checked && !hit };
      });
      // Hiding every item would empty the report; the desktop refuses that too.
      if (!items.some((entry) => entry.checked)) return true;
      commitPivotChange(commit, sheet, info.pivotIndex, () =>
        setPivotFieldItemsChecked(wb, sheet, info.pivotIndex, index, items),
      );
      return true;
    }
    default:
      return false;
  }
}
