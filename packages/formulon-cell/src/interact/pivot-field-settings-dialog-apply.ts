import { type History, recordingHistory } from '../commands/history.js';
import {
  type PivotAuthorization,
  pivotFilterBlockRows,
  pivotReportEffects,
  pivotReportRange,
  pivotWrite,
} from '../commands/pivot-actions.js';
import {
  applyPivotEdit,
  type PivotEditField,
  replacePivotDataFields,
  restorePivotFieldLayout,
} from '../commands/pivot-edit.js';
import type { Addr } from '../engine/types.js';
import {
  PivotAggregation,
  PivotAxis,
  type PivotDataFieldSpec,
  PivotReportLayout,
  PivotShowValuesAs,
  type PivotShowValuesAs as PivotShowValuesAsValue,
} from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { PivotCellInfo, PivotDefinition } from '../engine/workbook-handle-pivot.js';
import type { Strings } from '../i18n/strings.js';
import { pivotFieldOfCell } from './pivot-cell-at.js';
import { pivotShowAsNeedsBaseField, pivotShowAsNeedsBaseItem } from './pivot-field-settings.js';
import {
  type PivotFieldSettingsResult,
  type PivotFieldSettingsTarget,
  type PivotNumberFormatPicker,
  showPivotFieldSettingsDialog,
} from './pivot-field-settings-dialog.js';

/** Where a pivot rewrite records its undo step and announces itself. */
export interface PivotCommitContext extends PivotAuthorization {
  wb: WorkbookHandle;
  history?: History | null;
  /** The report was rewritten, by the change or by its undo or redo. */
  onChanged(): void;
}

/** Where the report proper sits: its top-left below any filter block, and its size. */
interface PivotAnchor {
  row: number;
  col: number;
  rows: number;
  cols: number;
}

function pivotAnchorOf(
  wb: WorkbookHandle,
  sheetIndex: number,
  pivotIndex: number,
  definition: PivotDefinition,
): PivotAnchor | null {
  const range = pivotReportRange(wb, sheetIndex, pivotIndex);
  if (!range) return null;
  const row = range.r0 + pivotFilterBlockRows(definition);
  return { row, col: range.c0, rows: range.r1 - row + 1, cols: range.c1 - range.c0 + 1 };
}

/** Runs `apply` against one report as a single undo step: the definition and
 *  the report's position read before and after (after the host's own follow-up
 *  in `onChanged`) are what undo and redo restore. Returns what `apply` did. */
export function commitPivotChange(
  ctx: PivotCommitContext,
  sheetIndex: number,
  pivotIndex: number,
  apply: () => boolean,
): boolean {
  const { wb } = ctx;
  return pivotWrite(ctx, ctx.history, pivotReportEffects(wb, sheetIndex, pivotIndex), false, () =>
    commitAuthorizedPivotChange(ctx, sheetIndex, pivotIndex, apply),
  );
}

/** The change itself, with its undo entry restoring the report's layout and anchor. */
function commitAuthorizedPivotChange(
  ctx: PivotCommitContext,
  sheetIndex: number,
  pivotIndex: number,
  apply: () => boolean,
): boolean {
  const { wb } = ctx;
  const capture = () => {
    const definition = wb.getPivotDefinition(sheetIndex, pivotIndex);
    return definition
      ? {
          definition,
          anchor: pivotAnchorOf(wb, sheetIndex, pivotIndex, definition),
          name: wb.pivotTableName(sheetIndex, pivotIndex),
        }
      : null;
  };
  const before = capture();
  const ok = apply();
  ctx.onChanged();
  const after = before ? capture() : null;
  const history = recordingHistory(ctx.history);
  if (history && before && after) {
    const restore = (target: NonNullable<typeof before>) => (): void => {
      restorePivotFieldLayout(wb, sheetIndex, pivotIndex, target.definition);
      const now = wb.getPivotDefinition(sheetIndex, pivotIndex);
      const at = now ? pivotAnchorOf(wb, sheetIndex, pivotIndex, now) : null;
      const anchor = target.anchor;
      if (anchor && (at?.row !== anchor.row || at?.col !== anchor.col)) {
        wb.setPivotTableAnchor(sheetIndex, pivotIndex, anchor);
      }
      if (target.name && wb.pivotTableName(sheetIndex, pivotIndex) !== target.name) {
        wb.renamePivotTable(sheetIndex, pivotIndex, target.name);
      }
      ctx.onChanged();
    };
    history.push({ undo: restore(before), redo: restore(after) });
  }
  return ok;
}

const isRunningTotal = (kind: number): boolean =>
  kind === PivotShowValuesAs.RunningTotalInRow || kind === PivotShowValuesAs.RunningTotalInCol;

/** The dialog's single Running Total kind; the engine splits it by the axis of
 *  the base field, so the axis decides which of the two is written. */
export const runningTotalFor = (definition: PivotDefinition, baseField: number): number =>
  definition.colFieldOrder.includes(baseField)
    ? PivotShowValuesAs.RunningTotalInCol
    : PivotShowValuesAs.RunningTotalInRow;

const itemLabels = (
  definition: PivotDefinition,
  fieldIndex: number,
  inferred: (sourceName: string) => readonly string[],
): string[] => {
  const field = definition.fields[fieldIndex];
  if (!field) return [];
  return field.items.length > 0
    ? field.items.map((item) => item.label)
    : [...inferred(field.sourceName)];
};

/** The name the engine gives the data field under each function, read by
 *  rewriting the unnamed spec once per function and then restoring the field.
 *  Empty when the field cannot be restored exactly as it was. */
function profileDataFieldNames(
  wb: WorkbookHandle,
  sheetIndex: number,
  pivotIndex: number,
  dataFieldIndex: number,
  spec: PivotDataFieldSpec,
): Partial<Record<PivotAggregation, string>> {
  const { name: _name, ...unnamed } = spec;
  const names: Partial<Record<PivotAggregation, string>> = {};
  for (const aggregation of Object.values(PivotAggregation)) {
    if (
      !wb.setPivotDataField(sheetIndex, pivotIndex, dataFieldIndex, { ...unnamed, aggregation })
    ) {
      break;
    }
    const named = wb.getPivotDataField(sheetIndex, pivotIndex, dataFieldIndex)?.name;
    if (named) names[aggregation] = named;
  }
  wb.setPivotDataField(sheetIndex, pivotIndex, dataFieldIndex, spec);
  const restored = wb.getPivotDataField(sheetIndex, pivotIndex, dataFieldIndex);
  return restored?.name === spec.name && restored?.aggregation === spec.aggregation ? names : {};
}

/** The dialog model of a value field (`dataFieldIndex` into the data fields) or
 *  of a row, column or filter field (`fieldIndex`). Null when the engine cannot
 *  read the report back. */
export function pivotFieldSettingsTarget(
  wb: WorkbookHandle,
  sheetIndex: number,
  pivotIndex: number,
  which: { kind: 'value'; dataFieldIndex: number } | { kind: 'field'; fieldIndex: number },
  inferredItems: (sourceName: string) => readonly string[] = () => [],
): PivotFieldSettingsTarget | null {
  const definition = wb.getPivotDefinition(sheetIndex, pivotIndex);
  if (!definition) return null;
  if (which.kind === 'field') {
    const field = definition.fields[which.fieldIndex];
    if (!field) return null;
    return {
      kind: 'field',
      name: field.customName || field.sourceName,
      subtotal: field.subtotalFns.length > 0 ? 'custom' : field.defaultSubtotal ? 'auto' : 'none',
      subtotalFns: [...field.subtotalFns],
    };
  }
  const spec = definition.dataFields[which.dataFieldIndex];
  const field = spec ? definition.fields[spec.fieldIndex] : undefined;
  if (!spec || !field) return null;
  return {
    kind: 'value',
    autoNames: profileDataFieldNames(wb, sheetIndex, pivotIndex, which.dataFieldIndex, spec),
    sourceName: field.sourceName,
    name: spec.name ?? '',
    aggregation: spec.aggregation,
    numberFormat: spec.numberFormat ?? '',
    showValuesAs: spec.showValuesAs ?? PivotShowValuesAs.Normal,
    ...(spec.showAsBaseField !== undefined && spec.showAsBaseField >= 0
      ? { baseField: spec.showAsBaseField }
      : {}),
    ...(spec.showAsBaseItem !== undefined && spec.showAsBaseItem >= 0
      ? { baseItem: spec.showAsBaseItem }
      : {}),
    baseFields: definition.fields.map((candidate) => ({
      index: candidate.index,
      name: candidate.customName || candidate.sourceName,
      items: itemLabels(definition, candidate.index, inferredItems),
    })),
  };
}

/** Every field as it stands, for an applyPivotEdit that changes only `override`. */
function standingEditFields(
  definition: PivotDefinition,
  override: (field: PivotEditField) => PivotEditField,
): PivotEditField[] {
  const order = [
    ...new Set([
      ...definition.rowFieldOrder,
      ...definition.colFieldOrder,
      ...definition.fields.map((field) => field.index),
    ]),
  ];
  return order.flatMap((index) => {
    const field = definition.fields[index];
    if (!field) return [];
    const dataField = definition.dataFields.find((spec) => spec.fieldIndex === index);
    return [
      override({
        fieldIndex: index,
        axis: field.axis,
        aggregation: dataField?.aggregation ?? PivotAggregation.Sum,
        numberFormat: dataField?.numberFormat ?? '',
        filterItems: [],
        filterChecklist:
          field.axis === PivotAxis.Page
            ? field.items.map((item) => ({
                value: item.label,
                checked: item.visible,
                ...(item.cacheIndex >= 0 ? { cacheIndex: item.cacheIndex } : {}),
              }))
            : null,
      }),
    ];
  });
}

const fieldListEdit = (fields: readonly PivotEditField[], definition: PivotDefinition) => ({
  fieldListOnly: true,
  name: '',
  anchor: { row: 0, col: 0 },
  rowGrandTotals: true,
  colGrandTotals: true,
  layout: PivotReportLayout.Compact,
  dataFields: definition.dataFields,
  fields,
  filters: null,
});

/** Writes a dialog result to the report as one undo step. */
export function applyPivotFieldSettings(
  ctx: PivotCommitContext,
  pivot: { sheetIndex: number; pivotIndex: number; rows: number; cols: number },
  which: { kind: 'value'; dataFieldIndex: number } | { kind: 'field'; fieldIndex: number },
  result: PivotFieldSettingsResult,
): boolean {
  const { wb } = ctx;
  const { sheetIndex, pivotIndex } = pivot;
  return commitPivotChange(ctx, sheetIndex, pivotIndex, () => {
    const definition = wb.getPivotDefinition(sheetIndex, pivotIndex);
    if (!definition) return false;
    if (which.kind === 'field') {
      const { fieldIndex } = which;
      const fields = standingEditFields(definition, (field) => {
        if (field.fieldIndex !== fieldIndex) return field;
        if (result.kind === 'hide') return { ...field, axis: null };
        if (result.kind !== 'field' || result.subtotal === 'none') return field;
        return { ...field, subtotalFns: result.subtotal === 'auto' ? [] : result.subtotalFns };
      });
      return applyPivotEdit(wb, pivot, fieldListEdit(fields, definition));
    }
    if (result.kind !== 'value') return false;
    const current = definition.dataFields[which.dataFieldIndex];
    if (!current) return false;
    const typed = result.name.trim();
    // A name the dialog could not rename after a function change is left to
    // the engine, which names it from the profile ("平均 / Amt").
    const name = typed && result.keepName ? typed : undefined;
    const showValuesAs =
      result.showValuesAs === PivotShowValuesAs.RunningTotalInRow && result.baseField !== undefined
        ? runningTotalFor(definition, result.baseField)
        : result.showValuesAs;
    const next: PivotDataFieldSpec = {
      fieldIndex: current.fieldIndex,
      aggregation: result.aggregation,
      ...(name ? { name } : {}),
      ...(result.numberFormat ? { numberFormat: result.numberFormat } : {}),
      ...(showValuesAs !== PivotShowValuesAs.Normal
        ? { showValuesAs: showValuesAs as PivotDataFieldSpec['showValuesAs'] }
        : {}),
      ...(pivotShowAsNeedsBaseField(showValuesAs as PivotShowValuesAsValue) ||
      (isRunningTotal(showValuesAs) && result.baseField !== undefined)
        ? { showAsBaseField: result.baseField }
        : {}),
      ...(pivotShowAsNeedsBaseItem(showValuesAs as PivotShowValuesAsValue) &&
      result.baseItem !== undefined
        ? { showAsBaseItem: result.baseItem }
        : {}),
    };
    const list = definition.dataFields.map((spec, index) =>
      index === which.dataFieldIndex ? next : spec,
    );
    return replacePivotDataFields(wb, sheetIndex, pivotIndex, list);
  });
}

/** Everything opening the PivotTable Field dialog on a report needs. */
export interface PivotFieldSettingsOpenContext extends PivotCommitContext {
  host: HTMLElement;
  strings: Strings;
  pickNumberFormat?: PivotNumberFormatPicker;
}

/** The field a report cell belongs to: a value cell's data field, or the row,
 *  column or filter field of a label. Null for a cell that names no field. */
function pivotFieldForCell(
  definition: PivotDefinition,
  info: PivotCellInfo,
): { kind: 'value'; dataFieldIndex: number } | { kind: 'field'; fieldIndex: number } | null {
  if (info.isValue) {
    // A value cell carries its data field's caption; a source index is a fallback.
    const byName = definition.dataFields.findIndex((spec) => spec.name === info.fieldName);
    const bySource = definition.dataFields.findIndex(
      (spec) => definition.fields[spec.fieldIndex]?.cacheFieldIndex === info.fieldIndex,
    );
    const dataFieldIndex = byName >= 0 ? byName : bySource;
    return definition.dataFields.length > 0
      ? { kind: 'value', dataFieldIndex: Math.max(dataFieldIndex, 0) }
      : null;
  }
  const field = pivotFieldOfCell(definition, info);
  return field ? { kind: 'field', fieldIndex: field.index } : null;
}

/** Opens the PivotTable Field dialog on the field of the report cell at `addr`
 *  and applies its result as one undo step. False when the cell is in no report
 *  or names no field. The context menu, the panel and the instance all open it here. */
export function openPivotFieldSettingsAt(ctx: PivotFieldSettingsOpenContext, addr: Addr): boolean {
  const { wb } = ctx;
  const info = wb.pivotCellInfo(addr);
  if (!info) return false;
  const definition = wb.getPivotDefinition(addr.sheet, info.pivotIndex);
  if (!definition) return false;
  const which = pivotFieldForCell(definition, info);
  if (!which) return false;
  const target = pivotFieldSettingsTarget(wb, addr.sheet, info.pivotIndex, which);
  if (!target) return false;
  void showPivotFieldSettingsDialog({
    host: ctx.host,
    strings: ctx.strings,
    target,
    ...(ctx.pickNumberFormat ? { pickNumberFormat: ctx.pickNumberFormat } : {}),
  }).then((result) => {
    if (!result) return;
    applyPivotFieldSettings(
      ctx,
      { sheetIndex: addr.sheet, pivotIndex: info.pivotIndex, rows: 0, cols: 0 },
      which,
      result,
    );
  });
  return true;
}

/** Moves fields between axes as one undo step. `axisOf` gives a field's new
 *  axis, null to take it out of the report, or undefined to leave it. */
export function setPivotFieldAxes(
  ctx: PivotCommitContext,
  pivot: { sheetIndex: number; pivotIndex: number },
  axisOf: (fieldIndex: number, axis: PivotAxis | null) => PivotAxis | null | undefined,
): boolean {
  const { wb } = ctx;
  const { sheetIndex, pivotIndex } = pivot;
  return commitPivotChange(ctx, sheetIndex, pivotIndex, () => {
    const definition = wb.getPivotDefinition(sheetIndex, pivotIndex);
    if (!definition) return false;
    const fields = standingEditFields(definition, (field) => {
      const next = axisOf(field.fieldIndex, field.axis);
      return next === undefined ? field : { ...field, axis: next };
    });
    return applyPivotEdit(wb, { ...pivot, rows: 0, cols: 0 }, fieldListEdit(fields, definition));
  });
}

/** Rewrites one data field as one undo step. `rename` drops its caption so the
 *  engine names it from the profile, as a function change does on the desktop. */
export function updatePivotDataField(
  ctx: PivotCommitContext,
  pivot: { sheetIndex: number; pivotIndex: number },
  dataFieldIndex: number,
  patch: { aggregation?: PivotAggregation; numberFormat?: string; showValuesAs?: number },
  rename: boolean,
): boolean {
  const { wb } = ctx;
  const { sheetIndex, pivotIndex } = pivot;
  return commitPivotChange(ctx, sheetIndex, pivotIndex, () => {
    const definition = wb.getPivotDefinition(sheetIndex, pivotIndex);
    const current = definition?.dataFields[dataFieldIndex];
    if (!definition || !current) return false;
    const { name, numberFormat, showValuesAs, showAsBaseField, showAsBaseItem, ...rest } = current;
    const kind = patch.showValuesAs ?? showValuesAs ?? PivotShowValuesAs.Normal;
    const format = patch.numberFormat ?? numberFormat ?? '';
    const next: PivotDataFieldSpec = {
      ...rest,
      ...(patch.aggregation === undefined ? {} : { aggregation: patch.aggregation }),
      ...(!rename && name ? { name } : {}),
      ...(format ? { numberFormat: format } : {}),
      ...(kind !== PivotShowValuesAs.Normal
        ? { showValuesAs: kind as PivotDataFieldSpec['showValuesAs'] }
        : {}),
      ...(pivotShowAsNeedsBaseField(kind as PivotShowValuesAsValue) && showAsBaseField !== undefined
        ? { showAsBaseField }
        : {}),
      ...(pivotShowAsNeedsBaseItem(kind as PivotShowValuesAsValue) && showAsBaseItem !== undefined
        ? { showAsBaseItem }
        : {}),
    };
    const list = definition.dataFields.map((spec, index) =>
      index === dataFieldIndex ? next : spec,
    );
    return replacePivotDataFields(wb, sheetIndex, pivotIndex, list);
  });
}
