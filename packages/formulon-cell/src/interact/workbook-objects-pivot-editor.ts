import type { History } from '../commands/history.js';
import { type PivotAuthorization, pivotReportRange } from '../commands/pivot-actions.js';
import { pivotCacheFieldIsNumeric } from '../commands/pivot-cache-source.js';
import { applyPivotEdit, type PivotEditField } from '../commands/pivot-edit.js';
import { deletePivotTables } from '../commands/pivot-table.js';
import { formatA1Cell, parseA1Atom } from '../engine/address.js';
import type { PivotDefinition } from '../engine/pivot-definition.js';
import {
  PivotAggregation,
  PivotAxis,
  type PivotFilterSpec,
  PivotReportLayout,
} from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { Strings } from '../i18n/strings.js';
import { createDialogSelect } from '../toolbar/dialogs/form-controls.js';
import { commitPivotChange } from './pivot-field-settings-dialog-apply.js';
import {
  createWorkbookObjectsActionButton,
  pivotEditCheck,
  pivotEditField,
} from './workbook-objects-dom.js';

/** What the pivot edit form needs from the panel that hosts it. */
export interface PivotEditorContext extends PivotAuthorization {
  host: HTMLElement;
  wb: WorkbookHandle;
  strings: Strings;
  history?: History | null;
  /** The report was rewritten, by an edit or by its undo or redo. */
  onChanged(): void;
  /** Error shown when the form is built. */
  errorText: string;
  setError(message: string): void;
  rerender(): void;
  /** The pivot was deleted. */
  onRemoved(): void;
  /** The edits were written back. */
  onApplied(): void;
}

/** The pivot summary the editors are built from. */
export interface PivotEditorTarget {
  sheetIndex: number;
  pivotIndex: number;
  top: number;
  left: number;
  rows: number;
  cols: number;
  fields: readonly string[];
  fieldItems?: Record<string, readonly string[]>;
  fieldItemIndexes?: Record<string, readonly number[]>;
}

const AXIS_NONE = 'none';

/** One field of the edit form: its stored settings, or a guess when the
 *  engine cannot read the definition back. */
interface PivotEditorField {
  index: number;
  sourceName: string;
  label: string;
  /** Null for a field that sits on no axis. */
  axis: PivotAxis | null;
  aggregation: PivotAggregation;
  numberFormat: string;
  /** Whether the automatic subtotal shows; false with no custom functions means none. */
  defaultSubtotal: boolean;
  subtotalFns: readonly PivotAggregation[];
  /** The area a field takes when it is ticked into the report: Values for a
   *  numeric column, Rows otherwise. */
  defaultAxis: PivotAxis;
  items: { label: string; cacheIndex?: number; visible: boolean }[];
}

/** Fields in the order the report shows them: row axis, column axis, then the rest. */
function definitionEditorFields(
  wb: WorkbookHandle,
  definition: PivotDefinition,
  inferredItems: (sourceName: string) => PivotEditorField['items'],
): PivotEditorField[] {
  const order = [
    ...new Set([
      ...definition.rowFieldOrder,
      ...definition.colFieldOrder,
      ...definition.fields.map((field) => field.index),
    ]),
  ];
  const out: PivotEditorField[] = [];
  for (const index of order) {
    const field = definition.fields[index];
    if (!field) continue;
    const dataField = definition.dataFields.find((d) => d.fieldIndex === index);
    out.push({
      index,
      sourceName: field.sourceName,
      label: field.customName || field.sourceName,
      axis: field.axis,
      aggregation: dataField?.aggregation ?? PivotAggregation.Sum,
      numberFormat: dataField?.numberFormat ?? '',
      defaultSubtotal: field.defaultSubtotal,
      subtotalFns: field.subtotalFns,
      defaultAxis: pivotCacheFieldIsNumeric(wb, definition.cacheId, field.cacheFieldIndex)
        ? PivotAxis.Value
        : PivotAxis.Row,
      items:
        field.items.length > 0
          ? field.items.map((item) => ({
              label: item.label,
              ...(item.cacheIndex >= 0 ? { cacheIndex: item.cacheIndex } : {}),
              visible: item.visible,
            }))
          : inferredItems(field.sourceName),
    });
  }
  return out;
}

/** The pivot as the editors see it: the stored definition when the engine can
 *  read it back, and every field with its settings. */
export interface PivotEditorModel {
  definition: PivotDefinition | null;
  fields: PivotEditorField[];
  /** Listed filters as read when the editor opened. */
  pivotFilters: readonly PivotFilterSpec[];
}

export function readPivotEditorModel(
  wb: WorkbookHandle,
  pivot: PivotEditorTarget,
): PivotEditorModel {
  const inferredItemsOf = (sourceName: string): PivotEditorField['items'] => {
    const labels = pivot.fieldItems?.[sourceName] ?? [];
    const indexes = pivot.fieldItemIndexes?.[sourceName] ?? [];
    return labels.map((label, position) => {
      const cacheIndex = indexes[position];
      return { label, ...(cacheIndex !== undefined ? { cacheIndex } : {}), visible: true };
    });
  };
  const definition = wb.getPivotDefinition(pivot.sheetIndex, pivot.pivotIndex);
  const fields: PivotEditorField[] = definition
    ? definitionEditorFields(wb, definition, inferredItemsOf)
    : pivot.fields.map((sourceName, index) => ({
        index,
        sourceName,
        label: sourceName,
        axis: index === 0 ? PivotAxis.Row : PivotAxis.Value,
        aggregation: PivotAggregation.Sum,
        numberFormat: '',
        defaultSubtotal: true,
        subtotalFns: [],
        defaultAxis: PivotAxis.Row,
        items: inferredItemsOf(sourceName),
      }));
  const pivotFilters = (pivot as { pivotFilters?: readonly PivotFilterSpec[] }).pivotFilters ?? [];
  return { definition, fields, pivotFilters };
}

/** The settings controls of one field, shared by the edit form and the field list pane. */
export interface PivotFieldEditor {
  readonly model: PivotEditorField;
  /** Aggregation and number format, for a field on Values. */
  readonly valueSettings: HTMLElement;
  /** Item checklist, for a field on Filters. */
  readonly filterSettings: HTMLElement;
  /** Subtotal mode and functions, for a field on Rows or Columns. */
  readonly subtotalSettings: HTMLElement;
  /** Shows the settings blocks that apply to a field on `axis`. */
  showFor(axis: PivotAxis | null): void;
  /** The field as written back with the field on `axis`. */
  toEditField(axis: PivotAxis | null): PivotEditField;
  /** True when Custom subtotals are shown with no function ticked. */
  subtotalEmpty(): boolean;
}

const aggregationOptionsOf = (t: Strings['workbookObjects']) => [
  { value: String(PivotAggregation.Sum), label: t.pivotAggregateSum },
  { value: String(PivotAggregation.Count), label: t.pivotAggregateCount },
  { value: String(PivotAggregation.Average), label: t.pivotAggregateAverage },
  { value: String(PivotAggregation.Max), label: t.pivotAggregateMax },
  { value: String(PivotAggregation.Min), label: t.pivotAggregateMin },
  { value: String(PivotAggregation.Product), label: t.pivotAggregateProduct },
  { value: String(PivotAggregation.CountNumbers), label: t.pivotAggregateCountNumbers },
  { value: String(PivotAggregation.StdDev), label: t.pivotAggregateStdDev },
  { value: String(PivotAggregation.StdDevP), label: t.pivotAggregateStdDevP },
  { value: String(PivotAggregation.Var), label: t.pivotAggregateVariance },
  { value: String(PivotAggregation.VarP), label: t.pivotAggregateVarianceP },
];

/** Builds the settings controls of one field. */
export function createPivotFieldEditor(
  ctx: { strings: Strings },
  model: PivotEditorField,
): PivotFieldEditor {
  const t = ctx.strings.workbookObjects;
  const { index } = model;
  const aggregationOptions = aggregationOptionsOf(t);
  const aggregation = createDialogSelect(aggregationOptions, String(model.aggregation), {
    className: 'fc-objects__input',
  });
  aggregation.dataset.pivotAggregationFieldIndex = String(index);
  const numberFormat = document.createElement('input');
  numberFormat.className = 'fc-objects__input';
  numberFormat.type = 'text';
  numberFormat.placeholder = t.pivotNumberFormatPlaceholder;
  numberFormat.value = model.numberFormat;
  numberFormat.dataset.pivotNumberFormatFieldIndex = String(index);
  const filterItems = document.createElement('textarea');
  filterItems.className = 'fc-objects__input';
  filterItems.rows = 3;
  filterItems.placeholder = t.pivotFilterItemsPlaceholder;
  filterItems.dataset.pivotFilterItemsFieldIndex = String(index);
  const items = model.items;
  const filterChecklist = document.createElement('div');
  filterChecklist.className = 'fc-objects__pivot-filter-items';
  filterChecklist.dataset.pivotFilterChecklistFieldIndex = String(index);
  for (const item of items) {
    const itemName = item.label;
    const check = document.createElement('input');
    check.type = 'checkbox';
    check.checked = item.visible;
    check.value = itemName;
    // The blank member has no label to be matched by, so the filter can
    // only name it by its cache index. Carry the index for every item and
    // the writeback never has to guess which form to use.
    if (item.cacheIndex !== undefined) check.dataset.pivotItemCacheIndex = String(item.cacheIndex);
    const checkLabel = document.createElement('label');
    checkLabel.className = 'fc-objects__pivot-field-list-item';
    checkLabel.append(check, document.createTextNode(itemName || t.pivotBlankItem));
    filterChecklist.appendChild(checkLabel);
  }
  const valueSettings = document.createElement('div');
  valueSettings.className = 'fc-objects__pivot-value-settings';
  valueSettings.append(
    pivotEditField(t.pivotAggregation, aggregation),
    pivotEditField(t.pivotNumberFormat, numberFormat),
  );
  const filterSettings = document.createElement('div');
  filterSettings.className = 'fc-objects__pivot-value-settings';
  filterSettings.appendChild(
    items.length > 0
      ? pivotEditField(t.pivotFilterItems, filterChecklist)
      : pivotEditField(t.pivotFilterItems, filterItems),
  );
  const subtotalNone = !model.defaultSubtotal && model.subtotalFns.length === 0;
  const subtotalMode = createDialogSelect(
    [
      { value: 'auto', label: t.pivotSubtotalAutomatic },
      ...(subtotalNone ? [{ value: 'none', label: t.pivotSubtotalNone }] : []),
      { value: 'custom', label: t.pivotSubtotalCustom },
    ],
    subtotalNone ? 'none' : model.subtotalFns.length > 0 ? 'custom' : 'auto',
    { className: 'fc-objects__input' },
  );
  subtotalMode.dataset.pivotSubtotalModeFieldIndex = String(index);
  const subtotalFnList = document.createElement('div');
  subtotalFnList.className = 'fc-objects__pivot-filter-items';
  subtotalFnList.dataset.pivotSubtotalFnsFieldIndex = String(index);
  for (const option of aggregationOptions) {
    const check = document.createElement('input');
    check.type = 'checkbox';
    check.value = option.value;
    check.checked = model.subtotalFns.includes(Number(option.value) as PivotAggregation);
    const checkLabel = document.createElement('label');
    checkLabel.className = 'fc-objects__pivot-field-list-item';
    checkLabel.append(check, document.createTextNode(option.label));
    subtotalFnList.appendChild(checkLabel);
  }
  const subtotalSettings = document.createElement('div');
  subtotalSettings.className = 'fc-objects__pivot-value-settings';
  subtotalSettings.append(
    pivotEditField(t.pivotSubtotals, subtotalMode),
    pivotEditField(t.pivotSubtotalCustom, subtotalFnList),
  );
  const selectedSubtotalFns = (): PivotAggregation[] =>
    subtotalMode.value === 'custom'
      ? Array.from(subtotalFnList.querySelectorAll<HTMLInputElement>('input:checked'))
          .map((check) => Number(check.value) as PivotAggregation)
          .sort((a, b) => a - b)
      : [];
  const subtotalsChanged = (): boolean => {
    if (subtotalMode.value === 'none') return false;
    // A field with no subtotal goes back to Automatic by clearing its (empty) function list.
    if (subtotalNone) return true;
    const next = selectedSubtotalFns();
    return (
      next.length !== model.subtotalFns.length ||
      next.some((fn, position) => fn !== model.subtotalFns[position])
    );
  };
  const syncSubtotalFns = (): void => {
    subtotalFnList.hidden = subtotalMode.value !== 'custom';
  };
  subtotalMode.addEventListener('change', syncSubtotalFns);
  syncSubtotalFns();
  const showFor = (axis: PivotAxis | null): void => {
    valueSettings.hidden = axis !== PivotAxis.Value;
    filterSettings.hidden = axis !== PivotAxis.Page;
    subtotalSettings.hidden = axis !== PivotAxis.Row && axis !== PivotAxis.Col;
  };
  const subtotalFns = (): readonly PivotAggregation[] | undefined =>
    !subtotalSettings.hidden && subtotalsChanged() ? selectedSubtotalFns() : undefined;
  return {
    model,
    valueSettings,
    filterSettings,
    subtotalSettings,
    showFor,
    toEditField: (axis) => {
      const fns = subtotalFns();
      return {
        fieldIndex: index,
        axis,
        ...(fns ? { subtotalFns: fns } : {}),
        aggregation: Number(aggregation.value) as PivotAggregation,
        numberFormat: numberFormat.value.trim(),
        filterItems: filterItems.value
          .split(/\r?\n/)
          .map((item) => item.trim())
          .filter(Boolean),
        filterChecklist:
          items.length > 0
            ? Array.from(
                filterChecklist.querySelectorAll<HTMLInputElement>('input[type="checkbox"]'),
              ).map((item) => {
                const cacheIndex = Number.parseInt(item.dataset.pivotItemCacheIndex ?? '', 10);
                return {
                  value: item.value,
                  checked: item.checked,
                  ...(Number.isFinite(cacheIndex) ? { cacheIndex } : {}),
                };
              })
            : null,
      };
    },
    subtotalEmpty: () =>
      !subtotalSettings.hidden &&
      subtotalMode.value === 'custom' &&
      selectedSubtotalFns().length === 0,
  };
}

/** The filters that survive the edit: those whose field is still placed on Rows or
 *  Columns, since only those fields carry a condition. Null when none is dropped, so
 *  the report's filters are left as they are. */
export function retainedPivotFilters(
  pivotFilters: readonly PivotFilterSpec[],
  placed: readonly { editor: { model: { sourceName: string } }; axis: PivotAxis | null }[],
): PivotFilterSpec[] | null {
  const kept = new Set(
    placed
      .filter(({ axis }) => axis === PivotAxis.Row || axis === PivotAxis.Col)
      .map(({ editor }) => editor.model.sourceName),
  );
  const retained = pivotFilters.filter((spec) => kept.has(spec.fieldName));
  return retained.length === pivotFilters.length ? null : retained;
}

/** The full pivot edit form: name, anchor, layout, per-field areas and totals. */
export function renderPivotEditForm(
  ctx: PivotEditorContext,
  pivot: PivotEditorTarget,
): HTMLFormElement {
  const { wb, strings } = ctx;
  const t = strings.workbookObjects;
  const form = document.createElement('form');
  form.className = 'fc-objects__pivot-edit';
  form.setAttribute('aria-label', t.editPivotTable);
  const editorModel = readPivotEditorModel(wb, pivot);
  const { definition } = editorModel;
  const name = document.createElement('input');
  name.className = 'fc-objects__input';
  name.type = 'text';
  name.value = definition?.name ?? `${t.pivot} ${pivot.pivotIndex + 1}`;
  const anchor = document.createElement('input');
  anchor.className = 'fc-objects__input';
  anchor.type = 'text';
  anchor.value = formatA1Cell(pivot.top, pivot.left);
  const rowTotals = document.createElement('input');
  rowTotals.type = 'checkbox';
  rowTotals.checked = definition?.grandTotals.rows ?? true;
  const colTotals = document.createElement('input');
  colTotals.type = 'checkbox';
  colTotals.checked = definition?.grandTotals.cols ?? true;
  const layoutSelect = createDialogSelect(
    [
      { value: String(PivotReportLayout.Compact), label: t.pivotReportLayoutCompact },
      { value: String(PivotReportLayout.Outline), label: t.pivotReportLayoutOutline },
      { value: String(PivotReportLayout.Tabular), label: t.pivotReportLayoutTabular },
    ],
    String(
      wb.getPivotReportLayout(pivot.sheetIndex, pivot.pivotIndex) ?? PivotReportLayout.Compact,
    ),
    { className: 'fc-objects__input' },
  );
  const fieldAreas = document.createElement('div');
  fieldAreas.className = 'fc-objects__pivot-field-areas';
  const fieldAreasTitle = document.createElement('span');
  fieldAreasTitle.textContent = t.pivotFieldAreas;
  fieldAreas.appendChild(fieldAreasTitle);
  const axisOptions = [
    { value: String(PivotAxis.Row), label: t.pivotAreaRows },
    { value: String(PivotAxis.Col), label: t.pivotAreaColumns },
    { value: String(PivotAxis.Page), label: t.pivotAreaFilters },
    { value: String(PivotAxis.Value), label: t.pivotAreaValues },
    { value: AXIS_NONE, label: t.pivotAreaNone },
  ];
  const editors: { editor: PivotFieldEditor; axisSelect: HTMLSelectElement }[] = [];
  for (const model of editorModel.fields) {
    const { index } = model;
    const editor = createPivotFieldEditor(ctx, model);
    const select = createDialogSelect(
      axisOptions,
      model.axis === null ? AXIS_NONE : String(model.axis),
      { className: 'fc-objects__input' },
    );
    select.dataset.pivotFieldIndex = String(index);
    const row = document.createElement('div');
    row.className = 'fc-objects__pivot-field-row';
    row.appendChild(pivotEditField(model.label, select));
    const syncArea = (): void =>
      editor.showFor(select.value === AXIS_NONE ? null : (Number(select.value) as PivotAxis));
    select.addEventListener('change', syncArea);
    syncArea();
    editors.push({ editor, axisSelect: select });
    row.append(editor.valueSettings, editor.filterSettings, editor.subtotalSettings);
    fieldAreas.appendChild(row);
  }
  const error = document.createElement('div');
  error.className = 'fc-objects__error';
  error.setAttribute('role', 'alert');
  error.hidden = !ctx.errorText;
  error.textContent = ctx.errorText;
  const actions = document.createElement('div');
  actions.className = 'fc-objects__actions';
  const remove = createWorkbookObjectsActionButton(t.deletePivotTable);
  const apply = createWorkbookObjectsActionButton(t.apply, { primary: true, type: 'submit' });
  actions.append(remove, apply);
  remove.addEventListener('click', () => {
    ctx.setError('');
    const range = pivotReportRange(wb, pivot.sheetIndex, pivot.pivotIndex);
    const removed =
      range !== null &&
      deletePivotTables({ ...ctx, workbook: wb, commandId: 'pivotDelete' }, pivot.sheetIndex, [
        { pivotIndex: pivot.pivotIndex, range },
      ]);
    if (!removed) {
      ctx.setError(t.pivotEditFailed);
      ctx.rerender();
      return;
    }
    ctx.onRemoved();
    ctx.rerender();
  });
  form.addEventListener('submit', (event) => {
    event.preventDefault();
    ctx.setError('');
    const nextAnchor = parseA1Atom(anchor.value);
    if (!nextAnchor) {
      ctx.setError(t.invalidPivotAnchor);
      ctx.rerender();
      return;
    }
    if (editors.some(({ editor }) => editor.subtotalEmpty())) {
      ctx.setError(t.pivotSubtotalCustomEmpty);
      ctx.rerender();
      return;
    }
    const placed = editors.map(({ editor, axisSelect }) => ({
      editor,
      axis: axisSelect.value === AXIS_NONE ? null : (Number(axisSelect.value) as PivotAxis),
    }));
    const applied = commitPivotChange({ ...ctx, wb }, pivot.sheetIndex, pivot.pivotIndex, () =>
      applyPivotEdit(wb, pivot, {
        fieldListOnly: false,
        name: name.value.trim(),
        anchor: nextAnchor,
        rowGrandTotals: rowTotals.checked,
        colGrandTotals: colTotals.checked,
        layout: Number(layoutSelect.value) as PivotReportLayout,
        ...(definition ? { dataFields: definition.dataFields } : {}),
        fields: placed.map(({ editor, axis }) => editor.toEditField(axis)),
        filters: retainedPivotFilters(editorModel.pivotFilters, placed),
      }),
    );
    if (!applied) {
      ctx.setError(t.pivotEditFailed);
      ctx.rerender();
      return;
    }
    ctx.onApplied();
    ctx.rerender();
  });
  form.append(
    pivotEditField(t.pivotName, name),
    pivotEditField(t.pivotAnchorCell, anchor),
    pivotEditField(t.pivotReportLayout, layoutSelect),
    fieldAreas,
    pivotEditCheck(t.rowGrandTotals, rowTotals),
    pivotEditCheck(t.columnGrandTotals, colTotals),
    error,
    actions,
  );
  return form;
}
