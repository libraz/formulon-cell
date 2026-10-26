import { applyPivotEdit } from '../commands/pivot-edit.js';
import { pivotCacheFieldIsNumeric } from '../commands/pivot-table.js';
import { formatA1Cell, parseA1Atom } from '../engine/address.js';
import {
  PivotAggregation,
  PivotAxis,
  type PivotFilterSpec,
  PivotReportLayout,
} from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { PivotDefinition } from '../engine/workbook-handle-pivot.js';
import type { Strings } from '../i18n/strings.js';
import { createDialogSelect } from '../toolbar/dialogs/form-controls.js';
import {
  createPivotFilterConditionControls,
  type PivotFilterConditionState,
  pivotFilterConditionToSpec,
  pivotFilterSpecToCondition,
  showPivotFilterDialog,
} from './pivot-field-settings.js';
import {
  createWorkbookObjectsActionButton,
  pivotEditCheck,
  pivotEditField,
} from './workbook-objects-dom.js';

/** What the pivot edit form needs from the panel that hosts it. */
export interface PivotEditorContext {
  host: HTMLElement;
  wb: WorkbookHandle;
  strings: Strings;
  /** Error shown when the form is built. */
  errorText: string;
  setError(message: string): void;
  rerender(): void;
  /** The pivot was deleted. */
  onRemoved(): void;
  /** The edits were written back. */
  onApplied(): void;
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

export function renderPivotEditForm(
  ctx: PivotEditorContext,
  pivot: {
    sheetIndex: number;
    pivotIndex: number;
    top: number;
    left: number;
    rows: number;
    cols: number;
    fields: readonly string[];
    fieldItems?: Record<string, readonly string[]>;
    fieldItemIndexes?: Record<string, readonly number[]>;
  },
  opts: { fieldListOnly?: boolean } = {},
): HTMLFormElement {
  const { host, wb, strings } = ctx;
  const t = strings.workbookObjects;
  const fieldListOnly = opts.fieldListOnly === true;
  const form = document.createElement('form');
  form.className = 'fc-objects__pivot-edit';
  form.setAttribute('aria-label', fieldListOnly ? t.pivotFieldList : t.editPivotTable);
  const inferredItemsOf = (sourceName: string): PivotEditorField['items'] => {
    const labels = pivot.fieldItems?.[sourceName] ?? [];
    const indexes = pivot.fieldItemIndexes?.[sourceName] ?? [];
    return labels.map((label, position) => {
      const cacheIndex = indexes[position];
      return { label, ...(cacheIndex !== undefined ? { cacheIndex } : {}), visible: true };
    });
  };
  const definition = wb.getPivotDefinition(pivot.sheetIndex, pivot.pivotIndex);
  const editorFields: PivotEditorField[] = definition
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
  const fieldListTitle = document.createElement('div');
  fieldListTitle.className = 'fc-objects__pivot-field-list-title';
  fieldListTitle.textContent = t.pivotFieldList;
  const availableFields = document.createElement('div');
  availableFields.className = 'fc-objects__pivot-field-list';
  const availableTitle = document.createElement('span');
  availableTitle.textContent = t.pivotAvailableFields;
  availableFields.appendChild(availableTitle);
  const axisOptions = [
    { value: String(PivotAxis.Row), label: t.pivotAreaRows },
    { value: String(PivotAxis.Col), label: t.pivotAreaColumns },
    { value: String(PivotAxis.Page), label: t.pivotAreaFilters },
    { value: String(PivotAxis.Value), label: t.pivotAreaValues },
  ];
  const aggregationOptions = [
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
  const filterConditions = new Map<number, PivotFilterConditionState>();
  const pivotFilters = (pivot as { pivotFilters?: readonly PivotFilterSpec[] }).pivotFilters ?? [];
  for (const spec of pivotFilters) {
    const fieldIndex = editorFields.find((field) => field.sourceName === spec.fieldName)?.index;
    if (fieldIndex === undefined || filterConditions.has(fieldIndex)) continue;
    const condition = pivotFilterSpecToCondition(spec);
    if (condition) filterConditions.set(fieldIndex, condition);
  }
  const filterConditionDirty = new Set<number>();
  const valueFieldSettings: {
    fieldIndex: number;
    axisSelect: HTMLSelectElement;
    aggregationSelect: HTMLSelectElement;
    numberFormatInput: HTMLInputElement;
    filterItemsInput: HTMLTextAreaElement;
    filterChecklist: HTMLElement | null;
    filterCondition(): PivotFilterConditionState | undefined;
    /** Custom subtotal functions to write, or undefined when unchanged. */
    subtotalFns(): readonly PivotAggregation[] | undefined;
    /** True when Custom is chosen with no function ticked. */
    subtotalEmpty(): boolean;
  }[] = [];
  for (const model of editorFields) {
    const { index } = model;
    const field = model.label;
    // A field on no axis is offered "not in report"; ticking it back in
    // restores its last area, or a default by data type.
    const select = createDialogSelect(
      [...axisOptions, { value: AXIS_NONE, label: t.pivotAreaNone }],
      filterConditions.has(index)
        ? String(PivotAxis.Page)
        : model.axis === null
          ? AXIS_NONE
          : String(model.axis),
      { className: 'fc-objects__input' },
    );
    select.dataset.pivotFieldIndex = String(index);
    const row = document.createElement('div');
    row.className = 'fc-objects__pivot-field-row';
    row.appendChild(pivotEditField(field, select));
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
    if (fieldListOnly && items.length > 0) {
      filterItems.value = items.map((item) => item.label).join('\n');
    }
    const syncFilterCondition = (condition: PivotFilterConditionState): void => {
      if (condition.kind === 'none' || !condition.value.trim()) filterConditions.delete(index);
      else filterConditions.set(index, condition);
    };
    const filterConditionControls = createPivotFilterConditionControls({
      strings: strings.pivotTableDialog,
      condition: filterConditions.get(index),
      selectClassName: 'fc-objects__input',
      valueClassName: 'fc-objects__input',
      valuesContainerClassName: 'fc-objects__pivot-filter-condition-values',
      categoryDataset: { pivotFilterCategoryFieldIndex: String(index) },
      conditionDataset: { pivotFilterConditionFieldIndex: String(index) },
      fieldRow: pivotEditField,
      onChange: syncFilterCondition,
      onUserChange: () => filterConditionDirty.add(index),
    });
    const filterDialogButton = createWorkbookObjectsActionButton(
      strings.pivotTableDialog.filterDialog,
    );
    filterDialogButton.addEventListener('click', () => {
      void showPivotFilterDialog({
        host,
        strings: strings.pivotTableDialog,
        fieldName: field,
        condition: filterConditions.get(index),
        okLabel: strings.pageSetup.ok,
        cancelLabel: strings.pageSetup.cancel,
      }).then((condition) => {
        if (!condition) return;
        syncFilterCondition(condition);
        filterConditionDirty.add(index);
        ctx.rerender();
      });
    });
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
      if (item.cacheIndex !== undefined)
        check.dataset.pivotItemCacheIndex = String(item.cacheIndex);
      const checkLabel = document.createElement('label');
      checkLabel.className = 'fc-objects__pivot-field-list-item';
      checkLabel.append(check, document.createTextNode(itemName || t.pivotBlankItem));
      filterChecklist.appendChild(checkLabel);
    }
    const settings = document.createElement('div');
    settings.className = 'fc-objects__pivot-value-settings';
    settings.hidden = select.value !== String(PivotAxis.Value);
    settings.append(
      pivotEditField(t.pivotAggregation, aggregation),
      pivotEditField(t.pivotNumberFormat, numberFormat),
    );
    const filterSettings = document.createElement('div');
    filterSettings.className = 'fc-objects__pivot-value-settings';
    filterSettings.hidden = select.value !== String(PivotAxis.Page);
    filterSettings.appendChild(
      items.length > 0
        ? pivotEditField(t.pivotFilterItems, filterChecklist)
        : pivotEditField(t.pivotFilterItems, filterItems),
    );
    filterSettings.append(...filterConditionControls);
    filterSettings.appendChild(filterDialogButton);
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
    let listCheck: HTMLInputElement | null = null;
    // The area the field returns to when it is ticked back into the report.
    let lastArea = model.axis === null ? String(model.defaultAxis) : String(model.axis);
    const syncArea = (): void => {
      if (select.value !== AXIS_NONE) lastArea = select.value;
      settings.hidden = select.value !== String(PivotAxis.Value);
      filterSettings.hidden = select.value !== String(PivotAxis.Page);
      const grouped =
        select.value === String(PivotAxis.Row) || select.value === String(PivotAxis.Col);
      subtotalSettings.hidden = !grouped;
      subtotalFnList.hidden = subtotalMode.value !== 'custom';
      if (listCheck) listCheck.checked = select.value !== AXIS_NONE;
    };
    if (fieldListOnly) {
      listCheck = document.createElement('input');
      listCheck.type = 'checkbox';
      listCheck.dataset.pivotFieldListFieldIndex = String(index);
      const listItem = document.createElement('label');
      listItem.className = 'fc-objects__pivot-field-list-item';
      listItem.append(listCheck, document.createTextNode(field));
      availableFields.appendChild(listItem);
      const toggle = listCheck;
      toggle.addEventListener('change', () => {
        select.value = toggle.checked ? lastArea : AXIS_NONE;
        syncArea();
      });
    }
    select.addEventListener('change', syncArea);
    subtotalMode.addEventListener('change', syncArea);
    syncArea();
    valueFieldSettings.push({
      fieldIndex: index,
      axisSelect: select,
      aggregationSelect: aggregation,
      numberFormatInput: numberFormat,
      filterItemsInput: filterItems,
      filterChecklist: items.length > 0 ? filterChecklist : null,
      filterCondition: () => filterConditions.get(index),
      subtotalFns: () =>
        !subtotalSettings.hidden && subtotalsChanged() ? selectedSubtotalFns() : undefined,
      subtotalEmpty: () =>
        !subtotalSettings.hidden &&
        subtotalMode.value === 'custom' &&
        selectedSubtotalFns().length === 0,
    });
    row.appendChild(settings);
    row.appendChild(filterSettings);
    row.appendChild(subtotalSettings);
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
  if (fieldListOnly) actions.append(apply);
  else actions.append(remove, apply);
  remove.addEventListener('click', () => {
    ctx.setError('');
    if (!wb.removePivotTable(pivot.sheetIndex, pivot.pivotIndex)) {
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
    const nextAnchor = fieldListOnly
      ? { row: pivot.top, col: pivot.left }
      : parseA1Atom(anchor.value);
    if (!nextAnchor) {
      ctx.setError(t.invalidPivotAnchor);
      ctx.rerender();
      return;
    }
    if (valueFieldSettings.some((field) => field.subtotalEmpty())) {
      ctx.setError(t.pivotSubtotalCustomEmpty);
      ctx.rerender();
      return;
    }
    const sourceNameOf = (fieldIndex: number): string =>
      editorFields.find((model) => model.index === fieldIndex)?.sourceName ?? '';
    // Filters of fields the user did not touch are written back as read, and
    // those of a field now out of the report went with it.
    const editedFilterFields = new Set(
      valueFieldSettings
        .filter((field) => filterConditionDirty.has(field.fieldIndex))
        .map((field) => sourceNameOf(field.fieldIndex)),
    );
    const outOfReport = new Set(
      valueFieldSettings
        .filter((field) => field.axisSelect.value === AXIS_NONE)
        .map((field) => sourceNameOf(field.fieldIndex)),
    );
    const dirtyFilterSettings = valueFieldSettings.filter(
      (field) =>
        field.axisSelect.value === String(PivotAxis.Page) &&
        filterConditionDirty.has(field.fieldIndex),
    );
    const applied = applyPivotEdit(wb, pivot, {
      fieldListOnly,
      name: name.value.trim(),
      anchor: nextAnchor,
      rowGrandTotals: rowTotals.checked,
      colGrandTotals: colTotals.checked,
      layout: Number(layoutSelect.value) as PivotReportLayout,
      ...(definition ? { dataFields: definition.dataFields } : {}),
      fields: valueFieldSettings.map((field) => ({
        fieldIndex: field.fieldIndex,
        axis:
          field.axisSelect.value === AXIS_NONE
            ? null
            : (Number(field.axisSelect.value) as PivotAxis),
        ...(field.subtotalFns() ? { subtotalFns: field.subtotalFns() } : {}),
        aggregation: Number(field.aggregationSelect.value) as PivotAggregation,
        numberFormat: field.numberFormatInput.value.trim(),
        filterItems: field.filterItemsInput.value
          .split(/\r?\n/)
          .map((item) => item.trim())
          .filter(Boolean),
        filterChecklist: field.filterChecklist
          ? Array.from(
              field.filterChecklist.querySelectorAll<HTMLInputElement>('input[type="checkbox"]'),
            ).map((item) => {
              const cacheIndex = Number.parseInt(item.dataset.pivotItemCacheIndex ?? '', 10);
              return {
                value: item.value,
                checked: item.checked,
                ...(Number.isFinite(cacheIndex) ? { cacheIndex } : {}),
              };
            })
          : null,
      })),
      filters:
        editedFilterFields.size === 0
          ? null
          : [
              ...pivotFilters.filter(
                (spec) =>
                  !editedFilterFields.has(spec.fieldName) && !outOfReport.has(spec.fieldName),
              ),
              ...dirtyFilterSettings
                .map((field) =>
                  pivotFilterConditionToSpec(
                    sourceNameOf(field.fieldIndex),
                    field.filterCondition(),
                  ),
                )
                .filter((filter): filter is PivotFilterSpec => filter !== null),
            ],
    });
    if (!applied) {
      ctx.setError(t.pivotEditFailed);
      ctx.rerender();
      return;
    }
    ctx.onApplied();
    ctx.rerender();
  });
  if (fieldListOnly) {
    form.append(fieldListTitle, availableFields, fieldAreas, error, actions);
  } else {
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
  }
  return form;
}
