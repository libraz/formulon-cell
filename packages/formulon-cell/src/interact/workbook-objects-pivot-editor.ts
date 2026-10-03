import { applyPivotEdit } from '../commands/pivot-edit.js';
import { formatA1Cell, parseA1Atom } from '../engine/address.js';
import {
  PivotAggregation,
  PivotAxis,
  type PivotFilterSpec,
  PivotReportLayout,
} from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { Strings } from '../i18n/strings.js';
import { createDialogSelect } from '../toolbar/dialogs/form-controls.js';
import { projectDisabledState } from '../toolbar/menu-a11y.js';
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
  const name = document.createElement('input');
  name.className = 'fc-objects__input';
  name.type = 'text';
  name.value = `${t.pivot} ${pivot.pivotIndex + 1}`;
  const anchor = document.createElement('input');
  anchor.className = 'fc-objects__input';
  anchor.type = 'text';
  anchor.value = formatA1Cell(pivot.top, pivot.left);
  const rowTotals = document.createElement('input');
  rowTotals.type = 'checkbox';
  rowTotals.checked = true;
  const colTotals = document.createElement('input');
  colTotals.type = 'checkbox';
  colTotals.checked = true;
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
  ];
  const filterConditions = new Map<number, PivotFilterConditionState>();
  for (const spec of (pivot as { pivotFilters?: readonly PivotFilterSpec[] }).pivotFilters ?? []) {
    const fieldIndex = pivot.fields.indexOf(spec.fieldName);
    if (fieldIndex < 0 || filterConditions.has(fieldIndex)) continue;
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
  }[] = [];
  for (const [index, field] of pivot.fields.entries()) {
    if (fieldListOnly) {
      const item = document.createElement('label');
      item.className = 'fc-objects__pivot-field-list-item';
      const check = document.createElement('input');
      check.type = 'checkbox';
      check.checked = true;
      projectDisabledState(check, true, t.pivotFieldListCheckboxReadOnly, {
        datasetKey: 'disabledReason',
      });
      item.append(check, document.createTextNode(field));
      availableFields.appendChild(item);
    }
    const select = createDialogSelect(
      axisOptions,
      filterConditions.has(index)
        ? String(PivotAxis.Page)
        : index === 0
          ? String(PivotAxis.Row)
          : String(PivotAxis.Value),
      { className: 'fc-objects__input' },
    );
    select.dataset.pivotFieldIndex = String(index);
    const row = document.createElement('div');
    row.className = 'fc-objects__pivot-field-row';
    row.appendChild(pivotEditField(field, select));
    const aggregation = createDialogSelect(aggregationOptions, String(PivotAggregation.Sum), {
      className: 'fc-objects__input',
    });
    aggregation.dataset.pivotAggregationFieldIndex = String(index);
    const numberFormat = document.createElement('input');
    numberFormat.className = 'fc-objects__input';
    numberFormat.type = 'text';
    numberFormat.placeholder = t.pivotNumberFormatPlaceholder;
    numberFormat.dataset.pivotNumberFormatFieldIndex = String(index);
    const filterItems = document.createElement('textarea');
    filterItems.className = 'fc-objects__input';
    filterItems.rows = 3;
    filterItems.placeholder = t.pivotFilterItemsPlaceholder;
    filterItems.dataset.pivotFilterItemsFieldIndex = String(index);
    const inferredItems = pivot.fieldItems?.[field] ?? [];
    const inferredIndexes = pivot.fieldItemIndexes?.[field] ?? [];
    if (fieldListOnly && inferredItems.length > 0) filterItems.value = inferredItems.join('\n');
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
    for (const [position, itemName] of inferredItems.entries()) {
      const check = document.createElement('input');
      check.type = 'checkbox';
      check.checked = true;
      check.value = itemName;
      // The blank member has no label to be matched by, so the filter can
      // only name it by its cache index. Carry the index for every item and
      // the writeback never has to guess which form to use.
      const cacheIndex = inferredIndexes[position];
      if (cacheIndex !== undefined) check.dataset.pivotItemCacheIndex = String(cacheIndex);
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
      inferredItems.length > 0
        ? pivotEditField(t.pivotFilterItems, filterChecklist)
        : pivotEditField(t.pivotFilterItems, filterItems),
    );
    filterSettings.append(...filterConditionControls);
    filterSettings.appendChild(filterDialogButton);
    select.addEventListener('change', () => {
      settings.hidden = select.value !== String(PivotAxis.Value);
      filterSettings.hidden = select.value !== String(PivotAxis.Page);
    });
    valueFieldSettings.push({
      fieldIndex: index,
      axisSelect: select,
      aggregationSelect: aggregation,
      numberFormatInput: numberFormat,
      filterItemsInput: filterItems,
      filterChecklist: inferredItems.length > 0 ? filterChecklist : null,
      filterCondition: () => filterConditions.get(index),
    });
    row.appendChild(settings);
    row.appendChild(filterSettings);
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
      fields: valueFieldSettings.map((field) => ({
        fieldIndex: field.fieldIndex,
        axis: Number(field.axisSelect.value) as PivotAxis,
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
        dirtyFilterSettings.length === 0
          ? null
          : dirtyFilterSettings
              .map((field) =>
                pivotFilterConditionToSpec(
                  pivot.fields[field.fieldIndex] ?? '',
                  field.filterCondition(),
                ),
              )
              .filter((filter): filter is PivotFilterSpec => filter !== null),
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
