import {
  type SpreadsheetCompatibilityId,
  type SpreadsheetCompatibilityStatus,
  summarizeSpreadsheetCompatibility,
} from '../engine/compatibility.js';
import {
  listWorkbookObjects,
  summarizePassthroughs,
  summarizePivotTables,
  summarizeTables,
  WORKBOOK_OBJECT_KINDS,
  workbookObjectKindCounts,
} from '../engine/passthrough-sync.js';
import {
  PivotAggregation,
  PivotAxis,
  type PivotDataFieldSpec,
  type PivotFilterSpec,
  PivotReportLayout,
} from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import type { SessionIllustration } from '../store/store.js';
import { createDialogSelect } from '../toolbar/dialogs/form-controls.js';
import { projectDisabledState } from '../toolbar/menu-a11y.js';
import { appendDialogIconButton, createDialogButton } from './dialog-shell.js';
import {
  createPivotFilterConditionControls,
  type PivotFilterConditionState,
  pivotFilterConditionToSpec,
  pivotFilterSpecToCondition,
  showPivotFilterDialog,
} from './pivot-field-settings.js';

type WorkbookObjectsStrings = Strings['workbookObjects'];

export interface SpreadsheetCompatibilityReportItem {
  severity: 'info' | 'warning';
  label: string;
  detail: string;
}

const COMPATIBILITY_LABEL_KEYS: Record<
  SpreadsheetCompatibilityId,
  keyof WorkbookObjectsStrings['compatibilityLabels']
> = {
  'cell-formatting': 'cellFormatting',
  'conditional-formatting': 'conditionalFormatting',
  'data-validation': 'dataValidation',
  hyperlinks: 'hyperlinks',
  comments: 'comments',
  'defined-names': 'definedNames',
  'sheet-protection': 'sheetProtection',
  'sheet-views': 'sheetViews',
  'loaded-tables': 'loadedTables',
  'format-as-table': 'formatAsTable',
  'pivot-layouts': 'pivotLayouts',
  'pivot-authoring': 'pivotAuthoring',
  'session-charts': 'sessionCharts',
  'charts-drawings': 'chartsDrawings',
  'chart-authoring': 'chartAuthoring',
  'external-links': 'externalLinks',
};

const STATUS_LABEL_KEYS: Record<
  SpreadsheetCompatibilityStatus,
  keyof Pick<WorkbookObjectsStrings, 'writable' | 'readOnly' | 'sessionOnly' | 'unsupported'>
> = {
  writable: 'writable',
  'read-only': 'readOnly',
  session: 'sessionOnly',
  unsupported: 'unsupported',
};

export const spreadsheetCompatibilityLabel = (
  id: SpreadsheetCompatibilityId,
  strings: WorkbookObjectsStrings,
): string => strings.compatibilityLabels[COMPATIBILITY_LABEL_KEYS[id]];

export const spreadsheetCompatibilityDetail = (
  id: SpreadsheetCompatibilityId,
  strings: WorkbookObjectsStrings,
): string => strings.compatibilityDetails[COMPATIBILITY_LABEL_KEYS[id]];

export const spreadsheetCompatibilityStatusLabel = (
  status: SpreadsheetCompatibilityStatus,
  strings: WorkbookObjectsStrings,
): string => strings[STATUS_LABEL_KEYS[status]];

/** Build the flat severity/label/detail list rendered by the React, Vue, and
 *  playground "inspect workbook" backstage actions. Three call sites used to
 *  carry ~100 lines of switch-case duplication each; this helper is the single
 *  source of truth for that mapping. */
export const buildSpreadsheetCompatibilityReport = (
  wb: WorkbookHandle,
  strings: WorkbookObjectsStrings,
): SpreadsheetCompatibilityReportItem[] => {
  const summary = summarizeSpreadsheetCompatibility(wb);
  const banner: SpreadsheetCompatibilityReportItem = {
    severity: 'info',
    label: strings.compatibility,
    detail:
      `${strings.writable} ${summary.byStatus.writable}, ` +
      `${strings.readOnly} ${summary.byStatus['read-only']}, ` +
      `${strings.sessionOnly} ${summary.byStatus.session}, ` +
      `${strings.unsupported} ${summary.byStatus.unsupported}`,
  };
  const rows = summary.items.map<SpreadsheetCompatibilityReportItem>((entry) => {
    const detail = spreadsheetCompatibilityDetail(entry.id, strings);
    return {
      severity: entry.status === 'unsupported' || entry.status === 'read-only' ? 'warning' : 'info',
      label: `${spreadsheetCompatibilityLabel(entry.id, strings)} · ${spreadsheetCompatibilityStatusLabel(entry.status, strings)}`,
      detail: entry.count === undefined ? detail : `${detail} (${entry.count})`,
    };
  });
  return [banner, ...rows];
};

export interface WorkbookObjectsPanelDeps {
  host: HTMLElement;
  wb: WorkbookHandle;
  strings?: Strings;
  onOpenPivotTableDialog?: () => void;
  onAfterPivotEdit?: () => void;
  onSelectSessionIllustration?: (id: string) => void;
  onDuplicateSessionIllustration?: (id: string) => void;
  onClearSessionIllustration?: (id: string) => void;
  onUpdateSessionIllustration?: (
    id: string,
    patch: Partial<Omit<SessionIllustration, 'id'>>,
  ) => void;
  listSessionIllustrations?: () => readonly SessionIllustration[];
  subscribeSessionObjects?: (listener: () => void) => () => void;
}

export interface WorkbookObjectsPanelHandle {
  open(): void;
  openPivotFieldList(sheetIndex: number, pivotIndex: number): boolean;
  isPivotFieldListOpen(): boolean;
  close(): void;
  refresh(): void;
  setStrings(next: Strings): void;
  bindWorkbook(next: WorkbookHandle): void;
  detach(): void;
}

function createWorkbookObjectsActionButton(
  label: string,
  opts: { primary?: boolean; type?: 'button' | 'submit' } = {},
): HTMLButtonElement {
  const button = createDialogButton({
    label,
    baseClass: 'fc-objects__action',
    variant: opts.primary ? 'primary' : undefined,
    primaryClass: 'fc-objects__action--primary',
  });
  button.type = opts.type ?? 'button';
  return button;
}

const compatibilityLabelKey = (
  id: SpreadsheetCompatibilityId,
): keyof WorkbookObjectsStrings['compatibilityLabels'] => COMPATIBILITY_LABEL_KEYS[id];

const colLetter = (n: number): string => {
  let v = n;
  let out = '';
  do {
    out = String.fromCharCode(65 + (v % 26)) + out;
    v = Math.floor(v / 26) - 1;
  } while (v >= 0);
  return out;
};

const cellRef = (row: number, col: number): string => `${colLetter(col)}${row + 1}`;

const parseCellRef = (input: string): { row: number; col: number } | null => {
  const m = input
    .trim()
    .replace(/\$/g, '')
    .match(/^([A-Za-z]+)([1-9][0-9]*)$/);
  if (!m) return null;
  const letters = m[1];
  const rows = m[2];
  if (!letters || !rows) return null;
  let col = 0;
  for (const ch of letters.toUpperCase()) col = col * 26 + (ch.charCodeAt(0) - 64);
  return { row: Number(rows) - 1, col: col - 1 };
};

export function attachWorkbookObjectsPanel(
  deps: WorkbookObjectsPanelDeps,
): WorkbookObjectsPanelHandle {
  const { host } = deps;
  let wb = deps.wb;
  let strings = deps.strings ?? defaultStrings;
  let open = false;
  let restoreFocusEl: HTMLElement | null = null;
  let activePivotEditKey = '';
  let activePivotFieldListKey = '';
  let pivotEditError = '';

  const root = document.createElement('div');
  root.className = 'fc-objects';
  root.setAttribute('role', 'dialog');
  root.setAttribute('aria-modal', 'false');
  root.hidden = true;
  root.tabIndex = -1;
  host.appendChild(root);

  const close = (restoreFocus = false): void => {
    const wasOpen = open;
    open = false;
    const focusTarget = restoreFocusEl;
    restoreFocusEl = null;
    root.hidden = true;
    if (
      wasOpen &&
      restoreFocus &&
      focusTarget &&
      (root.contains(document.activeElement) || document.activeElement === document.body)
    ) {
      focusTarget.focus({ preventScroll: true });
    }
  };

  const item = (label: string, value: string | number): HTMLDivElement => {
    const row = document.createElement('div');
    row.className = 'fc-objects__row';
    const k = document.createElement('span');
    k.className = 'fc-objects__key';
    k.textContent = label;
    const v = document.createElement('span');
    v.className = 'fc-objects__value';
    v.textContent = String(value);
    row.append(k, v);
    return row;
  };

  const pivotAnchor = (pivot: {
    top: number;
    left: number;
    rows: number;
    cols: number;
    cells: number;
  }): string =>
    `R${pivot.top + 1}C${pivot.left + 1} · ${pivot.rows} x ${pivot.cols} · ${pivot.cells} ${strings.workbookObjects.cells}`;

  const fieldChips = (fields: readonly string[]): HTMLDivElement => {
    const wrap = document.createElement('div');
    wrap.className = 'fc-objects__chips';
    for (const field of fields) {
      const chip = document.createElement('span');
      chip.className = 'fc-objects__chip';
      chip.textContent = field;
      wrap.appendChild(chip);
    }
    return wrap;
  };

  const pivotKey = (sheet: number, index: number): string => `${sheet}:${index}`;

  const pivotEditField = (label: string, control: HTMLElement): HTMLLabelElement => {
    const row = document.createElement('label');
    row.className = 'fc-objects__pivot-edit-field';
    const text = document.createElement('span');
    text.textContent = label;
    row.append(text, control);
    return row;
  };

  const pivotEditCheck = (label: string, control: HTMLInputElement): HTMLLabelElement => {
    const row = document.createElement('label');
    row.className = 'fc-objects__pivot-edit-check';
    row.append(control, document.createTextNode(label));
    return row;
  };

  const appendIllustrationActions = (
    li: HTMLLIElement,
    illustration: SessionIllustration,
  ): void => {
    const actions = document.createElement('div');
    actions.className = 'fc-objects__actions fc-objects__illustration-actions';
    const select = createWorkbookObjectsActionButton(strings.workbookObjects.objectSelect);
    select.addEventListener('click', () => deps.onSelectSessionIllustration?.(illustration.id));
    actions.appendChild(select);
    if (deps.onDuplicateSessionIllustration) {
      const duplicate = createWorkbookObjectsActionButton(strings.workbookObjects.objectDuplicate);
      duplicate.addEventListener('click', () =>
        deps.onDuplicateSessionIllustration?.(illustration.id),
      );
      actions.appendChild(duplicate);
    }
    if (deps.onClearSessionIllustration) {
      const remove = createWorkbookObjectsActionButton(strings.workbookObjects.objectDelete);
      remove.addEventListener('click', () => deps.onClearSessionIllustration?.(illustration.id));
      actions.appendChild(remove);
    }
    li.appendChild(actions);
  };

  const renderIllustrationEditForm = (illustration: SessionIllustration): HTMLFormElement => {
    const t = strings.workbookObjects;
    const form = document.createElement('form');
    form.className = 'fc-objects__pivot-edit fc-objects__illustration-edit';
    const color = document.createElement('input');
    color.className = 'fc-objects__input fc-objects__color-input';
    color.type = 'color';
    color.value = illustration.color ?? '#0f6cbd';
    const radius = document.createElement('input');
    radius.className = 'fc-objects__input';
    radius.type = 'number';
    radius.min = '0';
    radius.max = '48';
    radius.step = '1';
    radius.value = String(
      illustration.radius ?? (illustration.shape === 'rounded-rectangle' ? 12 : 0),
    );
    const lineWidth = document.createElement('input');
    lineWidth.className = 'fc-objects__input';
    lineWidth.type = 'number';
    lineWidth.min = '1';
    lineWidth.max = '16';
    lineWidth.step = '1';
    lineWidth.value = String(illustration.lineWidth ?? 3);
    const opacity = document.createElement('input');
    opacity.className = 'fc-objects__input';
    opacity.type = 'range';
    opacity.min = '0';
    opacity.max = '1';
    opacity.step = '0.05';
    opacity.value = String(illustration.opacity ?? 0.16);
    const actions = document.createElement('div');
    actions.className = 'fc-objects__actions';
    const apply = createWorkbookObjectsActionButton(t.apply, { primary: true, type: 'submit' });
    actions.appendChild(apply);
    form.addEventListener('submit', (event) => {
      event.preventDefault();
      deps.onUpdateSessionIllustration?.(illustration.id, {
        color: color.value,
        radius: Math.max(0, Math.min(48, Number(radius.value) || 0)),
        lineWidth: Math.max(1, Math.min(16, Number(lineWidth.value) || 1)),
        opacity: Math.max(0, Math.min(1, Number(opacity.value) || 0)),
      });
    });
    form.append(
      pivotEditField(t.shapeColor, color),
      pivotEditField(t.shapeRadius, radius),
      pivotEditField(t.shapeLineWidth, lineWidth),
      pivotEditField(t.shapeOpacity, opacity),
      actions,
    );
    return form;
  };

  const renderPivotEditForm = (
    pivot: {
      sheetIndex: number;
      pivotIndex: number;
      top: number;
      left: number;
      rows: number;
      cols: number;
      fields: readonly string[];
      fieldItems?: Record<string, readonly string[]>;
    },
    opts: { fieldListOnly?: boolean } = {},
  ): HTMLFormElement => {
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
    anchor.value = cellRef(pivot.top, pivot.left);
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
    const fieldAreaSelects: HTMLSelectElement[] = [];
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
    for (const spec of (pivot as { pivotFilters?: readonly PivotFilterSpec[] }).pivotFilters ??
      []) {
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
      fieldAreaSelects.push(select);
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
          render();
        });
      });
      const filterChecklist = document.createElement('div');
      filterChecklist.className = 'fc-objects__pivot-filter-items';
      filterChecklist.dataset.pivotFilterChecklistFieldIndex = String(index);
      for (const itemName of inferredItems) {
        const check = document.createElement('input');
        check.type = 'checkbox';
        check.checked = true;
        check.value = itemName;
        const checkLabel = document.createElement('label');
        checkLabel.className = 'fc-objects__pivot-field-list-item';
        checkLabel.append(check, document.createTextNode(itemName));
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
        filterCondition: () => filterConditions.get(index),
      });
      row.appendChild(settings);
      row.appendChild(filterSettings);
      fieldAreas.appendChild(row);
    }
    const error = document.createElement('div');
    error.className = 'fc-objects__error';
    error.setAttribute('role', 'alert');
    error.hidden = !pivotEditError;
    error.textContent = pivotEditError;
    const actions = document.createElement('div');
    actions.className = 'fc-objects__actions';
    const remove = createWorkbookObjectsActionButton(t.deletePivotTable);
    const apply = createWorkbookObjectsActionButton(t.apply, { primary: true, type: 'submit' });
    if (fieldListOnly) actions.append(apply);
    else actions.append(remove, apply);
    remove.addEventListener('click', () => {
      pivotEditError = '';
      if (!wb.removePivotTable(pivot.sheetIndex, pivot.pivotIndex)) {
        pivotEditError = t.pivotEditFailed;
        render();
        return;
      }
      activePivotEditKey = '';
      deps.onAfterPivotEdit?.();
      render();
    });
    form.addEventListener('submit', (event) => {
      event.preventDefault();
      pivotEditError = '';
      const nextAnchor = fieldListOnly
        ? { row: pivot.top, col: pivot.left }
        : parseCellRef(anchor.value);
      if (!nextAnchor) {
        pivotEditError = t.invalidPivotAnchor;
        render();
        return;
      }
      const renamed =
        fieldListOnly || wb.renamePivotTable(pivot.sheetIndex, pivot.pivotIndex, name.value.trim());
      const moved =
        fieldListOnly ||
        wb.setPivotTableAnchor(pivot.sheetIndex, pivot.pivotIndex, {
          row: nextAnchor.row,
          col: nextAnchor.col,
          rows: pivot.rows,
          cols: pivot.cols,
        });
      const totaled =
        fieldListOnly ||
        wb.setPivotTableGrandTotals(
          pivot.sheetIndex,
          pivot.pivotIndex,
          rowTotals.checked,
          colTotals.checked,
        );
      const layoutUpdated =
        fieldListOnly ||
        wb.setPivotReportLayout(
          pivot.sheetIndex,
          pivot.pivotIndex,
          Number(layoutSelect.value) as PivotReportLayout,
        );
      const fieldsUpdated = fieldAreaSelects.every((select) =>
        wb.setPivotFieldAxis(
          pivot.sheetIndex,
          pivot.pivotIndex,
          Number(select.dataset.pivotFieldIndex),
          Number(select.value) as PivotAxis,
        ),
      );
      const dataFieldCount = wb.pivotDataFieldCount(pivot.sheetIndex, pivot.pivotIndex);
      let nextDataFieldIndex = 0;
      const valueFieldsUpdated = valueFieldSettings.every((field) => {
        if (field.axisSelect.value !== String(PivotAxis.Value)) return true;
        const format = field.numberFormatInput.value.trim();
        const spec: PivotDataFieldSpec = {
          fieldIndex: field.fieldIndex,
          aggregation: Number(field.aggregationSelect.value) as PivotAggregation,
          ...(format.length > 0 ? { numberFormat: format } : {}),
        };
        const dataFieldIndex = nextDataFieldIndex;
        nextDataFieldIndex += 1;
        if (dataFieldIndex < dataFieldCount) {
          return wb.setPivotDataField(pivot.sheetIndex, pivot.pivotIndex, dataFieldIndex, spec);
        }
        return wb.addPivotDataField(pivot.sheetIndex, pivot.pivotIndex, spec) >= 0;
      });
      const filterItemsUpdated = valueFieldSettings.every((field) => {
        if (field.axisSelect.value !== String(PivotAxis.Page)) return true;
        if (!wb.clearPivotFieldItems(pivot.sheetIndex, pivot.pivotIndex, field.fieldIndex)) {
          return false;
        }
        const items = field.filterItemsInput.value
          .split(/\r?\n/)
          .map((item) => item.trim())
          .filter(Boolean);
        const checklist = form.querySelector<HTMLElement>(
          `[data-pivot-filter-checklist-field-index="${field.fieldIndex}"]`,
        );
        if (checklist) {
          const checkedItems = Array.from(
            checklist.querySelectorAll<HTMLInputElement>('input[type="checkbox"]'),
          );
          return checkedItems.every((item) =>
            wb.addPivotFieldItem(
              pivot.sheetIndex,
              pivot.pivotIndex,
              field.fieldIndex,
              item.value,
              item.checked,
            ),
          );
        }
        return items.every((item) =>
          wb.addPivotFieldItem(pivot.sheetIndex, pivot.pivotIndex, field.fieldIndex, item, true),
        );
      });
      const dirtyFilterSettings = valueFieldSettings.filter(
        (field) =>
          field.axisSelect.value === String(PivotAxis.Page) &&
          filterConditionDirty.has(field.fieldIndex),
      );
      const nextPivotFilters = dirtyFilterSettings
        .map((field) =>
          pivotFilterConditionToSpec(pivot.fields[field.fieldIndex] ?? '', field.filterCondition()),
        )
        .filter((filter): filter is PivotFilterSpec => filter !== null);
      const pivotFiltersUpdated =
        dirtyFilterSettings.length === 0 ||
        (wb.clearPivotFilters(pivot.sheetIndex, pivot.pivotIndex) &&
          nextPivotFilters.every((filter) =>
            wb.addPivotFilter(pivot.sheetIndex, pivot.pivotIndex, filter),
          ));
      if (
        !renamed ||
        !moved ||
        !totaled ||
        !layoutUpdated ||
        !fieldsUpdated ||
        !valueFieldsUpdated ||
        !filterItemsUpdated ||
        !pivotFiltersUpdated
      ) {
        pivotEditError = t.pivotEditFailed;
        render();
        return;
      }
      activePivotEditKey = '';
      activePivotFieldListKey = '';
      deps.onAfterPivotEdit?.();
      render();
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
  };

  const render = (): void => {
    const t = strings.workbookObjects;
    const objects = listWorkbookObjects(wb);
    const passthroughs = summarizePassthroughs(wb);
    const tables = summarizeTables(wb);
    const pivots = summarizePivotTables(wb);
    const illustrations = deps.listSessionIllustrations?.() ?? [];
    const support = summarizeSpreadsheetCompatibility(wb);
    const activeFieldListPivot = activePivotFieldListKey
      ? pivots.items.find(
          (pivot) => pivotKey(pivot.sheetIndex, pivot.pivotIndex) === activePivotFieldListKey,
        )
      : undefined;
    if (activePivotFieldListKey && !activeFieldListPivot) activePivotFieldListKey = '';
    root.replaceChildren();
    root.className = `fc-objects${activeFieldListPivot ? ' fc-objects--taskpane' : ''}`;
    root.setAttribute('aria-label', activeFieldListPivot ? t.pivotFieldList : t.title);

    const header = document.createElement('div');
    header.className = 'fc-objects__header';
    const title = document.createElement('div');
    title.className = 'fc-objects__title';
    title.textContent = activeFieldListPivot ? t.pivotFieldList : t.title;
    const headerActions = document.createElement('div');
    headerActions.className = 'fc-objects__header-actions';
    if (activeFieldListPivot) {
      const back = createWorkbookObjectsActionButton(t.backToWorkbookObjects);
      back.addEventListener('click', () => {
        activePivotFieldListKey = '';
        pivotEditError = '';
        render();
      });
      headerActions.appendChild(back);
    }
    const closeBtn = appendDialogIconButton(headerActions, {
      label: '',
      ariaLabel: t.close,
      baseClass: 'fc-objects__close',
    });
    closeBtn.addEventListener('click', () => close(false));
    header.append(title, headerActions);
    root.appendChild(header);

    const body = document.createElement('div');
    body.className = 'fc-objects__body';
    if (activeFieldListPivot) {
      const section = document.createElement('section');
      section.className = 'fc-objects__section';
      const heading = document.createElement('div');
      heading.className = 'fc-objects__heading';
      heading.textContent = [
        `${t.pivot} ${activeFieldListPivot.pivotIndex + 1}`,
        `${t.sheet} ${activeFieldListPivot.sheetIndex + 1}`,
        pivotAnchor(activeFieldListPivot),
      ].join(' · ');
      section.append(heading, renderPivotEditForm(activeFieldListPivot, { fieldListOnly: true }));
      body.appendChild(section);
      root.appendChild(body);
      return;
    }
    const summary = document.createElement('section');
    summary.className = 'fc-objects__section';
    summary.append(
      item(t.preservedParts, passthroughs.count),
      item(t.tables, tables.count),
      item(t.pivotTables, pivots.count),
      item(strings.ribbon.illustrations, illustrations.length),
      item(t.writable, support.byStatus.writable),
      item(t.readOnly, support.byStatus['read-only']),
      item(t.sessionOnly, support.byStatus.session),
      item(t.unsupported, support.byStatus.unsupported),
      item(t.noteLabel, t.readOnlyNote),
    );
    body.appendChild(summary);

    const supportSection = document.createElement('section');
    supportSection.className = 'fc-objects__section';
    const supportHeading = document.createElement('div');
    supportHeading.className = 'fc-objects__heading';
    supportHeading.textContent = t.compatibility;
    supportSection.appendChild(supportHeading);
    const supportList = document.createElement('ul');
    supportList.className = 'fc-objects__paths';
    for (const entry of support.items) {
      const li = document.createElement('li');
      li.textContent = [
        t.compatibilityLabels[compatibilityLabelKey(entry.id)],
        t[
          entry.status === 'read-only'
            ? 'readOnly'
            : entry.status === 'session'
              ? 'sessionOnly'
              : entry.status
        ],
        entry.count === undefined ? '' : `${entry.count}`,
      ]
        .filter(Boolean)
        .join(' · ');
      supportList.appendChild(li);
    }
    supportSection.appendChild(supportList);
    body.appendChild(supportSection);

    const objectCounts = workbookObjectKindCounts(objects);
    const cats = WORKBOOK_OBJECT_KINDS.filter((kind) => objectCounts[kind] > 0);
    if (cats.length > 0) {
      const section = document.createElement('section');
      section.className = 'fc-objects__section';
      const heading = document.createElement('div');
      heading.className = 'fc-objects__heading';
      heading.textContent = t.categories;
      section.appendChild(heading);
      for (const category of cats) {
        section.appendChild(item(t.kindLabels[category], objectCounts[category]));
      }
      body.appendChild(section);
    }

    if (tables.names.length > 0) {
      const section = document.createElement('section');
      section.className = 'fc-objects__section';
      const heading = document.createElement('div');
      heading.className = 'fc-objects__heading';
      heading.textContent = t.tableNames;
      section.appendChild(heading);
      const list = document.createElement('div');
      list.className = 'fc-objects__list';
      list.textContent = tables.names.join(', ');
      section.appendChild(list);
      body.appendChild(section);
    }

    if (tables.items.length > 0) {
      const section = document.createElement('section');
      section.className = 'fc-objects__section';
      const heading = document.createElement('div');
      heading.className = 'fc-objects__heading';
      heading.textContent = t.tableDetails;
      section.appendChild(heading);
      const list = document.createElement('ul');
      list.className = 'fc-objects__paths';
      for (const table of tables.items) {
        const li = document.createElement('li');
        const name = table.displayName || table.name;
        const cols = table.columns.length;
        li.textContent = [
          name,
          `${t.sheet} ${table.sheetIndex + 1}`,
          table.ref,
          `${cols} ${cols === 1 ? t.columnSingular : t.columnPlural}`,
        ].join(' · ');
        list.appendChild(li);
      }
      section.appendChild(list);
      body.appendChild(section);
    }

    if (pivots.items.length > 0) {
      const section = document.createElement('section');
      section.className = 'fc-objects__section';
      const heading = document.createElement('div');
      heading.className = 'fc-objects__heading';
      heading.textContent = t.pivotDetails;
      section.appendChild(heading);
      for (const pivot of pivots.items) {
        const card = document.createElement('div');
        card.className = 'fc-objects__pivot-card';
        const titleRow = document.createElement('div');
        titleRow.className = 'fc-objects__pivot-title';
        const title = document.createElement('strong');
        title.textContent = `${t.pivot} ${pivot.pivotIndex + 1}`;
        const meta = document.createElement('span');
        meta.textContent = `${t.sheet} ${pivot.sheetIndex + 1} · ${pivotAnchor(pivot)}`;
        titleRow.append(title, meta);
        card.appendChild(titleRow);
        if (pivot.fields.length > 0) card.appendChild(fieldChips(pivot.fields));
        const canEditPivot = wb.capabilities.pivotTableMutate;
        if (deps.onOpenPivotTableDialog) {
          const actions = document.createElement('div');
          actions.className = 'fc-objects__actions';
          if (canEditPivot) {
            const edit = createWorkbookObjectsActionButton(t.editPivotTable);
            edit.addEventListener('click', () => {
              const key = pivotKey(pivot.sheetIndex, pivot.pivotIndex);
              activePivotEditKey = activePivotEditKey === key ? '' : key;
              activePivotFieldListKey = '';
              pivotEditError = '';
              render();
            });
            actions.appendChild(edit);
          }
          if (canEditPivot) {
            const fieldList = createWorkbookObjectsActionButton(t.pivotFieldList);
            fieldList.addEventListener('click', () => {
              const key = pivotKey(pivot.sheetIndex, pivot.pivotIndex);
              activePivotFieldListKey = activePivotFieldListKey === key ? '' : key;
              activePivotEditKey = '';
              pivotEditError = '';
              render();
            });
            actions.appendChild(fieldList);
          }
          const button = createWorkbookObjectsActionButton(t.createPivotTable);
          button.addEventListener('click', () => deps.onOpenPivotTableDialog?.());
          actions.appendChild(button);
          card.appendChild(actions);
        } else if (canEditPivot) {
          const actions = document.createElement('div');
          actions.className = 'fc-objects__actions';
          const edit = createWorkbookObjectsActionButton(t.editPivotTable);
          edit.addEventListener('click', () => {
            const key = pivotKey(pivot.sheetIndex, pivot.pivotIndex);
            activePivotEditKey = activePivotEditKey === key ? '' : key;
            activePivotFieldListKey = '';
            pivotEditError = '';
            render();
          });
          actions.appendChild(edit);
          const fieldList = createWorkbookObjectsActionButton(t.pivotFieldList);
          fieldList.addEventListener('click', () => {
            const key = pivotKey(pivot.sheetIndex, pivot.pivotIndex);
            activePivotFieldListKey = activePivotFieldListKey === key ? '' : key;
            activePivotEditKey = '';
            pivotEditError = '';
            render();
          });
          actions.appendChild(fieldList);
          card.appendChild(actions);
        }
        if (canEditPivot && activePivotEditKey === pivotKey(pivot.sheetIndex, pivot.pivotIndex)) {
          card.appendChild(renderPivotEditForm(pivot));
        }
        if (
          canEditPivot &&
          activePivotFieldListKey === pivotKey(pivot.sheetIndex, pivot.pivotIndex)
        ) {
          card.appendChild(renderPivotEditForm(pivot, { fieldListOnly: true }));
        }
        section.appendChild(card);
      }
      body.appendChild(section);
    }

    if (illustrations.length > 0) {
      const section = document.createElement('section');
      section.className = 'fc-objects__section';
      const heading = document.createElement('div');
      heading.className = 'fc-objects__heading';
      heading.textContent = strings.ribbon.illustrations;
      section.appendChild(heading);
      const list = document.createElement('ul');
      list.className = 'fc-objects__paths';
      for (const [index, illustration] of illustrations.entries()) {
        const li = document.createElement('li');
        li.textContent = [
          illustration.kind === 'image'
            ? strings.ribbon.pictures
            : (illustration.shape ?? strings.ribbon.shapes),
          `${t.sheet} ${illustration.sheet + 1}`,
          illustration.id || `${strings.ribbon.illustrations} ${index + 1}`,
        ].join(' · ');
        if (illustration.src) li.title = illustration.src;
        appendIllustrationActions(li, illustration);
        if (illustration.kind === 'shape') li.appendChild(renderIllustrationEditForm(illustration));
        list.appendChild(li);
      }
      section.appendChild(list);
      body.appendChild(section);
    }

    if (objects.length > 0) {
      const section = document.createElement('section');
      section.className = 'fc-objects__section';
      const heading = document.createElement('div');
      heading.className = 'fc-objects__heading';
      heading.textContent = t.paths;
      section.appendChild(heading);
      const list = document.createElement('ul');
      list.className = 'fc-objects__paths';
      for (const object of objects.slice(0, 32)) {
        const li = document.createElement('li');
        li.textContent = `${t.kindLabels[object.kind]} · ${object.path}`;
        li.title = object.path;
        list.appendChild(li);
      }
      section.appendChild(list);
      body.appendChild(section);
    }

    if (
      passthroughs.count === 0 &&
      tables.count === 0 &&
      pivots.count === 0 &&
      illustrations.length === 0
    ) {
      const empty = document.createElement('div');
      empty.className = 'fc-objects__empty';
      empty.textContent = t.empty;
      body.appendChild(empty);
    }
    root.appendChild(body);
  };

  const refresh = (): void => {
    if (open) render();
  };
  const unsubscribeSessionObjects = deps.subscribeSessionObjects?.(refresh) ?? null;

  const openPanel = (): void => {
    render();
    restoreFocusEl = document.activeElement instanceof HTMLElement ? document.activeElement : host;
    root.hidden = false;
    open = true;
    root.focus({ preventScroll: true });
  };

  const openPivotFieldList = (sheetIndex: number, pivotIndex: number): boolean => {
    const key = pivotKey(sheetIndex, pivotIndex);
    if (
      !summarizePivotTables(wb).items.some(
        (pivot) => pivotKey(pivot.sheetIndex, pivot.pivotIndex) === key,
      )
    ) {
      return false;
    }
    activePivotEditKey = '';
    activePivotFieldListKey = key;
    pivotEditError = '';
    openPanel();
    return true;
  };

  const onKey = (e: KeyboardEvent): void => {
    if (e.key === 'Escape') close(true);
  };
  root.addEventListener('keydown', onKey);

  return {
    open: openPanel,
    openPivotFieldList,
    isPivotFieldListOpen: () => open && activePivotFieldListKey.length > 0,
    close,
    refresh,
    setStrings(next) {
      strings = next;
      refresh();
    },
    bindWorkbook(next) {
      wb = next;
      refresh();
    },
    detach() {
      unsubscribeSessionObjects?.();
      root.removeEventListener('keydown', onKey);
      root.remove();
    },
  };
}
