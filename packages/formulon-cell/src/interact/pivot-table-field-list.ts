import type { PivotSourceField } from '../commands/pivot-table.js';
import type { Strings } from '../i18n/strings.js';
import { createExcelRibbonSvg } from '../toolbar/excel-ribbon-icons.js';
import {
  createPivotAreaSettingsButton,
  type PivotAreaKind,
  type PivotFieldSettingsPanelOptions,
  renderPivotFieldSettingsPanel,
} from './pivot-field-settings.js';
import type { PivotFieldAreas } from './pivot-table-field-areas.js';

/** Field placement as the field list and its four areas read and change it. */
export interface PivotFieldAreasModel {
  /** Every source field, in source order. */
  readonly fields: readonly string[];
  /** Field names on `kind`, in display order. */
  fieldsIn(kind: PivotAreaKind): readonly string[];
  /** Text shown on a field's chip in `kind`. */
  chipLabel?(kind: PivotAreaKind, fieldName: string): string;
  /** Captions of fixed entries listed after the fields of `kind`, such as the
   *  Values field that stands for the data fields; not draggable. */
  fixedChips?(kind: PivotAreaKind): readonly string[];
  canPlace(fieldName: string, kind: PivotAreaKind): boolean;
  /** Field-list checkbox: ticking places the field by its data type, unticking takes it out. */
  toggle(fieldName: string, checked: boolean): void;
  /** Puts `fieldName` on `kind`, before `before` or last when null. */
  place(fieldName: string, kind: PivotAreaKind, before: string | null): void;
  /** Takes `fieldName` out of every area. */
  remove(fieldName: string): void;
}

export interface PivotFieldAreasLabels {
  /** Heads the field list; sits beside the search box when there is one. */
  availableFields: string;
  /** Heads the area grid; no heading when omitted. */
  fieldAreas?: string;
  filters: string;
  columns: string;
  rows: string;
  values: string;
  /** Shown in an empty area; an empty area stays blank when omitted. */
  none?: string;
  fieldSettings: string;
  /** Accessible name of a chip's settings button; `{field}` is the field name. */
  fieldSettingsFor: string;
  /** Placeholder of the field search box; no search box when omitted. */
  search?: string;
  /** Hint below the area grid. */
  areaHint?: string;
}

export interface PivotFieldAreasViewOptions {
  model: PivotFieldAreasModel;
  labels: PivotFieldAreasLabels;
  /** Runs after every model change; re-renders the view. */
  refresh(): void;
  /** A chip in an area was activated. */
  onFieldSettings(kind: PivotAreaKind, fieldName: string): void;
  /** Field search text, kept by the caller across re-renders. */
  searchText?: string;
  onSearch?(text: string): void;
  /** Appended below the area grid. */
  areaFooter?: readonly HTMLElement[];
}

export const PIVOT_AREA_ORDER: readonly PivotAreaKind[] = ['filters', 'columns', 'rows', 'values'];

const AREA_ICON: Record<PivotAreaKind, string> = {
  filters: 'filter',
  columns: 'pivotAreaColumns',
  rows: 'pivotAreaRows',
  values: 'autosum',
};

interface PivotDrag {
  fieldName: string;
  /** The area the dragged chip came from; null for a field-list row. */
  from: PivotAreaKind | null;
  dropped: boolean;
}

const pivotDrags = new WeakMap<PivotFieldAreasModel, PivotDrag>();
const PIVOT_DRAG_TYPE = 'application/x-fc-pivot-field';

const pivotDragData = (event: DragEvent, model: PivotFieldAreasModel): string =>
  event.dataTransfer?.getData('text/plain') ||
  event.dataTransfer?.getData(PIVOT_DRAG_TYPE) ||
  pivotDrags.get(model)?.fieldName ||
  '';

const startPivotDrag = (
  event: DragEvent,
  model: PivotFieldAreasModel,
  fieldName: string,
  from: PivotAreaKind | null,
): void => {
  pivotDrags.set(model, { fieldName, from, dropped: false });
  event.dataTransfer?.setData('text/plain', fieldName);
  event.dataTransfer?.setData(PIVOT_DRAG_TYPE, fieldName);
  if (event.dataTransfer) event.dataTransfer.effectAllowed = 'move';
};

/** Ends a drag; a chip dropped outside every area leaves the report. */
const endPivotDrag = (model: PivotFieldAreasModel, refresh: () => void): void => {
  const drag = pivotDrags.get(model);
  pivotDrags.delete(model);
  if (!drag || drag.dropped || drag.from === null) return;
  model.remove(drag.fieldName);
  refresh();
};

const markPivotDropped = (model: PivotFieldAreasModel): void => {
  const drag = pivotDrags.get(model);
  if (drag) drag.dropped = true;
};

/** Renders the field list (checkboxes) and the four area drop targets into `container`. */
export function renderPivotFieldAreas(
  container: HTMLElement,
  opts: PivotFieldAreasViewOptions,
): void {
  const { model, labels, refresh } = opts;
  const assigned = new Set(PIVOT_AREA_ORDER.flatMap((kind) => model.fieldsIn(kind)));

  const available = document.createElement('div');
  available.className = 'fc-pivotdlg__field-list-available';
  const availableLabel = document.createElement('div');
  availableLabel.className = 'fc-pivotdlg__field-list-label';
  availableLabel.textContent = labels.availableFields;
  const fieldGrid = document.createElement('div');
  fieldGrid.className = 'fc-pivotdlg__field-list-grid';
  const rows: { name: string; label: HTMLLabelElement }[] = [];
  for (const fieldName of model.fields) {
    const label = document.createElement('label');
    label.className = 'fc-pivotdlg__field-chip';
    const input = document.createElement('input');
    input.type = 'checkbox';
    input.checked = assigned.has(fieldName);
    input.dataset.pivotFieldListField = fieldName;
    input.addEventListener('change', () => {
      model.toggle(fieldName, input.checked);
      refresh();
    });
    const name = document.createElement('span');
    name.textContent = fieldName;
    label.draggable = true;
    label.addEventListener('dragstart', (event) => startPivotDrag(event, model, fieldName, null));
    label.addEventListener('dragend', () => endPivotDrag(model, refresh));
    label.append(input, name);
    fieldGrid.appendChild(label);
    rows.push({ name: fieldName, label });
  }
  if (labels.search !== undefined) {
    const search = document.createElement('input');
    search.type = 'search';
    search.className = 'fc-pivotdlg__field-search';
    search.placeholder = labels.search;
    search.setAttribute('aria-label', labels.search);
    search.dataset.pivotFieldSearch = '';
    search.value = opts.searchText ?? '';
    const applySearch = (): void => {
      const needle = search.value.trim().toLocaleLowerCase();
      for (const row of rows) {
        row.label.hidden = needle.length > 0 && !row.name.toLocaleLowerCase().includes(needle);
      }
    };
    search.addEventListener('input', () => {
      applySearch();
      opts.onSearch?.(search.value);
    });
    applySearch();
    const searchBox = document.createElement('span');
    searchBox.className = 'fc-pivotdlg__field-search-box';
    const lens = createExcelRibbonSvg('find', 'fc-pivotdlg__field-search-icon');
    if (lens) searchBox.appendChild(lens);
    searchBox.appendChild(search);
    const searchRow = document.createElement('div');
    searchRow.className = 'fc-pivotdlg__field-search-row';
    searchRow.append(availableLabel, searchBox);
    available.appendChild(searchRow);
  } else {
    available.appendChild(availableLabel);
  }
  available.appendChild(fieldGrid);
  // Dropping a chip back on the field list takes it out of the report.
  available.addEventListener('dragover', (event) => {
    if (pivotDrags.get(model)?.from) event.preventDefault();
  });
  available.addEventListener('drop', (event) => {
    if (!pivotDrags.get(model)?.from) return;
    event.preventDefault();
  });

  const areaWrap = document.createElement('div');
  areaWrap.className = 'fc-pivotdlg__areas';
  if (labels.fieldAreas !== undefined) {
    const areasLabel = document.createElement('div');
    areasLabel.className = 'fc-pivotdlg__field-list-label';
    areasLabel.textContent = labels.fieldAreas;
    areaWrap.appendChild(areasLabel);
  }
  const areaGrid = document.createElement('div');
  areaGrid.className = 'fc-pivotdlg__area-grid';
  const areaLabel: Record<PivotAreaKind, string> = {
    filters: labels.filters,
    columns: labels.columns,
    rows: labels.rows,
    values: labels.values,
  };
  const area = (kind: PivotAreaKind): HTMLDivElement => {
    const wrap = document.createElement('div');
    wrap.className = 'fc-pivotdlg__area';
    wrap.dataset.pivotArea = kind;
    wrap.addEventListener('dragover', (event) => {
      const fieldName = pivotDragData(event, model);
      if (!fieldName || !model.canPlace(fieldName, kind)) return;
      event.preventDefault();
      wrap.dataset.pivotDragOver = 'true';
    });
    wrap.addEventListener('dragleave', () => {
      delete wrap.dataset.pivotDragOver;
    });
    wrap.addEventListener('drop', (event) => {
      const fieldName = pivotDragData(event, model);
      if (!fieldName || !model.canPlace(fieldName, kind)) return;
      event.preventDefault();
      delete wrap.dataset.pivotDragOver;
      markPivotDropped(model);
      const target = event.target instanceof Element ? event.target : null;
      const before =
        target?.closest<HTMLElement>('[data-pivot-area-field]')?.dataset.pivotAreaField ?? null;
      if (before === fieldName) return;
      model.place(fieldName, kind, before);
      refresh();
    });
    const heading = document.createElement('span');
    const icon = createExcelRibbonSvg(AREA_ICON[kind], 'fc-pivotdlg__area-icon');
    if (icon) heading.appendChild(icon);
    heading.append(areaLabel[kind]);
    const list = document.createElement('div');
    list.className = 'fc-pivotdlg__area-fields';
    const present = model.fieldsIn(kind).filter(Boolean);
    if (present.length === 0 && labels.none !== undefined) {
      const none = document.createElement('strong');
      none.textContent = labels.none;
      list.appendChild(none);
    }
    for (const value of present) {
      const chip = document.createElement('div');
      chip.className = 'fc-pivotdlg__area-field';
      chip.draggable = true;
      chip.dataset.pivotAreaField = value;
      chip.addEventListener('dragstart', (event) => startPivotDrag(event, model, value, kind));
      chip.addEventListener('dragend', () => endPivotDrag(model, refresh));
      const name = document.createElement('strong');
      name.textContent = model.chipLabel?.(kind, value) ?? value;
      const settings = createPivotAreaSettingsButton(
        labels.fieldSettings,
        labels.fieldSettingsFor.replace('{field}', value),
      );
      chip.addEventListener('click', () => opts.onFieldSettings(kind, value));
      chip.append(name, settings);
      list.appendChild(chip);
    }
    for (const caption of model.fixedChips?.(kind) ?? []) {
      const chip = document.createElement('div');
      chip.className = 'fc-pivotdlg__area-field fc-pivotdlg__area-field--fixed';
      chip.dataset.pivotAreaFixed = caption;
      const name = document.createElement('strong');
      name.textContent = caption;
      chip.appendChild(name);
      list.appendChild(chip);
    }
    wrap.append(heading, list);
    return wrap;
  };
  areaGrid.append(...PIVOT_AREA_ORDER.map(area));
  areaWrap.appendChild(areaGrid);
  if (labels.areaHint !== undefined) {
    const hint = document.createElement('div');
    hint.className = 'fc-pivotdlg__area-hint';
    hint.textContent = labels.areaHint;
    areaWrap.appendChild(hint);
  }
  areaWrap.append(...(opts.areaFooter ?? []));
  container.append(available, areaWrap);
}

export interface PivotFieldListContext {
  host: HTMLElement;
  readonly strings: Strings;
  areas: PivotFieldAreas;
  controls: PivotFieldSettingsPanelOptions['controls'];
  fields: readonly PivotSourceField[];
  inferFilterItems(fieldName: string): readonly string[];
}

const renderFieldSettingsPanel = (
  panelEl: HTMLDivElement,
  ctx: PivotFieldListContext,
  refreshFieldList: () => void,
): void => {
  const { areas, fields } = ctx;
  renderPivotFieldSettingsPanel({
    host: ctx.host,
    panelEl,
    active: areas.active,
    strings: ctx.strings.pivotTableDialog,
    okLabel: ctx.strings.pageSetup.ok,
    cancelLabel: ctx.strings.pageSetup.cancel,
    fields,
    controls: ctx.controls,
    selectedValueFields: areas.valueFields,
    selectedFilterFields: areas.filterFields,
    selectedValueSetting: areas.selectedValueSetting,
    setValueFieldSetting: areas.setValueFieldSetting,
    fieldCanBeValue: areas.fieldCanBeValue,
    replaceFilterField: areas.replaceFilterField,
    normalizeSelectedFilters: areas.normalizeSelectedFilters,
    refreshFieldList,
    filterItemVisibility: areas.filterItemVisibility,
    setFilterItemVisibility: areas.setFilterItemVisibility,
    selectedFilterCondition: areas.selectedFilterCondition,
    setFilterCondition: (fieldName, condition) =>
      areas.setFilterCondition(fieldName, {
        kind: condition.kind,
        value: condition.value,
      }),
    inferFilterItems: ctx.inferFilterItems,
  });
};

/** The Create PivotTable dialog's field list: the shared areas view over the
 *  dialog's single row/column selects and its filter and value lists. */
export function renderPivotFieldList(container: HTMLElement, ctx: PivotFieldListContext): void {
  const { areas, controls, fields } = ctx;
  const t = ctx.strings.pivotTableDialog;
  container.replaceChildren();
  const refresh = (): void => renderPivotFieldList(container, ctx);
  const model: PivotFieldAreasModel = {
    fields: fields.map((field) => field.name),
    fieldsIn: (kind) =>
      kind === 'filters'
        ? areas.filterFields
        : kind === 'values'
          ? areas.valueFields
          : [(kind === 'rows' ? controls.rowSelect : controls.colSelect).value].filter(Boolean),
    canPlace: (fieldName, kind) => kind !== 'values' || areas.fieldCanBeValue(fieldName),
    toggle: (fieldName, checked) => areas.toggleFromFieldList(fieldName, checked, fields),
    place: (fieldName, kind) => areas.assignFieldToArea(fieldName, kind, fields),
    remove: (fieldName) => areas.toggleFromFieldList(fieldName, false, fields),
  };

  const title = document.createElement('div');
  title.className = 'fc-pivotdlg__field-list-title';
  title.textContent = t.fieldList;
  container.appendChild(title);
  const settingsPanel = document.createElement('div');
  settingsPanel.className = 'fc-pivotdlg__area-settings-panel';
  settingsPanel.hidden = true;
  settingsPanel.setAttribute('role', 'status');
  settingsPanel.setAttribute('aria-live', 'polite');
  renderPivotFieldAreas(container, {
    model,
    labels: {
      availableFields: t.availableFields,
      fieldAreas: t.fieldAreas,
      filters: t.filtersArea,
      columns: t.columnField,
      rows: t.rowField,
      values: t.valueField,
      none: t.none,
      fieldSettings: t.fieldSettings,
      fieldSettingsFor: t.fieldSettingsFor,
    },
    refresh,
    onFieldSettings: (kind, fieldName) => {
      areas.setActive({ kind, fieldName });
      renderFieldSettingsPanel(settingsPanel, ctx, refresh);
      settingsPanel.querySelector<HTMLElement>('select, input')?.focus();
    },
    areaFooter: [settingsPanel],
  });
  areas.pruneActiveSettings();
  renderFieldSettingsPanel(settingsPanel, ctx, refresh);
}
