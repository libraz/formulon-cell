import type { PivotSourceField } from '../commands/pivot-table.js';
import type { Strings } from '../i18n/strings.js';
import {
  createPivotAreaSettingsButton,
  type PivotAreaKind,
  type PivotFieldSettingsPanelOptions,
  renderPivotFieldSettingsPanel,
} from './pivot-field-settings.js';
import type { PivotFieldAreas } from './pivot-table-field-areas.js';

export interface PivotFieldListContext {
  host: HTMLElement;
  readonly strings: Strings;
  areas: PivotFieldAreas;
  controls: PivotFieldSettingsPanelOptions['controls'];
  fields: readonly PivotSourceField[];
  inferFilterItems(fieldName: string): readonly string[];
}

const draggedPivotFields = new WeakMap<PivotFieldAreas, string>();

const pivotDragData = (event: DragEvent, areas: PivotFieldAreas): string =>
  event.dataTransfer?.getData('text/plain') ||
  event.dataTransfer?.getData('application/x-fc-pivot-field') ||
  draggedPivotFields.get(areas) ||
  '';

const setPivotDragData = (event: DragEvent, areas: PivotFieldAreas, fieldName: string): void => {
  draggedPivotFields.set(areas, fieldName);
  event.dataTransfer?.setData('text/plain', fieldName);
  event.dataTransfer?.setData('application/x-fc-pivot-field', fieldName);
  if (event.dataTransfer) event.dataTransfer.effectAllowed = 'move';
};

const clearPivotDragData = (areas: PivotFieldAreas): void => {
  draggedPivotFields.set(areas, '');
};

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

/** Renders the available-fields grid and the four area drop targets into `container`. */
export function renderPivotFieldList(container: HTMLElement, ctx: PivotFieldListContext): void {
  const { areas, controls, fields } = ctx;
  const t = ctx.strings.pivotTableDialog;
  const assigned = new Set(
    [
      controls.rowSelect.value,
      controls.colSelect.value,
      ...areas.filterFields,
      ...areas.valueFields,
    ].filter(Boolean),
  );
  container.replaceChildren();
  const refresh = (): void => renderPivotFieldList(container, ctx);

  const title = document.createElement('div');
  title.className = 'fc-pivotdlg__field-list-title';
  title.textContent = t.fieldList;
  const available = document.createElement('div');
  available.className = 'fc-pivotdlg__field-list-available';
  const availableLabel = document.createElement('div');
  availableLabel.className = 'fc-pivotdlg__field-list-label';
  availableLabel.textContent = t.availableFields;
  const fieldGrid = document.createElement('div');
  fieldGrid.className = 'fc-pivotdlg__field-list-grid';
  for (const field of fields) {
    const label = document.createElement('label');
    label.className = 'fc-pivotdlg__field-chip';
    const input = document.createElement('input');
    input.type = 'checkbox';
    input.checked = assigned.has(field.name);
    input.dataset.pivotFieldListField = field.name;
    input.addEventListener('change', () => {
      areas.toggleFromFieldList(field.name, input.checked, fields);
      refresh();
    });
    const name = document.createElement('span');
    name.textContent = field.name;
    label.draggable = true;
    label.addEventListener('dragstart', (event) => setPivotDragData(event, areas, field.name));
    label.addEventListener('dragend', () => clearPivotDragData(areas));
    label.append(input, name);
    fieldGrid.appendChild(label);
  }
  available.append(availableLabel, fieldGrid);

  const areaWrap = document.createElement('div');
  areaWrap.className = 'fc-pivotdlg__areas';
  const areasLabel = document.createElement('div');
  areasLabel.className = 'fc-pivotdlg__field-list-label';
  areasLabel.textContent = t.fieldAreas;
  const areaGrid = document.createElement('div');
  areaGrid.className = 'fc-pivotdlg__area-grid';
  const settingsPanel = document.createElement('div');
  settingsPanel.className = 'fc-pivotdlg__area-settings-panel';
  settingsPanel.hidden = true;
  settingsPanel.setAttribute('role', 'status');
  settingsPanel.setAttribute('aria-live', 'polite');
  const showFieldSettings = (kind: PivotAreaKind, fieldName: string): void => {
    areas.setActive({ kind, fieldName });
    renderFieldSettingsPanel(settingsPanel, ctx, refresh);
    settingsPanel.querySelector<HTMLElement>('select, input')?.focus();
  };
  const area = (label: string, values: readonly string[], kind: PivotAreaKind): HTMLDivElement => {
    const wrap = document.createElement('div');
    wrap.className = 'fc-pivotdlg__area';
    wrap.dataset.pivotArea = kind;
    wrap.addEventListener('dragover', (event) => {
      const fieldName = pivotDragData(event, areas);
      if (!fieldName) return;
      if (kind === 'values' && !areas.fieldCanBeValue(fieldName)) return;
      event.preventDefault();
      wrap.dataset.pivotDragOver = 'true';
    });
    wrap.addEventListener('dragleave', () => {
      delete wrap.dataset.pivotDragOver;
    });
    wrap.addEventListener('drop', (event) => {
      const fieldName = pivotDragData(event, areas);
      if (!fieldName) return;
      event.preventDefault();
      delete wrap.dataset.pivotDragOver;
      areas.assignFieldToArea(fieldName, kind, fields);
      refresh();
    });
    const heading = document.createElement('span');
    heading.textContent = label;
    const list = document.createElement('div');
    list.className = 'fc-pivotdlg__area-fields';
    const present = values.filter(Boolean);
    if (present.length === 0) {
      const none = document.createElement('strong');
      none.textContent = t.none;
      list.appendChild(none);
    } else {
      for (const value of present) {
        const chip = document.createElement('div');
        chip.className = 'fc-pivotdlg__area-field';
        chip.draggable = true;
        chip.addEventListener('dragstart', (event) => setPivotDragData(event, areas, value));
        chip.addEventListener('dragend', () => clearPivotDragData(areas));
        const name = document.createElement('strong');
        name.textContent = value;
        const settings = createPivotAreaSettingsButton(
          t.fieldSettings,
          t.fieldSettingsFor.replace('{field}', value),
        );
        settings.addEventListener('click', () => showFieldSettings(kind, value));
        chip.append(name, settings);
        list.appendChild(chip);
      }
    }
    wrap.append(heading, list);
    return wrap;
  };
  areaGrid.append(
    area(t.filtersArea, areas.filterFields, 'filters'),
    area(t.columnField, [controls.colSelect.value], 'columns'),
    area(t.rowField, [controls.rowSelect.value], 'rows'),
    area(t.valueField, areas.valueFields, 'values'),
  );
  areas.pruneActiveSettings();
  renderFieldSettingsPanel(settingsPanel, ctx, refresh);
  areaWrap.append(areasLabel, areaGrid, settingsPanel);
  container.append(title, available, areaWrap);
}
