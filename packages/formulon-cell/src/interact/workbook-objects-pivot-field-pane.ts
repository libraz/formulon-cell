import type { History } from '../commands/history.js';
import type { PivotAuthorization } from '../commands/pivot-actions.js';
import { applyPivotEdit } from '../commands/pivot-edit.js';
import { PivotAxis, PivotReportLayout } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { Strings } from '../i18n/strings.js';
import { type PivotAreaKind, pivotFilterInputFor } from './pivot-field-settings.js';
import {
  type PivotNumberFormatPicker,
  showPivotFieldSettingsDialog,
} from './pivot-field-settings-dialog.js';
import {
  applyPivotFieldSettings,
  commitPivotChange,
  pivotFieldSettingsTarget,
} from './pivot-field-settings-dialog-apply.js';
import {
  PIVOT_AREA_ORDER,
  type PivotFieldAreasModel,
  renderPivotFieldAreas,
} from './pivot-table-field-list.js';
import {
  createPivotFieldEditor,
  editedPivotFilters,
  type PivotEditorTarget,
  type PivotFieldEditor,
  pivotEditorFilterConditions,
  readPivotEditorModel,
} from './workbook-objects-pivot-editor.js';

/** Field indices per area, in display order. */
export type PivotPaneLayout = Record<PivotAreaKind, number[]>;

/** Pane state the host keeps across re-renders. */
export interface PivotFieldPaneState {
  /** The chip whose settings are open. */
  active: { kind: PivotAreaKind; fieldName: string } | null;
  search: string;
  /** The layout as last written, for an engine that cannot read the definition back. */
  layout: PivotPaneLayout | null;
  /** Field list height set by dragging the splitter; null for the stylesheet default. */
  listHeight: number | null;
}

export const createPivotFieldPaneState = (): PivotFieldPaneState => ({
  active: null,
  search: '',
  layout: null,
  listHeight: null,
});

const LIST_MIN_HEIGHT = 44;
const LIST_MAX_HEIGHT = 600;
const LIST_KEY_STEP = 22;

export interface PivotFieldPaneContext extends PivotAuthorization {
  host: HTMLElement;
  wb: WorkbookHandle;
  strings: Strings;
  history?: History | null;
  state: PivotFieldPaneState;
  errorText: string;
  setError(message: string): void;
  /** Re-renders the pane. */
  rerender(): void;
  /** Opens Format Cells on the number tab for the field dialog's Number Format button. */
  pickNumberFormat?: PivotNumberFormatPicker;
  /** The report was rewritten, by an edit, an undo or a redo; repaint the grid and the pane. */
  onChanged(): void;
}

const AREA_AXIS: Record<PivotAreaKind, PivotAxis> = {
  filters: PivotAxis.Page,
  columns: PivotAxis.Col,
  rows: PivotAxis.Row,
  values: PivotAxis.Value,
};

const AXIS_AREA = new Map<PivotAxis, PivotAreaKind>(
  PIVOT_AREA_ORDER.map((kind) => [AREA_AXIS[kind], kind]),
);

const emptyLayout = (): PivotPaneLayout => ({ filters: [], columns: [], rows: [], values: [] });

/** The docked PivotTable Fields pane: field checkboxes and the four areas.
 *  Every change is written to the report at once, as one undo step. */
export function renderPivotFieldListPane(
  ctx: PivotFieldPaneContext,
  pivot: PivotEditorTarget,
): HTMLElement {
  const { wb, strings, state } = ctx;
  const t = strings.workbookObjects;
  const td = strings.pivotTableDialog;
  const root = document.createElement('div');
  root.className = 'fc-objects__pivot-pane';
  root.setAttribute('aria-label', t.pivotFieldList);
  const editorModel = readPivotEditorModel(wb, pivot);
  const { definition } = editorModel;
  const filterConditions = pivotEditorFilterConditions(editorModel, pivotFilterInputFor(wb));
  const error = document.createElement('div');
  error.className = 'fc-objects__error';
  error.setAttribute('role', 'alert');
  error.hidden = !ctx.errorText;
  error.textContent = ctx.errorText;
  let write: (layout: PivotPaneLayout) => void = () => {};
  const editors = editorModel.fields
    .map((model) =>
      createPivotFieldEditor(
        { ...ctx, dataFields: definition?.dataFields.map((spec) => spec.name ?? '') ?? [] },
        model,
        filterConditions.get(model.index),
        () => write(layout),
      ),
    )
    .sort((a, b) => a.model.index - b.model.index);
  const editorOf = new Map<string, PivotFieldEditor>(
    editors.map((editor) => [editor.model.label, editor]),
  );
  const labelOf = (index: number): string =>
    editors.find((editor) => editor.model.index === index)?.model.label ?? '';

  const layoutFromModel = (): PivotPaneLayout => {
    if (!definition) {
      if (state.layout) return state.layout;
      const guessed = emptyLayout();
      for (const field of editorModel.fields) {
        const kind = field.axis === null ? undefined : AXIS_AREA.get(field.axis);
        if (kind) guessed[kind].push(field.index);
      }
      return guessed;
    }
    const valueOrder = [
      ...new Set([
        ...definition.dataFields.map((spec) => spec.fieldIndex),
        ...definition.fields.filter((f) => f.axis === PivotAxis.Value).map((f) => f.index),
      ]),
    ];
    return {
      filters: definition.fields.filter((f) => f.axis === PivotAxis.Page).map((f) => f.index),
      columns: [...definition.colFieldOrder],
      rows: [...definition.rowFieldOrder],
      values: valueOrder.filter((index) => definition.fields[index]?.axis === PivotAxis.Value),
    };
  };
  const layout = layoutFromModel();
  const axisIn = (next: PivotPaneLayout, index: number): PivotAxis | null => {
    const kind = PIVOT_AREA_ORDER.find((area) => next[area].includes(index));
    return kind ? AREA_AXIS[kind] : null;
  };
  for (const editor of editors) editor.showFor(axisIn(layout, editor.model.index));

  write = (next: PivotPaneLayout): void => {
    ctx.setError('');
    const placedIndexes = [...next.rows, ...next.columns];
    const ordered = [
      ...placedIndexes.map((index) => editors.find((editor) => editor.model.index === index)),
      ...editors.filter((editor) => !placedIndexes.includes(editor.model.index)),
    ].filter((editor): editor is PivotFieldEditor => editor !== undefined);
    const placed = ordered.map((editor) => {
      const axis = axisIn(next, editor.model.index);
      editor.showFor(axis);
      return { editor, axis };
    });
    // Shown in place: a re-render would drop the Custom choice being made.
    if (placed.some(({ editor }) => editor.subtotalEmpty())) {
      ctx.setError(t.pivotSubtotalCustomEmpty);
      error.textContent = t.pivotSubtotalCustomEmpty;
      error.hidden = false;
      return;
    }
    const { sheetIndex, pivotIndex } = pivot;
    commitPivotChange(ctx, sheetIndex, pivotIndex, () => {
      const before = wb.getPivotDefinition(sheetIndex, pivotIndex);
      let ok = applyPivotEdit(wb, pivot, {
        fieldListOnly: true,
        name: '',
        anchor: { row: pivot.top, col: pivot.left },
        rowGrandTotals: true,
        colGrandTotals: true,
        layout: PivotReportLayout.Compact,
        ...(before ? { dataFields: before.dataFields } : {}),
        fields: placed.map(({ editor, axis }) => editor.toEditField(axis)),
        filters: editedPivotFilters(editorModel.pivotFilters, placed),
      });
      // Value fields keep their written order; a reorder within Values rewrites the list.
      const written = before ? wb.getPivotDefinition(sheetIndex, pivotIndex) : null;
      if (written) {
        const rank = (fieldIndex: number): number => {
          const position = next.values.indexOf(fieldIndex);
          return position < 0 ? next.values.length : position;
        };
        const sorted = written.dataFields
          .map((spec, position) => ({ spec, position }))
          .sort(
            (a, b) => rank(a.spec.fieldIndex) - rank(b.spec.fieldIndex) || a.position - b.position,
          )
          .map(({ spec }) => spec);
        if (sorted.some((spec, position) => spec !== written.dataFields[position])) {
          ok =
            wb.clearPivotDataFields(sheetIndex, pivotIndex) &&
            sorted.every((spec) => wb.addPivotDataField(sheetIndex, pivotIndex, spec) >= 0) &&
            ok;
        }
      }
      state.layout = before ? null : next;
      if (!ok) ctx.setError(t.pivotEditFailed);
      return ok;
    });
  };

  // The PivotTable Field dialog for a placed field; its result is one undo step.
  const openFieldDialog = (kind: PivotAreaKind, fieldName: string): void => {
    const index = indexOf(fieldName);
    if (index === undefined || !definition) return;
    const which =
      kind === 'values'
        ? {
            kind: 'value' as const,
            dataFieldIndex: definition.dataFields.findIndex((spec) => spec.fieldIndex === index),
          }
        : { kind: 'field' as const, fieldIndex: index };
    if (which.kind === 'value' && which.dataFieldIndex < 0) return;
    const target = pivotFieldSettingsTarget(
      wb,
      pivot.sheetIndex,
      pivot.pivotIndex,
      which,
      (sourceName) => pivot.fieldItems?.[sourceName] ?? [],
    );
    if (!target) return;
    void showPivotFieldSettingsDialog({
      host: ctx.host,
      strings,
      target,
      ...(ctx.pickNumberFormat ? { pickNumberFormat: ctx.pickNumberFormat } : {}),
    }).then((result) => {
      if (result) applyPivotFieldSettings(ctx, pivot, which, result);
    });
  };

  const without = (index: number): PivotPaneLayout => {
    const next = emptyLayout();
    for (const kind of PIVOT_AREA_ORDER) next[kind] = layout[kind].filter((i) => i !== index);
    return next;
  };
  const indexOf = (fieldName: string): number | undefined => editorOf.get(fieldName)?.model.index;
  const model: PivotFieldAreasModel = {
    fields: editors.map((editor) => editor.model.label),
    fieldsIn: (kind) => layout[kind].map(labelOf),
    // Two or more value fields add the Values field; the engine lays it across
    // the columns and does not report a values-on-rows placement.
    fixedChips: (kind) =>
      kind === 'columns' && layout.values.length >= 2 ? [t.pivotAreaValues] : [],
    chipLabel: (kind, fieldName) => {
      if (kind !== 'values' || !definition) return fieldName;
      const index = indexOf(fieldName);
      return definition.dataFields.find((spec) => spec.fieldIndex === index)?.name || fieldName;
    },
    canPlace: (fieldName) => indexOf(fieldName) !== undefined,
    toggle: (fieldName, checked) => {
      const editor = editorOf.get(fieldName);
      if (!editor) return;
      const next = without(editor.model.index);
      if (checked) {
        // A ticked field goes to the top of its area, as the desktop pane does.
        const kind = AXIS_AREA.get(editor.model.defaultAxis) ?? 'rows';
        next[kind].unshift(editor.model.index);
      }
      write(next);
    },
    place: (fieldName, kind, before) => {
      const index = indexOf(fieldName);
      if (index === undefined) return;
      const next = without(index);
      const target = before === null ? undefined : indexOf(before);
      const at = target === undefined ? -1 : next[kind].indexOf(target);
      if (at < 0) next[kind].push(index);
      else next[kind].splice(at, 0, index);
      write(next);
    },
    remove: (fieldName) => {
      const index = indexOf(fieldName);
      if (index !== undefined) write(without(index));
    },
  };

  // Filter items and conditions stay in the pane; every other setting is in the dialog.
  const active = state.active?.kind === 'filters' ? state.active : null;
  const activeEditor =
    active && model.fieldsIn(active.kind).includes(active.fieldName)
      ? editorOf.get(active.fieldName)
      : undefined;
  if (!activeEditor) state.active = null;
  const settingsPanel = document.createElement('div');
  settingsPanel.className = 'fc-pivotdlg__area-settings-panel';
  settingsPanel.hidden = !activeEditor;
  if (active && activeEditor) {
    const heading = document.createElement('span');
    heading.textContent = td.fieldSettingsFor.replace('{field}', active.fieldName);
    settingsPanel.append(heading, activeEditor.filterSettings);
    settingsPanel.addEventListener('change', () => write(layout));
  }

  renderPivotFieldAreas(root, {
    model,
    labels: {
      availableFields: t.pivotFieldName,
      filters: t.pivotAreaFilters,
      columns: t.pivotAreaColumns,
      rows: t.pivotAreaRows,
      values: t.pivotAreaValues,
      fieldSettings: td.fieldSettings,
      fieldSettingsFor: td.fieldSettingsFor,
      search: t.pivotFieldSearch,
      areaHint: t.pivotDragFieldsHint,
    },
    refresh: () => {},
    onFieldSettings: (kind, fieldName) => {
      if (kind !== 'filters') {
        openFieldDialog(kind, fieldName);
        return;
      }
      const same = state.active?.kind === kind && state.active.fieldName === fieldName;
      state.active = same ? null : { kind, fieldName };
      ctx.rerender();
    },
    onFieldInfo: openFieldDialog,
    searchText: state.search,
    onSearch: (text) => {
      state.search = text;
    },
    areaFooter: [settingsPanel, error],
  });
  const areas = root.querySelector('.fc-pivotdlg__areas');
  if (areas) root.insertBefore(createListSplitter(root, state), areas);
  if (state.listHeight !== null) {
    root.style.setProperty('--fc-pivot-pane-list-height', `${state.listHeight}px`);
  }
  return root;
}

/** The grip between the field list and the areas: dragging it, or the arrow
 *  keys, resizes the list. */
function createListSplitter(root: HTMLElement, state: PivotFieldPaneState): HTMLElement {
  const splitter = document.createElement('div');
  splitter.className = 'fc-pivotdlg__pane-splitter';
  splitter.setAttribute('role', 'separator');
  splitter.setAttribute('aria-orientation', 'horizontal');
  splitter.tabIndex = 0;
  splitter.dataset.pivotPaneSplitter = '';
  const listHeight = (): number =>
    state.listHeight ??
    root.querySelector<HTMLElement>('.fc-pivotdlg__field-list-grid')?.getBoundingClientRect()
      .height ??
    LIST_MIN_HEIGHT;
  const resize = (height: number): void => {
    state.listHeight = Math.round(Math.min(LIST_MAX_HEIGHT, Math.max(LIST_MIN_HEIGHT, height)));
    root.style.setProperty('--fc-pivot-pane-list-height', `${state.listHeight}px`);
  };
  splitter.addEventListener('pointerdown', (event) => {
    event.preventDefault();
    const startY = event.clientY;
    const startHeight = listHeight();
    splitter.setPointerCapture?.(event.pointerId);
    const move = (next: PointerEvent): void => resize(startHeight + next.clientY - startY);
    const end = (): void => {
      splitter.removeEventListener('pointermove', move);
      splitter.removeEventListener('pointerup', end);
      splitter.removeEventListener('pointercancel', end);
    };
    splitter.addEventListener('pointermove', move);
    splitter.addEventListener('pointerup', end);
    splitter.addEventListener('pointercancel', end);
  });
  splitter.addEventListener('keydown', (event) => {
    if (event.key !== 'ArrowUp' && event.key !== 'ArrowDown') return;
    event.preventDefault();
    resize(listHeight() + (event.key === 'ArrowDown' ? LIST_KEY_STEP : -LIST_KEY_STEP));
  });
  return splitter;
}
