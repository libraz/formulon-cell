import { canExecuteBuiltInRequest } from '../../commands/built-in-command-policy.js';
import { DEFAULT_PIVOT_TOGGLES, PIVOT_STYLE_TOGGLES } from '../../commands/pivot-style-info.js';
import {
  activeTableOverlay,
  renameActiveTableAction,
  TABLE_OPTION_COMMANDS,
} from '../../commands/table-design.js';
import { findPivotTableAtCell } from '../../engine/passthrough-sync.js';
import { focusSheet } from '../../interact/dialog-focus.js';
import { isInnermostRowItem, pivotRowItemAt } from '../../interact/pivot-detail.js';
import { pivotGroupTargetAt } from '../../interact/pivot-group.js';
import type { SpreadsheetInstance } from '../../mount/types.js';
import { showMessage } from '../dialogs/prompt.js';
import { projectDisabledState } from '../menu-a11y.js';
import type { RibbonCommand, ToolbarText } from '../ribbon-model.js';
import { projectPressedState } from './button.js';

export interface TableNameInputCtx {
  getInst: () => SpreadsheetInstance | null;
  text: ToolbarText;
}

/** The Table Name box of the Table Design tab: shows the active table's name and renames on commit. */
export const createTableNameInput = (
  command: RibbonCommand,
  ctx: TableNameInputCtx,
): HTMLDivElement => {
  const wrap = document.createElement('div');
  wrap.className = 'fc-tb__rb-input';
  wrap.dataset.ribbonCommand = command.id;
  wrap.dataset.ribbonInput = command.id;

  const label = document.createElement('span');
  label.className = 'fc-tb__rb-input__label';
  label.textContent = command.label;
  const input = document.createElement('input');
  input.type = 'text';
  input.className = 'fc-tb__rb-input__field';
  input.title = command.title;
  input.setAttribute('aria-label', command.title);
  input.spellcheck = false;
  input.autocomplete = 'off';
  wrap.append(label, input);

  const currentName = (): string => {
    const inst = ctx.getInst();
    return (inst && activeTableOverlay(inst.store.getState())?.name) || '';
  };
  const revert = (): void => {
    input.value = currentName();
  };

  const commit = (): void => {
    const inst = ctx.getInst();
    if (!inst || input.value === currentName()) return;
    if (!canExecuteBuiltInRequest(inst.store, command.id, 'ribbon').allowed) {
      revert();
      return;
    }
    const result = renameActiveTableAction(
      { store: inst.store, history: inst.history, origin: 'ribbon', commandId: command.id },
      input.value,
    );
    if (result === 'invalid' || result === 'duplicate') {
      void showMessage({
        title: command.title,
        message: result === 'invalid' ? ctx.text.tableNameInvalid : ctx.text.tableNameDuplicate,
        okLabel: inst.i18n.strings.pageSetup.ok,
      }).then(() => input.focus());
    }
    revert();
  };

  input.addEventListener('change', commit);
  input.addEventListener('keydown', (event) => {
    if (event.key === 'Enter') {
      event.preventDefault();
      commit();
      const sheetHost = ctx.getInst()?.host;
      if (sheetHost) focusSheet(sheetHost);
    } else if (event.key === 'Escape') {
      event.preventDefault();
      revert();
      const sheetHost = ctx.getInst()?.host;
      if (sheetHost) focusSheet(sheetHost);
    }
  });
  revert();
  return wrap;
};

/** Sync the Table Design controls under `host` with the table under the active cell. */
export const projectTableDesign = (host: HTMLElement, instance: SpreadsheetInstance): void => {
  const table = activeTableOverlay(instance.store.getState());
  const input = host.querySelector<HTMLInputElement>('[data-ribbon-input="tableName"] input');
  if (input && input !== input.ownerDocument.activeElement) input.value = table?.name ?? '';
  const reason = instance.i18n.strings.ribbon.tableSessionOnly;
  const sessionOnly = table?.source === 'session';
  // The box is disabled exactly when its commit would be refused (protection, policy).
  const renamable =
    table?.source === 'engine' &&
    canExecuteBuiltInRequest(instance.store, 'tableName', 'ribbon').allowed;
  if (input) projectDisabledState(input, !renamable, sessionOnly ? reason : null);
  // Both read the engine's table, which a session table is not.
  for (const command of ['tableSummarizePivot', 'tableInsertSlicer']) {
    const button = host.querySelector<HTMLButtonElement>(`[data-ribbon-command="${command}"]`);
    if (!button) continue;
    if (sessionOnly) {
      button.dataset.tableSessionTitle ??= button.title;
      projectDisabledState(button, true, reason, { titlePrefix: button.dataset.tableSessionTitle });
      button.dataset.tableSessionOnly = 'true';
    } else if (button.dataset.tableSessionOnly) {
      projectDisabledState(button, false, null, { titlePrefix: button.dataset.tableSessionTitle });
      delete button.dataset.tableSessionOnly;
      delete button.dataset.tableSessionTitle;
    }
  }
  for (const [command, option] of Object.entries(TABLE_OPTION_COMMANDS)) {
    const button = host.querySelector<HTMLButtonElement>(`[data-ribbon-command="${command}"]`);
    if (!button) continue;
    const pressed =
      option === 'filterButton'
        ? !!table && table.filterButton !== false
        : table?.[option] === true;
    projectPressedState(button, pressed);
    // Leave other disabled reasons (policy, protection) alone when the table is engine-backed.
    if (sessionOnly) {
      projectDisabledState(button, true, reason);
      button.dataset.tableSessionOnly = 'true';
    } else if (button.dataset.tableSessionOnly) {
      projectDisabledState(button, false, null);
      delete button.dataset.tableSessionOnly;
    }
  }
  host
    .querySelector('[data-ribbon-gallery="tableQuickStyles"]')
    ?.dispatchEvent(new Event('fc-ribbon-control-refresh'));
};

/** Expand Field and Collapse Field run on a row label. Expand Field on the innermost field and
 *  Collapse Field on a report's only row field stay greyed, as does every other cell. Group
 *  Selection runs on an item label of a date or number field. */
const projectPivotDetailCommands = (
  host: HTMLElement,
  instance: SpreadsheetInstance,
  inReport: boolean,
): void => {
  const active = instance.store.getState().selection.active;
  const ref = inReport ? pivotRowItemAt(instance.workbook, active) : null;
  const runs: Record<string, boolean> = {
    pivotExpandField: !!ref && !isInnermostRowItem(ref),
    pivotCollapseField: !!ref && !(isInnermostRowItem(ref) && ref.level === 0),
    pivotGroupSelection:
      inReport && (pivotGroupTargetAt(instance.workbook, active)?.values ?? 'other') !== 'other',
  };
  for (const [command, enabled] of Object.entries(runs)) {
    for (const button of host.querySelectorAll<HTMLButtonElement>(
      `[data-ribbon-command="${command}"]`,
    )) {
      // Re-enable only what this projection greyed; policy and protection own the rest.
      if (enabled && button.dataset.pivotDetailOff !== 'true') continue;
      projectDisabledState(button, !enabled, null);
      if (enabled) delete button.dataset.pivotDetailOff;
      else button.dataset.pivotDetailOff = 'true';
    }
  }
};

/** Sync the PivotTable toggles under `host` (the Design options and the Analyze Show group) with
 *  the report under the active cell. */
export const projectPivotDesign = (host: HTMLElement, instance: SpreadsheetInstance): void => {
  const pivot = findPivotTableAtCell(instance.workbook, instance.store.getState().selection.active);
  const info = pivot
    ? instance.workbook.getPivotStyleInfo(pivot.sheetIndex, pivot.pivotIndex)
    : null;
  for (const [command, key] of Object.entries(PIVOT_STYLE_TOGGLES)) {
    const button = host.querySelector<HTMLButtonElement>(`[data-ribbon-command="${command}"]`);
    if (!button) continue;
    const pressed = info ? info[key] : DEFAULT_PIVOT_TOGGLES[key];
    projectPressedState(button, pressed);
  }
  const options = pivot ? instance.workbook.pivotOptions(pivot.sheetIndex, pivot.pivotIndex) : null;
  for (const [command, key] of [
    ['pivotButtons', 'showDrill'],
    ['pivotHeaders', 'showHeaders'],
  ] as const) {
    const button = host.querySelector<HTMLButtonElement>(`[data-ribbon-command="${command}"]`);
    if (button) projectPressedState(button, options?.[key] ?? true);
  }
  projectPivotDetailCommands(host, instance, pivot !== null);
  host
    .querySelector('[data-ribbon-gallery="pivotQuickStyles"]')
    ?.dispatchEvent(new Event('fc-ribbon-control-refresh'));
};

/** Read-only readout in a PivotTable tab: the report's name or its active field. */
export const createPivotReadout = (
  command: RibbonCommand,
  ctx: TableNameInputCtx,
): HTMLDivElement => {
  const wrap = document.createElement('div');
  wrap.className = 'fc-tb__rb-input';
  wrap.dataset.ribbonCommand = command.id;
  wrap.dataset.ribbonInput = command.id;
  const label = document.createElement('span');
  label.className = 'fc-tb__rb-input__label';
  label.textContent = command.label;
  const input = document.createElement('input');
  input.type = 'text';
  input.className = 'fc-tb__rb-input__field';
  input.title = command.title;
  input.setAttribute('aria-label', command.title);
  input.readOnly = true;
  const inst = ctx.getInst();
  const active = inst?.store.getState().selection.active;
  const pivot = inst && active ? findPivotTableAtCell(inst.workbook, active) : null;
  const definition =
    inst && pivot ? inst.workbook.getPivotDefinition(pivot.sheetIndex, pivot.pivotIndex) : null;
  if (command.id === 'pivotName') input.value = definition?.name || '';
  if (command.id === 'pivotActiveField' && inst && active && definition) {
    // A value cell names its data field; a label names its own field.
    const info = inst.workbook.pivotCellInfo(active);
    const field = definition.fields.find((f) => f.cacheFieldIndex === info?.fieldIndex);
    input.value =
      (info?.fieldIndex !== undefined && info.fieldIndex >= 0 && field
        ? field.customName || field.sourceName
        : info?.fieldName) || '';
  }
  wrap.append(label, input);
  return wrap;
};
