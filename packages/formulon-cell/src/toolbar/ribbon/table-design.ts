import { canExecuteBuiltIn } from '../../commands/built-in-command-policy.js';
import {
  activeTableOverlay,
  renameActiveTable,
  TABLE_OPTION_COMMANDS,
} from '../../commands/table-design.js';
import type { SpreadsheetInstance } from '../../mount/types.js';
import { showMessage } from '../dialogs/prompt.js';
import { projectDisabledState } from '../menu-a11y.js';
import type { RibbonCommand, ToolbarText } from '../ribbon-model.js';

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
    if (!canExecuteBuiltIn(inst.store, command.id, 'ribbon').allowed) {
      revert();
      return;
    }
    const result = renameActiveTable(inst.store, inst.history, input.value);
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
      ctx.getInst()?.host.focus();
    } else if (event.key === 'Escape') {
      event.preventDefault();
      revert();
      ctx.getInst()?.host.focus();
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
  if (input) projectDisabledState(input, table?.source !== 'engine', sessionOnly ? reason : null);
  for (const [command, option] of Object.entries(TABLE_OPTION_COMMANDS)) {
    const button = host.querySelector<HTMLButtonElement>(`[data-ribbon-command="${command}"]`);
    if (!button) continue;
    const pressed = table?.[option] === true;
    button.classList.toggle('fc-tb__rb--active', pressed);
    button.setAttribute('aria-pressed', pressed ? 'true' : 'false');
    // Leave other disabled reasons (policy, protection) alone when the table is engine-backed.
    if (sessionOnly) {
      projectDisabledState(button, true, reason);
      button.dataset.tableSessionOnly = 'true';
    } else if (button.dataset.tableSessionOnly) {
      projectDisabledState(button, false, null);
      delete button.dataset.tableSessionOnly;
    }
  }
};
