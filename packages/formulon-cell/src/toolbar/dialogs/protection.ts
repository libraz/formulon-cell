import { attachRangePickerButton } from '../../interact/range-picker-control.js';
import type { SheetProtectionPermissions } from '../../store/store.js';
import { projectDisabledState } from '../menu-a11y.js';
import {
  appendDialogActions,
  appendDialogButton,
  appendErrorRow,
  appendInputRow,
  createDialogShell,
  focusAndSelectInput,
  installDialogLifecycle,
  mountDialog,
  showInputError,
} from './shell.js';

export interface ProtectSheetDialogStrings {
  title: string;
  password: string;
  passwordPlaceholder: string;
  confirmPassword: string;
  passwordMismatch: string;
  allowLabel: string;
  allowSelectLockedCells: string;
  allowSelectUnlockedCells: string;
  allowFormatCells: string;
  allowFormatColumns: string;
  allowFormatRows: string;
  allowInsertColumns: string;
  allowInsertRows: string;
  allowInsertHyperlinks: string;
  allowDeleteColumns: string;
  allowDeleteRows: string;
  allowSort: string;
  allowAutoFilter: string;
  allowPivotTables: string;
  allowEditObjects: string;
  allowEditScenarios: string;
  ok: string;
  cancel: string;
}

export interface ProtectSheetDialogResult {
  password?: string;
  permissions: SheetProtectionPermissions;
}

export interface ProtectWorkbookDialogStrings {
  title: string;
  password: string;
  passwordPlaceholder: string;
  confirmPassword: string;
  passwordMismatch: string;
  ok: string;
  cancel: string;
}

export interface AllowEditRangeDialogStrings {
  title: string;
  range: string;
  invalid: string;
  rangePickerLabel: string;
  titleLabel: string;
  password: string;
  passwordPlaceholder: string;
  confirmPassword: string;
  passwordMismatch: string;
  ok: string;
  cancel: string;
}

export interface AllowEditRangeDialogResult {
  range: string;
  title: string;
  /** Absent when the password field was never touched; null when it was
   *  cleared (removes the password); otherwise the new password. */
  password?: string | null;
}

export interface AllowEditRangesManagerStrings {
  title: string;
  newButton: string;
  modifyButton: string;
  deleteButton: string;
  close: string;
  empty: string;
}

export interface AllowEditRangesManagerRow {
  id: string;
  title: string;
  ref: string;
}

export interface UnprotectSheetDialogStrings {
  title: string;
  password: string;
  ok: string;
  cancel: string;
}

export interface AllowEditRangeDialogOptions {
  strings: AllowEditRangeDialogStrings;
  initialRange: string;
  initialTitle?: string;
  /** The range is guarded: the fields open masked, and clearing them removes the password. */
  hasPassword?: boolean;
  pickRange: () => string;
  validateRange: (value: string) => boolean;
  subscribeToRangeChanges: (listener: () => void) => () => void;
}

const permissionOrder: Array<keyof SheetProtectionPermissions> = [
  'selectLockedCells',
  'selectUnlockedCells',
  'formatCells',
  'formatColumns',
  'formatRows',
  'insertColumns',
  'insertRows',
  'insertHyperlinks',
  'deleteColumns',
  'deleteRows',
  'sort',
  'autoFilter',
  'pivotTables',
  'objects',
  'scenarios',
];

const permissionLabel = (
  strings: ProtectSheetDialogStrings,
  key: keyof SheetProtectionPermissions,
): string => {
  switch (key) {
    case 'selectLockedCells':
      return strings.allowSelectLockedCells;
    case 'selectUnlockedCells':
      return strings.allowSelectUnlockedCells;
    case 'formatCells':
      return strings.allowFormatCells;
    case 'formatColumns':
      return strings.allowFormatColumns;
    case 'formatRows':
      return strings.allowFormatRows;
    case 'insertColumns':
      return strings.allowInsertColumns;
    case 'insertRows':
      return strings.allowInsertRows;
    case 'insertHyperlinks':
      return strings.allowInsertHyperlinks;
    case 'deleteColumns':
      return strings.allowDeleteColumns;
    case 'deleteRows':
      return strings.allowDeleteRows;
    case 'sort':
      return strings.allowSort;
    case 'autoFilter':
      return strings.allowAutoFilter;
    case 'pivotTables':
      return strings.allowPivotTables;
    case 'objects':
      return strings.allowEditObjects;
    case 'scenarios':
      return strings.allowEditScenarios;
  }
};

export const defaultSheetProtectionPermissions = (): SheetProtectionPermissions => ({
  selectLockedCells: true,
  selectUnlockedCells: true,
  formatCells: false,
  formatColumns: false,
  formatRows: false,
  insertColumns: false,
  insertRows: false,
  insertHyperlinks: false,
  deleteColumns: false,
  deleteRows: false,
  sort: false,
  autoFilter: false,
  pivotTables: false,
  objects: false,
  scenarios: false,
});

export const showProtectSheetDialog = (opts: {
  strings: ProtectSheetDialogStrings;
  initial?: SheetProtectionPermissions;
}): Promise<ProtectSheetDialogResult | null> =>
  new Promise<ProtectSheetDialogResult | null>((resolve) => {
    const strings = opts.strings;
    const initial = { ...defaultSheetProtectionPermissions(), ...(opts.initial ?? {}) };
    const shell = createDialogShell({ title: strings.title, bodyVariant: 'app' });
    const password = appendInputRow(shell.body, strings.password, {
      placeholder: strings.passwordPlaceholder,
    });
    password.type = 'password';
    const confirm = appendInputRow(shell.body, strings.confirmPassword);
    confirm.type = 'password';

    const allowLabel = document.createElement('div');
    allowLabel.className = 'fc-tb__dlg__label';
    allowLabel.textContent = strings.allowLabel;
    shell.body.appendChild(allowLabel);

    const choices = document.createElement('div');
    choices.className = 'fc-fmtdlg__choice-grid fc-tb__dlg__choices';
    const inputs = new Map<keyof SheetProtectionPermissions, HTMLInputElement>();
    for (const key of permissionOrder) {
      const row = document.createElement('label');
      row.className = 'fc-fmtdlg__radio';
      const input = document.createElement('input');
      input.type = 'checkbox';
      input.checked = initial[key] === true;
      const text = document.createElement('span');
      text.textContent = permissionLabel(strings, key);
      row.append(input, text);
      choices.appendChild(row);
      inputs.set(key, input);
    }
    shell.body.appendChild(choices);
    const errorRow = appendErrorRow(shell.body);

    const { cancelBtn, okBtn } = appendDialogActions(shell.footer, {
      cancelLabel: strings.cancel,
      okLabel: strings.ok,
    });

    const permissions = (): SheetProtectionPermissions => {
      const out: SheetProtectionPermissions = {};
      for (const key of permissionOrder) out[key] = inputs.get(key)?.checked === true;
      return out;
    };

    const lifecycle = installDialogLifecycle<ProtectSheetDialogResult | null>({
      shell,
      resolve,
      onCancel: () => null,
      onSubmit: () => onOk(),
    });
    const onOk = (): void => {
      if (password.value && password.value !== confirm.value) {
        showInputError(errorRow, confirm, strings.passwordMismatch);
        return;
      }
      lifecycle.finish({
        password: password.value || undefined,
        permissions: permissions(),
      });
    };
    okBtn.addEventListener('click', onOk);
    cancelBtn.addEventListener('click', () => lifecycle.finish(null));

    mountDialog(shell, () => focusAndSelectInput(password));
  });

export const showUnprotectSheetDialog = (
  strings: UnprotectSheetDialogStrings,
): Promise<string | null> =>
  new Promise<string | null>((resolve) => {
    const shell = createDialogShell({ title: strings.title, bodyVariant: 'app' });
    const password = appendInputRow(shell.body, strings.password);
    password.type = 'password';
    const { cancelBtn, okBtn } = appendDialogActions(shell.footer, {
      cancelLabel: strings.cancel,
      okLabel: strings.ok,
    });

    const lifecycle = installDialogLifecycle<string | null>({
      shell,
      resolve,
      onCancel: () => null,
      onSubmit: () => lifecycle.finish(password.value),
    });
    okBtn.addEventListener('click', () => lifecycle.finish(password.value));
    cancelBtn.addEventListener('click', () => lifecycle.finish(null));

    mountDialog(shell, () => focusAndSelectInput(password));
  });

const PASSWORD_MASK = '\u2022\u2022\u2022\u2022\u2022\u2022\u2022\u2022';

export const showProtectWorkbookDialog = (
  strings: ProtectWorkbookDialogStrings,
): Promise<{ password?: string } | null> =>
  new Promise<{ password?: string } | null>((resolve) => {
    const shell = createDialogShell({ title: strings.title, bodyVariant: 'app' });
    const password = appendInputRow(shell.body, strings.password, {
      placeholder: strings.passwordPlaceholder,
    });
    password.type = 'password';
    const confirm = appendInputRow(shell.body, strings.confirmPassword);
    confirm.type = 'password';
    const errorRow = appendErrorRow(shell.body);
    const { cancelBtn, okBtn } = appendDialogActions(shell.footer, {
      cancelLabel: strings.cancel,
      okLabel: strings.ok,
    });

    const lifecycle = installDialogLifecycle<{ password?: string } | null>({
      shell,
      resolve,
      onCancel: () => null,
      onSubmit: () => onOk(),
    });
    const onOk = (): void => {
      if (password.value && password.value !== confirm.value) {
        showInputError(errorRow, confirm, strings.passwordMismatch);
        return;
      }
      lifecycle.finish({ password: password.value || undefined });
    };
    okBtn.addEventListener('click', onOk);
    cancelBtn.addEventListener('click', () => lifecycle.finish(null));

    mountDialog(shell, () => focusAndSelectInput(password));
  });

export const showAllowEditRangeDialog = (
  opts: AllowEditRangeDialogOptions,
): Promise<AllowEditRangeDialogResult | null> =>
  new Promise<AllowEditRangeDialogResult | null>((resolve) => {
    const { strings } = opts;
    const shell = createDialogShell({ title: strings.title, bodyVariant: 'app' });
    const rangeInput = appendInputRow(shell.body, strings.range, {
      initial: opts.initialRange,
      placeholder: 'A1:B10',
    });
    rangeInput.autocomplete = 'off';
    rangeInput.spellcheck = false;
    attachRangePickerButton(rangeInput, {
      label: strings.rangePickerLabel,
      getValue: opts.pickRange,
      subscribeToRangeChanges: opts.subscribeToRangeChanges,
      kind: 'allow-edit-ranges-range',
    });
    const titleInput = appendInputRow(shell.body, strings.titleLabel, {
      initial: opts.initialTitle ?? '',
    });
    titleInput.autocomplete = 'off';
    const password = appendInputRow(shell.body, strings.password, {
      placeholder: strings.passwordPlaceholder,
    });
    password.type = 'password';
    const confirm = appendInputRow(shell.body, strings.confirmPassword);
    confirm.type = 'password';
    let passwordTouched = false;
    if (opts.hasPassword) {
      // A mask stands for the current password; editing the field replaces it.
      password.value = confirm.value = PASSWORD_MASK;
      for (const input of [password, confirm]) {
        input.addEventListener('focus', () => focusAndSelectInput(input));
      }
    }
    password.addEventListener('input', () => {
      passwordTouched = true;
    });
    const errorRow = appendErrorRow(shell.body);
    const { cancelBtn, okBtn } = appendDialogActions(shell.footer, {
      cancelLabel: strings.cancel,
      okLabel: strings.ok,
    });

    const lifecycle = installDialogLifecycle<AllowEditRangeDialogResult | null>({
      shell,
      resolve,
      onCancel: () => null,
      onSubmit: () => onOk(),
    });
    const onOk = (): void => {
      const value = rangeInput.value.trim();
      if (!opts.validateRange(value)) {
        showInputError(errorRow, rangeInput, strings.invalid);
        return;
      }
      if (passwordTouched && password.value && password.value !== confirm.value) {
        showInputError(errorRow, confirm, strings.passwordMismatch);
        return;
      }
      lifecycle.finish({
        range: value,
        title: titleInput.value.trim(),
        ...(passwordTouched ? { password: password.value || null } : {}),
      });
    };
    okBtn.addEventListener('click', onOk);
    cancelBtn.addEventListener('click', () => lifecycle.finish(null));

    mountDialog(shell, () => focusAndSelectInput(rangeInput));
  });

/**
 * The Allow Users to Edit Ranges list: New, Modify and Delete act on the
 * selected row through the callbacks, and `rows` is re-read after each. Resolves
 * when the dialog is closed.
 */
export const showAllowEditRangesManagerDialog = (opts: {
  strings: AllowEditRangesManagerStrings;
  rows: () => readonly AllowEditRangesManagerRow[];
  onNew: () => Promise<void>;
  onModify: (id: string) => Promise<void>;
  onDelete: (id: string) => Promise<void>;
}): Promise<void> =>
  new Promise<void>((resolve) => {
    const { strings } = opts;
    const shell = createDialogShell({ title: strings.title, bodyVariant: 'app' });
    const list = document.createElement('div');
    list.className = 'fc-tb__dlg__list';
    list.setAttribute('role', 'radiogroup');
    list.setAttribute('aria-label', strings.title);
    shell.body.appendChild(list);

    const newBtn = appendDialogButton(shell.footer, { label: strings.newButton });
    const modifyBtn = appendDialogButton(shell.footer, { label: strings.modifyButton });
    const deleteBtn = appendDialogButton(shell.footer, { label: strings.deleteButton });
    const closeBtn = appendDialogButton(shell.footer, {
      label: strings.close,
      variant: 'primary',
    });

    let selected = '';
    const render = (): void => {
      const rows = opts.rows();
      if (!rows.some((row) => row.id === selected)) selected = rows[0]?.id ?? '';
      list.replaceChildren();
      if (rows.length === 0) {
        const empty = document.createElement('div');
        empty.className = 'fc-tb__dlg__note';
        empty.textContent = strings.empty;
        list.appendChild(empty);
      }
      for (const row of rows) {
        const label = document.createElement('label');
        label.className = 'fc-fmtdlg__row fc-fmtdlg__row--block fc-tb__dlg__check';
        const radio = document.createElement('input');
        radio.type = 'radio';
        radio.name = 'fc-allow-edit-ranges';
        radio.value = row.id;
        radio.checked = row.id === selected;
        radio.addEventListener('change', () => {
          if (radio.checked) selected = row.id;
        });
        const body = document.createElement('span');
        const name = document.createElement('strong');
        name.textContent = row.title;
        const ref = document.createElement('span');
        ref.className = 'fc-tb__dlg__note';
        ref.textContent = row.ref;
        body.append(name, ref);
        label.append(radio, body);
        list.appendChild(label);
      }
      projectDisabledState(modifyBtn, selected === '', null);
      projectDisabledState(deleteBtn, selected === '', null);
    };
    const run = async (task: () => Promise<void>): Promise<void> => {
      await task();
      render();
    };
    newBtn.addEventListener('click', () => void run(opts.onNew));
    modifyBtn.addEventListener('click', () => void run(() => opts.onModify(selected)));
    deleteBtn.addEventListener('click', () => void run(() => opts.onDelete(selected)));

    const lifecycle = installDialogLifecycle<void>({
      shell,
      resolve,
      onCancel: () => undefined,
    });
    closeBtn.addEventListener('click', () => lifecycle.finish());

    render();
    mountDialog(shell, newBtn);
  });
