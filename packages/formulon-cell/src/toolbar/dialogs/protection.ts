import type { SheetProtectionPermissions } from '../../store/store.js';
import { attachRangePickerButton } from '../../interact/range-picker-control.js';
import {
  appendDialogActions,
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

export interface AllowEditRangeDialogStrings {
  title: string;
  range: string;
  invalid: string;
  rangePickerLabel: string;
  ok: string;
  cancel: string;
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
    allowLabel.className = 'app__dlg__label';
    allowLabel.textContent = strings.allowLabel;
    shell.body.appendChild(allowLabel);

    const choices = document.createElement('div');
    choices.className = 'fc-fmtdlg__choice-grid app__dlg__choices';
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

export const showAllowEditRangeDialog = (
  opts: AllowEditRangeDialogOptions,
): Promise<string | null> =>
  new Promise<string | null>((resolve) => {
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
    const errorRow = appendErrorRow(shell.body);
    const { cancelBtn, okBtn } = appendDialogActions(shell.footer, {
      cancelLabel: strings.cancel,
      okLabel: strings.ok,
    });

    const lifecycle = installDialogLifecycle<string | null>({
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
      lifecycle.finish(value);
    };
    okBtn.addEventListener('click', onOk);
    cancelBtn.addEventListener('click', () => lifecycle.finish(null));

    mountDialog(shell, () => focusAndSelectInput(rangeInput));
  });
