import { attachRangePickerButton } from '../../interact/range-picker-control.js';
import type { SheetProtectionPermissions } from '../../store/store.js';
import { projectDisabledState } from '../menu-a11y.js';
import { appendNamedRadioRow } from './form-controls.js';
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
  /** Lines above the password fields; omitted when absent or empty. */
  heading?: string;
  note?: string;
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

/** The commit buttons of the Windows layout: Cancel discards and Apply commits
 *  without closing; Permissions... (Windows ACLs) stays disabled. */
export interface AllowEditRangesManagerCommitStrings {
  permissions: string;
  protectSheet: string;
  cancel: string;
  apply: string;
}

/** How the manager closed: `commit` keeps the changes (OK, or Esc in the OK-only
 *  layout), `discard` drops them (Cancel, Esc), `protectSheet` keeps them and
 *  asks for Protect Sheet next. */
export type AllowEditRangesManagerResult = 'commit' | 'discard' | 'protectSheet';

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

const appendPasswordPair = (
  body: HTMLElement,
  strings: {
    password: string;
    passwordPlaceholder: string;
    confirmPassword: string;
    passwordMismatch: string;
  },
): {
  password: HTMLInputElement;
  confirm: HTMLInputElement;
  /** Shows the mismatch error and returns true when a password was typed and the confirmation differs. */
  mismatch: (errorRow: HTMLElement) => boolean;
} => {
  const password = appendInputRow(body, strings.password, {
    placeholder: strings.passwordPlaceholder,
  });
  password.type = 'password';
  const confirm = appendInputRow(body, strings.confirmPassword);
  confirm.type = 'password';
  return {
    password,
    confirm,
    mismatch: (errorRow) => {
      if (!password.value || password.value === confirm.value) return false;
      showInputError(errorRow, confirm, strings.passwordMismatch);
      return true;
    },
  };
};

const PERMISSIONS: ReadonlyArray<{
  key: keyof SheetProtectionPermissions;
  label: keyof ProtectSheetDialogStrings;
  allowed: boolean;
}> = [
  { key: 'selectLockedCells', label: 'allowSelectLockedCells', allowed: true },
  { key: 'selectUnlockedCells', label: 'allowSelectUnlockedCells', allowed: true },
  { key: 'formatCells', label: 'allowFormatCells', allowed: false },
  { key: 'formatColumns', label: 'allowFormatColumns', allowed: false },
  { key: 'formatRows', label: 'allowFormatRows', allowed: false },
  { key: 'insertColumns', label: 'allowInsertColumns', allowed: false },
  { key: 'insertRows', label: 'allowInsertRows', allowed: false },
  { key: 'insertHyperlinks', label: 'allowInsertHyperlinks', allowed: false },
  { key: 'deleteColumns', label: 'allowDeleteColumns', allowed: false },
  { key: 'deleteRows', label: 'allowDeleteRows', allowed: false },
  { key: 'sort', label: 'allowSort', allowed: false },
  { key: 'autoFilter', label: 'allowAutoFilter', allowed: false },
  { key: 'pivotTables', label: 'allowPivotTables', allowed: false },
  { key: 'objects', label: 'allowEditObjects', allowed: false },
  { key: 'scenarios', label: 'allowEditScenarios', allowed: false },
];

const defaultSheetProtectionPermissions = (): SheetProtectionPermissions =>
  Object.fromEntries(PERMISSIONS.map(({ key, allowed }) => [key, allowed]));

export const showProtectSheetDialog = (opts: {
  strings: ProtectSheetDialogStrings;
  initial?: SheetProtectionPermissions;
}): Promise<ProtectSheetDialogResult | null> =>
  new Promise<ProtectSheetDialogResult | null>((resolve) => {
    const strings = opts.strings;
    const initial = { ...defaultSheetProtectionPermissions(), ...(opts.initial ?? {}) };
    const shell = createDialogShell({ title: strings.title, bodyVariant: 'app' });
    for (const [text, className] of [
      [strings.heading, 'fc-tb__dlg__label'],
      [strings.note, 'fc-tb__dlg__message'],
    ] as const) {
      if (!text) continue;
      const line = document.createElement('div');
      line.className = className;
      line.textContent = text;
      shell.body.appendChild(line);
    }
    const { password, mismatch } = appendPasswordPair(shell.body, strings);

    const allowLabel = document.createElement('div');
    allowLabel.className = 'fc-tb__dlg__label';
    allowLabel.textContent = strings.allowLabel;
    shell.body.appendChild(allowLabel);

    const choices = document.createElement('div');
    choices.className = 'fc-fmtdlg__choice-grid fc-tb__dlg__choices fc-protectdlg__choices';
    choices.setAttribute('role', 'group');
    choices.setAttribute('aria-label', strings.allowLabel);
    const inputs = new Map<keyof SheetProtectionPermissions, HTMLInputElement>();
    for (const { key, label } of PERMISSIONS) {
      const row = document.createElement('label');
      row.className = 'fc-fmtdlg__radio';
      const input = document.createElement('input');
      input.type = 'checkbox';
      input.checked = initial[key] === true;
      const text = document.createElement('span');
      text.textContent = strings[label] as string;
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
      for (const { key } of PERMISSIONS) out[key] = inputs.get(key)?.checked === true;
      return out;
    };

    const lifecycle = installDialogLifecycle<ProtectSheetDialogResult | null>({
      shell,
      resolve,
      onCancel: () => null,
      onSubmit: () => onOk(),
      buttons: { okBtn, cancelBtn },
    });
    const onOk = (): void => {
      if (mismatch(errorRow)) return;
      lifecycle.finish({
        password: password.value || undefined,
        permissions: permissions(),
      });
    };

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
      buttons: { okBtn, cancelBtn },
    });

    mountDialog(shell, () => focusAndSelectInput(password));
  });

const PASSWORD_MASK = '\u2022\u2022\u2022\u2022\u2022\u2022\u2022\u2022';

export const showProtectWorkbookDialog = (
  strings: ProtectWorkbookDialogStrings,
): Promise<{ password?: string } | null> =>
  new Promise<{ password?: string } | null>((resolve) => {
    const shell = createDialogShell({ title: strings.title, bodyVariant: 'app' });
    const { password, mismatch } = appendPasswordPair(shell.body, strings);
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
      buttons: { okBtn, cancelBtn },
    });
    const onOk = (): void => {
      if (mismatch(errorRow)) return;
      lifecycle.finish({ password: password.value || undefined });
    };

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
    const { password, confirm, mismatch } = appendPasswordPair(shell.body, strings);
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
      buttons: { okBtn, cancelBtn },
    });
    const onOk = (): void => {
      const value = rangeInput.value.trim();
      if (!opts.validateRange(value)) {
        showInputError(errorRow, rangeInput, strings.invalid);
        return;
      }
      if (passwordTouched && mismatch(errorRow)) return;
      lifecycle.finish({
        range: value,
        title: titleInput.value.trim(),
        ...(passwordTouched ? { password: password.value || null } : {}),
      });
    };

    mountDialog(shell, () => focusAndSelectInput(rangeInput));
  });

/**
 * The Allow Users to Edit Ranges list: New, Modify and Delete act on the
 * selected row through the callbacks, and `rows` is re-read after each. With
 * `commit` it takes the Windows layout (Permissions..., Protect Sheet..., OK,
 * Cancel, Apply); without it, OK alone. Resolves with how it was closed.
 */
export const showAllowEditRangesManagerDialog = (opts: {
  strings: AllowEditRangesManagerStrings;
  rows: () => readonly AllowEditRangesManagerRow[];
  onNew: () => Promise<void>;
  onModify: (id: string) => Promise<void>;
  onDelete: (id: string) => Promise<void>;
  commit?: { strings: AllowEditRangesManagerCommitStrings; onApply: () => void };
}): Promise<AllowEditRangesManagerResult> =>
  new Promise<AllowEditRangesManagerResult>((resolve) => {
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
    const commit = opts.commit;
    const permissionsBtn = commit
      ? appendDialogButton(shell.footer, { label: commit.strings.permissions })
      : null;
    if (permissionsBtn) projectDisabledState(permissionsBtn, true, null);
    const protectBtn = commit
      ? appendDialogButton(shell.footer, { label: commit.strings.protectSheet })
      : null;
    const closeBtn = appendDialogButton(shell.footer, {
      label: strings.close,
      variant: 'primary',
    });
    const cancelBtn = commit
      ? appendDialogButton(shell.footer, { label: commit.strings.cancel })
      : null;
    const applyBtn = commit
      ? appendDialogButton(shell.footer, { label: commit.strings.apply })
      : null;

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
        appendNamedRadioRow(list, {
          name: 'fc-allow-edit-ranges',
          value: row.id,
          title: row.title,
          note: row.ref,
          checked: row.id === selected,
          onSelect: (id) => {
            selected = id;
          },
        });
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

    const lifecycle = installDialogLifecycle<AllowEditRangesManagerResult>({
      shell,
      resolve,
      onCancel: () => (commit ? 'discard' : 'commit'),
    });
    closeBtn.addEventListener('click', () => lifecycle.finish('commit'));
    cancelBtn?.addEventListener('click', () => lifecycle.finish('discard'));
    protectBtn?.addEventListener('click', () => lifecycle.finish('protectSheet'));
    applyBtn?.addEventListener('click', () => commit?.onApply());

    render();
    mountDialog(shell, newBtn);
  });
