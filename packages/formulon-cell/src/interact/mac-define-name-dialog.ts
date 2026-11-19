/** @file Mac Define Name sub-dialog: a names list with + / -, a name field and a range field. */
import type { namedRangeWrite } from '../commands/named-range-actions.js';
import {
  type DefinedNameEntry,
  definedNameDisplayFormula,
  deleteDefinedName,
  isSameDefinedName,
  upsertDefinedName,
} from '../commands/named-ranges.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { Strings } from '../i18n/strings.js';
import { projectDisabledState } from '../toolbar/menu-a11y.js';
import {
  appendDialogActions,
  appendDialogFrame,
  appendDialogIconButton,
  clearDialogError,
  createDialogShell,
  deferInitialFocus,
  focusAndSelectInput,
  showDialogError,
} from './dialog-shell.js';
import { errorMessageFor } from './named-range-dialog-model.js';
import { attachRangePickerButton, type RangePickerControlOptions } from './range-picker-control.js';

export interface MacDefineNameDialogContext {
  /** Element the dialog overlay mounts into. */
  host: HTMLElement;
  /** Locale strings. */
  strings: Strings;
  /** Current Name Manager strings (re-resolved on every open). */
  text(): Strings['namedRangeDialog'];
  /** Current workbook handle (re-bound by the host). */
  workbook(): WorkbookHandle;
  /** Runs a defined-name change as one authorized, undoable step. */
  writeNames: <T>(write: () => T) => ReturnType<typeof namedRangeWrite<T>>;
  /** Range-picker options, or null when the host provides no range picker. */
  rangePicker: RangePickerControlOptions | null;
  /** Absolute, sheet-qualified formula for the current selection. */
  seedFormula(): string;
  /** Called after a defined-name mutation writes through to the engine. */
  afterMutate(): void;
}

export function createMacDefineNameDialog(ctx: MacDefineNameDialogContext): {
  open(): void;
  dispose(): void;
} {
  // ── Mac Define Name: names list with + / -, a name and a range field ──
  const defineShell = createDialogShell({
    host: ctx.host,
    className: 'fc-namedlg-define',
    ariaLabel: ctx.text().title,
    onDismiss: () => defineShell.close(),
  });
  defineShell.overlay.classList.add('fc-fmtdlg');
  const defineFrame = appendDialogFrame(defineShell, {
    title: '',
    panelClasses: ['fc-fmtdlg__panel', 'fc-namedlg-define__panel'],
    bodyClass: 'fc-fmtdlg__body fc-namedlg-define__body',
    footerClass: 'fc-fmtdlg__footer fc-namedlg-define__footer',
  });
  const defineNamesLabel = document.createElement('div');
  defineNamesLabel.className = 'fc-namedlg-define__names-label';
  const defineNames = document.createElement('div');
  defineNames.className = 'fc-namedlg-define__names';
  defineNames.setAttribute('role', 'listbox');
  const defineListButtons = document.createElement('div');
  defineListButtons.className = 'fc-namedlg-define__list-buttons';
  const defineAddBtn = appendDialogIconButton(defineListButtons, {
    label: '+',
    ariaLabel: ctx.text().addButton,
    baseClass: 'fc-namedlg-define__list-btn',
  });
  defineAddBtn.dataset.defineAction = 'add';
  const defineRemoveBtn = appendDialogIconButton(defineListButtons, {
    label: '−',
    ariaLabel: ctx.text().deleteButton,
    baseClass: 'fc-namedlg-define__list-btn',
  });
  defineRemoveBtn.dataset.defineAction = 'remove';
  const defineLeft = document.createElement('div');
  defineLeft.className = 'fc-namedlg-define__left';
  defineLeft.append(defineNamesLabel, defineNames, defineListButtons);
  const defineField = (): {
    label: HTMLLabelElement;
    text: HTMLSpanElement;
    input: HTMLInputElement;
  } => {
    const label = document.createElement('label');
    label.className = 'fc-namedlg-define__field';
    const text = document.createElement('span');
    const input = document.createElement('input');
    input.type = 'text';
    input.className = 'fc-namedlg-editor__input';
    input.autocomplete = 'off';
    input.spellcheck = false;
    label.append(text, input);
    return { label, text, input };
  };
  const defineName = defineField();
  const defineRange = defineField();
  const defineRight = document.createElement('div');
  defineRight.className = 'fc-namedlg-define__right';
  defineRight.append(defineName.label, defineRange.label);
  if (ctx.rangePicker) {
    attachRangePickerButton(defineRange.input, {
      ...ctx.rangePicker,
      kind: 'named-range-define-refers-to',
    });
  }
  const defineError = document.createElement('div');
  defineError.className = 'fc-namedlg__error';
  defineError.setAttribute('role', 'alert');
  defineError.hidden = true;
  defineFrame.body.append(defineLeft, defineRight, defineError);
  const defineActions = appendDialogActions(defineFrame.footer, {
    cancelLabel: ctx.text().close,
    okLabel: ctx.text().ok,
  });
  let defineSelected: DefinedNameEntry | null = null;

  const renderDefineNames = (): void => {
    const entries = [...ctx.workbook().definedNames()].filter((entry) => !entry.hidden);
    defineNames.replaceChildren(
      ...entries.map((entry) => {
        const option = document.createElement('div');
        option.setAttribute('role', 'option');
        option.dataset.name = entry.name;
        option.textContent = ctx.workbook().localizeDefinedName(entry.name);
        const selected =
          defineSelected !== null && isSameDefinedName(defineSelected.name, entry.name);
        option.setAttribute('aria-selected', String(selected));
        option.addEventListener('click', () => {
          defineSelected = entry;
          defineName.input.value = ctx.workbook().localizeDefinedName(entry.name);
          defineRange.input.value = definedNameDisplayFormula(entry.formula);
          renderDefineNames();
        });
        return option;
      }),
    );
    projectDisabledState(defineRemoveBtn, defineSelected === null, null);
  };
  const showDefineError = (message: string): void => showDialogError(defineError, message);
  /** Add or update the name in the fields; false when it was refused. */
  const defineCommit = (): boolean => {
    const name = defineName.input.value.trim();
    if (!name) {
      showDefineError(ctx.text().errorEmptyName);
      focusAndSelectInput(defineName.input);
      return false;
    }
    const written = ctx.writeNames(() =>
      upsertDefinedName(ctx.workbook(), name, defineRange.input.value.trim(), -1),
    );
    if (!written.ok) return false;
    const result = written.value;
    if (!result.ok) {
      showDefineError(errorMessageFor(result.reason, ctx.text(), ctx.strings.commandRefusal));
      return false;
    }
    clearDialogError(defineError);
    ctx.afterMutate();
    defineSelected = null;
    renderDefineNames();
    return true;
  };
  defineShell.on(defineAddBtn, 'click', () => {
    defineCommit();
  });
  defineShell.on(defineRemoveBtn, 'click', () => {
    const entry = defineSelected;
    if (!entry) return;
    const written = ctx.writeNames(() =>
      deleteDefinedName(ctx.workbook(), entry.name, entry.localSheetId),
    );
    if (!written.ok) return;
    if (!written.value.ok) {
      showDefineError(ctx.strings.commandRefusal.engineFailed);
      return;
    }
    clearDialogError(defineError);
    ctx.afterMutate();
    defineSelected = null;
    defineName.input.value = '';
    renderDefineNames();
  });
  defineShell.on(defineActions.cancelBtn, 'click', defineShell.close);
  defineShell.on(defineActions.okBtn, 'click', () => {
    if (defineName.input.value.trim() && !defineCommit()) return;
    defineShell.close();
  });
  const openDefine = (): void => {
    const m = ctx.strings.mac.namedRangeDialog;
    defineFrame.header.textContent = m.defineTitle;
    defineShell.setAriaLabel(m.defineTitle);
    defineNamesLabel.textContent = m.defineNames;
    defineName.text.textContent = m.defineNamePrompt;
    defineRange.text.textContent = m.defineRangePrompt;
    defineSelected = null;
    defineName.input.value = '';
    // The range seeds as an absolute, sheet-qualified reference: =Sheet!$A$1:$D$7.
    defineRange.input.value = ctx.seedFormula();
    clearDialogError(defineError);
    renderDefineNames();
    defineShell.open();
    deferInitialFocus(defineShell.overlay, () => defineName.input.focus());
  };
  return { open: openDefine, dispose: () => defineShell.dispose() };
}
