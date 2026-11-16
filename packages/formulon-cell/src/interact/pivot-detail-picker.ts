import { PivotAxis } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { Strings } from '../i18n/strings.js';
import { projectDisabledState } from '../toolbar/menu-a11y.js';
import { createInteractionChipButton } from './chip-button.js';
import {
  appendDialogActions,
  appendDialogFrame,
  createDialogShell,
  deferInitialFocus,
} from './dialog-shell.js';
import { isSubmitEnter } from './mac-dialog-keys.js';
import { isInnermostRowItem, type PivotRowItemRef } from './pivot-detail.js';
import { commitPivotChange, type PivotCommitContext } from './pivot-field-settings-dialog-apply.js';

/** One row of the Show Detail picker. */
export interface PivotShowDetailField {
  fieldIdx: number;
  name: string;
  /** False for a column or filter field: listed, but not offered as a choice. */
  pickable: boolean;
}

/** The fields the Show Detail picker lists for `ref`: every field off the row area, in field
 *  order. */
export function pivotShowDetailFields(
  wb: WorkbookHandle,
  ref: PivotRowItemRef,
): PivotShowDetailField[] {
  const definition = wb.getPivotDefinition(ref.sheet, ref.pivotIndex);
  if (!definition) return [];
  return definition.fields
    .filter((field) => field.axis !== PivotAxis.Row)
    .map((field) => ({
      fieldIdx: field.index,
      name: field.customName || field.sourceName,
      pickable: field.axis !== PivotAxis.Col && field.axis !== PivotAxis.Page,
    }));
}

/** The picker opens on an innermost item nested under another row field, when it has a field
 *  to offer. */
export const canShowPivotDetail = (wb: WorkbookHandle, ref: PivotRowItemRef): boolean =>
  isInnermostRowItem(ref) &&
  ref.level > 0 &&
  pivotShowDetailFields(wb, ref).some((field) => field.pickable);

/** Adds `fieldIdx` as the row field after the item's own, under every item; one undo step. */
export function showPivotDetailField(
  ctx: PivotCommitContext,
  ref: PivotRowItemRef,
  fieldIdx: number,
): boolean {
  const order = ref.rowFieldOrder.filter((i) => i !== fieldIdx);
  order.splice(ref.level + 1, 0, fieldIdx);
  return commitPivotChange(
    ctx,
    ref.sheet,
    ref.pivotIndex,
    () =>
      ctx.wb.setPivotFieldAxis(ref.sheet, ref.pivotIndex, fieldIdx, PivotAxis.Row) &&
      ctx.wb.setPivotRowFieldOrder(ref.sheet, ref.pivotIndex, order),
  );
}

/** Open the Show Detail picker of an innermost row item. OK stays disabled until a field is
 *  chosen; Cancel leaves the report alone. False when `canShowPivotDetail` refuses the item. */
export function openPivotShowDetailDialog(
  ctx: PivotCommitContext & { host: HTMLElement; strings: Strings },
  ref: PivotRowItemRef,
): boolean {
  if (!canShowPivotDetail(ctx.wb, ref)) return false;
  const fields = pivotShowDetailFields(ctx.wb, ref);
  const t = ctx.strings.pivotShowDetailDialog;
  const shell = createDialogShell({
    host: ctx.host,
    className: 'fc-pivotdetail',
    ariaLabel: t.title,
    onDismiss: () => shell.dispose(),
  });
  shell.overlay.classList.add('fc-fmtdlg');
  const frame = appendDialogFrame(shell, {
    title: t.title,
    panelClasses: ['fc-fmtdlg__panel', 'fc-pivotdetail__panel'],
  });
  const prompt = document.createElement('p');
  prompt.className = 'fc-tb__dlg__message';
  prompt.textContent = t.prompt;
  const list = document.createElement('div');
  list.className = 'fc-movecopy__list';
  list.setAttribute('role', 'listbox');
  list.setAttribute('aria-label', t.prompt);
  frame.body.append(prompt, list);
  const { cancelBtn, okBtn } = appendDialogActions(frame.footer, {
    cancelLabel: ctx.strings.pivotTableDialog.cancel,
    okLabel: ctx.strings.pivotTableDialog.ok,
  });

  let chosen = -1;
  const options = fields.map((field, index) => {
    const option = createInteractionChipButton({
      className: 'fc-movecopy__item',
      label: field.name,
      role: 'option',
      selected: false,
      tabIndex: -1,
    });
    projectDisabledState(option, !field.pickable, null);
    shell.on(option, 'click', () => choose(index));
    list.appendChild(option);
    return option;
  });
  const choose = (index: number): void => {
    if (!fields[index]?.pickable) return;
    chosen = index;
    options.forEach((option, i) => {
      option.setAttribute('aria-selected', i === chosen ? 'true' : 'false');
      option.tabIndex = i === chosen ? 0 : -1;
    });
    options[chosen]?.focus({ preventScroll: true });
    projectDisabledState(okBtn, false, null);
  };
  const step = (delta: number): void => {
    for (let i = chosen + delta; i >= 0 && i < fields.length; i += delta) {
      if (fields[i]?.pickable) {
        choose(i);
        return;
      }
    }
  };
  projectDisabledState(okBtn, true, null);
  const firstPickable = options.find((_, i) => fields[i]?.pickable);
  if (firstPickable) firstPickable.tabIndex = 0;

  const run = (): void => {
    const field = fields[chosen];
    if (!field) return;
    showPivotDetailField(ctx, ref, field.fieldIdx);
    shell.dispose();
  };
  shell.on(okBtn, 'click', run);
  shell.on(cancelBtn, 'click', () => shell.dispose());
  shell.on(list, 'keydown', (event) => {
    const e = event as KeyboardEvent;
    if (e.key === 'ArrowDown') step(1);
    else if (e.key === 'ArrowUp') step(-1);
    else return;
    e.preventDefault();
  });
  shell.on(shell.overlay, 'keydown', (event) => {
    const e = event as KeyboardEvent;
    if (!isSubmitEnter(e)) return;
    e.preventDefault();
    run();
  });
  shell.open();
  const focusList = (): void => (firstPickable ?? cancelBtn).focus({ preventScroll: true });
  focusList();
  deferInitialFocus(shell.overlay, focusList);
  return true;
}
