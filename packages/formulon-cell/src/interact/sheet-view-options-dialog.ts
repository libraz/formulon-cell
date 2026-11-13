import {
  activateSheetView,
  DEFAULT_SHEET_VIEW_ID,
  deleteSheetView,
  duplicateSheetView,
  exitSheetView,
  nextSheetViewName,
  renameSheetView,
  sheetViewNameProblem,
} from '../commands/sheet-views.js';
import type { SpreadsheetInstance } from '../mount/types.js';
import { showConfirm, showPrompt } from '../toolbar/dialogs/prompt.js';
import { projectDisabledState } from '../toolbar/menu-a11y.js';
import {
  appendDialogActions,
  appendDialogButton,
  appendDialogFrame,
  createDialogShell,
} from './dialog-shell.js';

const openDialogs = new WeakMap<SpreadsheetInstance, () => void>();

export interface SheetViewOptionsDialogOptions {
  /** Called after a switch or deletion changes the rows and columns shown. */
  onViewChanged?: () => void;
}

/** Opens Sheet View Options: Default and the sheet's saved views, with Rename, Duplicate and
 *  Delete above the list and Switch to as the default button. */
export function openSheetViewOptionsDialog(
  instance: SpreadsheetInstance,
  opts: SheetViewOptionsDialogOptions = {},
): void {
  openDialogs.get(instance)?.();
  const { store, history } = instance;
  const t = instance.i18n.strings.mac.sheetViews;
  const shell = createDialogShell({
    host: instance.host,
    className: 'fc-sheetviewopts',
    ariaLabel: t.optionsTitle,
    onDismiss: () => close(),
  });
  shell.overlay.classList.add('fc-fmtdlg');
  const { body, footer } = appendDialogFrame(shell, {
    title: t.optionsTitle,
    panelClasses: ['fc-fmtdlg__panel', 'fc-sheetviewopts__panel'],
    bodyClass: 'fc-fmtdlg__body fc-sheetviewopts__body',
    footerClass: 'fc-fmtdlg__footer fc-sheetviewopts__footer',
  });

  const actions = document.createElement('div');
  actions.className = 'fc-sheetviewopts__actions';
  body.appendChild(actions);
  const renameBtn = appendDialogButton(actions, { label: t.rename, variant: 'secondary' });
  const duplicateBtn = appendDialogButton(actions, { label: t.duplicate, variant: 'secondary' });
  const deleteBtn = appendDialogButton(actions, { label: t.delete, variant: 'secondary' });

  const list = document.createElement('div');
  list.className = 'fc-sheetviewopts__list';
  list.setAttribute('role', 'listbox');
  list.setAttribute('aria-label', t.optionsTitle);
  list.tabIndex = 0;
  body.appendChild(list);

  const { cancelBtn, okBtn: switchBtn } = appendDialogActions(footer, {
    cancelLabel: t.cancel,
    okLabel: t.switchTo,
  });

  const rows = (): { id: string; name: string }[] => {
    const { sheetViews, data } = store.getState();
    return [
      { id: DEFAULT_SHEET_VIEW_ID, name: t.defaultView },
      ...sheetViews.views
        .filter((view) => view.sheet === data.sheetIndex)
        .map((view) => ({ id: view.id, name: view.name })),
    ];
  };
  let selected = store.getState().sheetViews.activeViewId ?? DEFAULT_SHEET_VIEW_ID;
  const render = (): void => {
    const all = rows();
    if (!all.some((row) => row.id === selected)) selected = DEFAULT_SHEET_VIEW_ID;
    list.replaceChildren(
      ...all.map((row, index) => {
        const option = document.createElement('div');
        option.className = 'fc-sheetviewopts__entry';
        option.setAttribute('role', 'option');
        option.id = `fc-sheetviewopts-entry-${index}`;
        option.dataset.viewId = row.id;
        option.setAttribute('aria-selected', String(row.id === selected));
        option.textContent = row.name;
        return option;
      }),
    );
    // Default can be switched to and copied, never deleted. Provisional.
    projectDisabledState(deleteBtn, selected === DEFAULT_SHEET_VIEW_ID, null);
  };
  const nameOf = (id: string): string => rows().find((row) => row.id === id)?.name ?? '';

  shell.on(list, 'click', (event) => {
    const option = (event.target as HTMLElement).closest<HTMLElement>('[role="option"]');
    if (!option?.dataset.viewId) return;
    selected = option.dataset.viewId;
    render();
  });
  shell.on(list, 'dblclick', () => switchTo());

  const switchTo = (): void => {
    if (selected === DEFAULT_SHEET_VIEW_ID) exitSheetView(store, history);
    else activateSheetView(store, selected, history);
    close();
    opts.onViewChanged?.();
  };
  shell.on(switchBtn, 'click', switchTo);
  shell.on(cancelBtn, 'click', () => close());

  shell.on(renameBtn, 'click', () => {
    const id = selected;
    void showPrompt({
      title: t.renameTitle,
      label: t.renameTitle,
      initial: nameOf(id),
      okLabel: t.rename,
      cancelLabel: t.cancel,
      validate: (name) => {
        if (id === DEFAULT_SHEET_VIEW_ID) return t.renameDefault;
        const problem = sheetViewNameProblem(store, name, id);
        if (problem === 'invalid') return t.nameInvalid;
        return problem === 'taken' ? t.nameTaken : null;
      },
    }).then((name) => {
      if (name === null || closed) return;
      renameSheetView(store, id, name, history);
      render();
    });
  });

  shell.on(duplicateBtn, 'click', () => {
    const base = nameOf(selected);
    let n = 2;
    while (sheetViewNameProblem(store, `${base} ${n}`)) n += 1;
    duplicateSheetView(
      store,
      selected,
      nextSheetViewName(store, instance.i18n.strings.viewToolbar.views).id,
      `${base} ${n}`,
      history,
    );
    render();
  });

  shell.on(deleteBtn, 'click', () => {
    const id = selected;
    if (id === DEFAULT_SHEET_VIEW_ID) return;
    void showConfirm({
      title: t.deleteTitle,
      headline: t.deleteTitle,
      message: t.deleteMessage.replace('{name}', nameOf(id)),
      alert: true,
      okLabel: t.delete,
      cancelLabel: t.cancel,
      destructive: true,
    }).then((ok) => {
      if (!ok || closed) return;
      deleteSheetView(store, id, history);
      render();
      opts.onViewChanged?.();
    });
  });

  let closed = false;
  function close(): void {
    if (closed) return;
    closed = true;
    openDialogs.delete(instance);
    shell.close();
    shell.dispose();
  }
  openDialogs.set(instance, close);

  render();
  shell.open();
  list.focus({ preventScroll: true });
}

/** Closes an open Sheet View Options window, as unmounting does. */
export function closeSheetViewOptionsDialog(instance: SpreadsheetInstance): void {
  openDialogs.get(instance)?.();
}
