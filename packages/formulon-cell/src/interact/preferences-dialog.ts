import { setZerosVisible } from '../commands/view.js';
import type { Strings } from '../i18n/strings.js';
import type { SpreadsheetInstance } from '../mount/types.js';
import {
  appendDialogFrame,
  appendDialogIconButton,
  appendDialogTabPair,
  createDialogShell,
} from './dialog-shell.js';

/** A pane of the Preferences window: its name in the pane bar and its settings. */
interface PreferencePane {
  readonly id: string;
  readonly label: (t: Strings['preferencesDialog']) => string;
  readonly render: (instance: SpreadsheetInstance, body: HTMLElement, on: Listen) => void;
}

type Listen = (target: EventTarget, event: string, handler: EventListener) => void;

const checkbox = (
  body: HTMLElement,
  on: Listen,
  opts: { label: string; tip: string; checked: boolean; set: (on: boolean) => void },
): void => {
  const row = document.createElement('label');
  row.className = 'fc-prefs__check';
  row.title = opts.tip;
  const box = document.createElement('input');
  box.type = 'checkbox';
  box.checked = opts.checked;
  on(box, 'change', () => opts.set(box.checked));
  row.append(box, document.createTextNode(opts.label));
  body.appendChild(row);
};

/** The panes in the window's order. */
const PANES: readonly PreferencePane[] = [
  {
    id: 'view',
    label: (t) => t.viewPane,
    render(instance, body, on) {
      const t = instance.i18n.strings.preferencesDialog;
      checkbox(body, on, {
        label: t.zeroValues,
        tip: t.zeroValuesTip,
        checked: instance.store.getState().ui.showZeros !== false,
        set: (show) => setZerosVisible(instance.store, show, instance.workbook),
      });
    },
  },
];

const openDialogs = new WeakMap<SpreadsheetInstance, () => void>();

/** Opens the Preferences window on its first pane. Every setting applies as it changes; the
 *  window is titled with the shown pane's name and has no buttons, only a close control. */
export function openPreferencesDialog(instance: SpreadsheetInstance): void {
  openDialogs.get(instance)?.();
  const t = instance.i18n.strings.preferencesDialog;
  const close = instance.i18n.strings.watchPanel.close;
  const shell = createDialogShell({
    host: instance.host,
    className: 'fc-prefs',
    ariaLabel: '',
    onDismiss: () => dismiss(),
  });
  shell.overlay.classList.add('fc-fmtdlg');
  const { header, body, footer } = appendDialogFrame(shell, {
    title: '',
    panelClasses: ['fc-fmtdlg__panel', 'fc-prefs__panel'],
    headerClass: 'fc-fmtdlg__header fc-prefs__header',
    bodyClass: 'fc-fmtdlg__body fc-prefs__body',
  });
  footer.remove();
  const title = document.createElement('span');
  header.replaceChildren(title);
  const closeBtn = appendDialogIconButton(header, {
    label: '',
    ariaLabel: close,
    title: close,
    baseClass: 'fc-fmtdlg__close',
  });
  shell.on(closeBtn, 'click', () => dismiss());

  const bar = document.createElement('div');
  bar.className = 'fc-prefs__panes';
  bar.setAttribute('role', 'tablist');
  body.appendChild(bar);
  const pairs = PANES.map((pane) => {
    const pair = appendDialogTabPair(bar, body, {
      id: pane.id,
      label: pane.label(t),
      tabId: `fc-prefs-tab-${pane.id}`,
      panelId: `fc-prefs-panel-${pane.id}`,
      tabClass: 'fc-prefs__pane-tab',
      panelClass: 'fc-prefs__pane',
      tabDatasetKey: 'pane',
      panelDatasetKey: 'pane',
    });
    pane.render(instance, pair.panel, shell.on);
    shell.on(pair.button, 'click', () => show(pane));
    return { pane, ...pair };
  });

  function show(pane: PreferencePane): void {
    const name = pane.label(t);
    title.textContent = name;
    shell.setAriaLabel(name);
    for (const pair of pairs) {
      const selected = pair.pane === pane;
      pair.button.setAttribute('aria-selected', String(selected));
      pair.button.tabIndex = selected ? 0 : -1;
      pair.panel.hidden = !selected;
    }
  }

  let closed = false;
  function dismiss(): void {
    if (closed) return;
    closed = true;
    openDialogs.delete(instance);
    shell.close();
    shell.dispose();
  }
  openDialogs.set(instance, dismiss);

  const [first] = PANES;
  if (first) show(first);
  shell.open();
}

/** Closes an open Preferences window, as unmounting does. */
export function closePreferencesDialog(instance: SpreadsheetInstance): void {
  openDialogs.get(instance)?.();
}
