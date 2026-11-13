import { setZerosVisible } from '../commands/view.js';
import type { Strings } from '../i18n/strings.js';
import type { SpreadsheetInstance } from '../mount/types.js';
import { createExcelRibbonSvg } from '../toolbar/excel-ribbon-icons.js';
import { projectDisabledState } from '../toolbar/menu-a11y.js';
import {
  appendDialogButton,
  appendDialogFrame,
  appendDialogIconButton,
  createDialogShell,
} from './dialog-shell.js';

type PrefStrings = Strings['preferencesDialog'];

/** A pane of the Preferences window: its name, the Show All group it sits in, and its settings. */
interface PreferencePane {
  readonly id: string;
  readonly label: (t: PrefStrings) => string;
  readonly group: (t: PrefStrings) => string;
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
    group: (t) => t.authoringGroup,
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

/** Show All, or one pane, as the window's history records them. */
type Page = PreferencePane | 'all';

/** Opens the Preferences window on its first pane. Every setting applies as it changes; the
 *  window has no buttons, only a close control and the desktop's toolbar: Back and Forward
 *  through the pages shown, Show All listing the panes under their group headings, the shown
 *  pane's name, and a search field narrowing Show All to the panes whose names match. */
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

  const toolbar = document.createElement('div');
  toolbar.className = 'fc-prefs__toolbar';
  toolbar.setAttribute('role', 'toolbar');
  body.appendChild(toolbar);
  const historyGroup = document.createElement('div');
  historyGroup.className = 'fc-prefs__history';
  historyGroup.setAttribute('role', 'group');
  historyGroup.setAttribute('aria-label', t.history);
  toolbar.appendChild(historyGroup);
  const toolButton = (parent: HTMLElement, label: string, glyph: string): HTMLButtonElement => {
    const button = appendDialogIconButton(parent, {
      label: '',
      ariaLabel: label,
      title: label,
      baseClass: 'fc-prefs__tool',
    });
    button.append(createExcelRibbonSvg(glyph, 'fc-prefs__glyph') ?? '');
    return button;
  };
  const backBtn = toolButton(historyGroup, t.back, 'navBack');
  const forwardBtn = toolButton(historyGroup, t.forward, 'navForward');
  const showAllBtn = toolButton(toolbar, t.showAll, 'showAll');
  const heading = document.createElement('span');
  heading.className = 'fc-prefs__heading';
  toolbar.appendChild(heading);
  const search = document.createElement('input');
  search.type = 'search';
  search.className = 'fc-prefs__search';
  search.placeholder = t.search;
  search.setAttribute('aria-label', t.search);
  toolbar.appendChild(search);

  const all = document.createElement('div');
  all.className = 'fc-prefs__all';
  body.appendChild(all);
  const groups = new Map<string, { section: HTMLElement; buttons: HTMLButtonElement[] }>();
  const paneButtons = PANES.map((pane) => {
    const name = pane.group(t);
    let group = groups.get(name);
    if (!group) {
      const section = document.createElement('section');
      section.className = 'fc-prefs__group';
      const groupTitle = document.createElement('h3');
      groupTitle.className = 'fc-prefs__group-title';
      groupTitle.textContent = name;
      section.appendChild(groupTitle);
      all.appendChild(section);
      group = { section, buttons: [] };
      groups.set(name, group);
    }
    const button = appendDialogButton(group.section, {
      label: pane.label(t),
      baseClass: 'fc-prefs__pane-button',
    });
    button.dataset.pane = pane.id;
    group.buttons.push(button);
    shell.on(button, 'click', () => go(pane));
    return { pane, button };
  });

  const panels = PANES.map((pane) => {
    const panel = document.createElement('div');
    panel.className = 'fc-prefs__pane';
    panel.dataset.pane = pane.id;
    panel.setAttribute('role', 'group');
    panel.setAttribute('aria-label', pane.label(t));
    panel.hidden = true;
    body.appendChild(panel);
    pane.render(instance, panel, shell.on);
    return { pane, panel };
  });

  // Pages shown, oldest first; Back and Forward move through them as a browser does.
  const history: Page[] = [];
  let at = -1;

  function render(): void {
    const page = history[at] ?? 'all';
    const name = page === 'all' ? t.windowTitle : page.label(t);
    title.textContent = name;
    shell.setAriaLabel(name);
    heading.textContent = page === 'all' ? '' : name;
    all.hidden = page !== 'all';
    for (const { pane, panel } of panels) panel.hidden = pane !== page;
    projectDisabledState(backBtn, at <= 0, null);
    projectDisabledState(forwardBtn, at >= history.length - 1, null);
    filter();
  }

  function go(page: Page): void {
    if (history[at] === page) return;
    history.splice(at + 1, history.length, page);
    at = history.length - 1;
    render();
  }

  // Search matches pane names only, without changing case. Provisional.
  function filter(): void {
    const query = search.value.trim().toLocaleLowerCase();
    for (const { pane, button } of paneButtons) {
      button.hidden = query !== '' && !pane.label(t).toLocaleLowerCase().includes(query);
    }
    for (const { section, buttons } of groups.values()) {
      section.hidden = buttons.every((button) => button.hidden);
    }
  }

  shell.on(backBtn, 'click', () => {
    if (at <= 0) return;
    at -= 1;
    render();
  });
  shell.on(forwardBtn, 'click', () => {
    if (at >= history.length - 1) return;
    at += 1;
    render();
  });
  shell.on(showAllBtn, 'click', () => go('all'));
  shell.on(search, 'input', () => {
    if (search.value.trim() !== '' && history[at] !== 'all') go('all');
    else filter();
  });

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
  go(first ?? 'all');
  shell.open();
}

/** Closes an open Preferences window, as unmounting does. */
export function closePreferencesDialog(instance: SpreadsheetInstance): void {
  openDialogs.get(instance)?.();
}
