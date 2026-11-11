import type { AutoCorrectEntry } from '../commands/autocorrect.js';
import { proofingLocale } from '../commands/spelling.js';
import type { SpreadsheetInstance } from '../mount/types.js';
import { projectDisabledState } from '../toolbar/menu-a11y.js';
import { appendDialogButton, appendDialogFrame, createDialogShell } from './dialog-shell.js';

const openDialogs = new WeakMap<SpreadsheetInstance, () => void>();

/** The name of the language the lists are for, in the UI language. */
const languageName = (locale: string): string => {
  try {
    return (
      new Intl.DisplayNames([locale], { type: 'language' }).of(proofingLocale(locale)) ?? locale
    );
  } catch {
    return locale;
  }
};

/** Opens the AutoCorrect options: the replacement list and its switches. Every change applies at
 *  once and reaches the host through `autoCorrect.onChange`. */
export function openAutoCorrectDialog(instance: SpreadsheetInstance): void {
  openDialogs.get(instance)?.();
  const store = instance.autoCorrect;
  const t = instance.i18n.strings.autoCorrectDialog;
  const title = t.title.replace('{language}', languageName(instance.i18n.locale));
  const shell = createDialogShell({
    host: instance.host,
    className: 'fc-autocorrect',
    ariaLabel: title,
    onDismiss: () => close(),
  });
  shell.overlay.classList.add('fc-fmtdlg');
  const { body, footer } = appendDialogFrame(shell, {
    title,
    panelClasses: ['fc-fmtdlg__panel', 'fc-autocorrect__panel'],
    bodyClass: 'fc-fmtdlg__body fc-autocorrect__body',
    footerClass: 'fc-fmtdlg__footer fc-autocorrect__footer',
  });

  const check = (label: string, checked: boolean, set: (on: boolean) => void): void => {
    const row = document.createElement('label');
    row.className = 'fc-autocorrect__check';
    const box = document.createElement('input');
    box.type = 'checkbox';
    box.checked = checked;
    shell.on(box, 'change', () => set(box.checked));
    row.append(box, document.createTextNode(label));
    body.appendChild(row);
  };
  // The desktop pane's order: hyperlinks, days, two capitals, sentences, then the list switch.
  check(t.replaceHyperlinks, store.state().replaceHyperlinks, (on) =>
    store.update({ replaceHyperlinks: on }),
  );
  check(t.capitalizeDays, store.state().capitalizeDays, (on) =>
    store.update({ capitalizeDays: on }),
  );
  check(t.correctTwoInitialCapitals, store.state().correctTwoInitialCapitals, (on) =>
    store.update({ correctTwoInitialCapitals: on }),
  );
  check(t.capitalizeSentences, store.state().capitalizeSentences, (on) =>
    store.update({ capitalizeSentences: on }),
  );
  check(t.replaceAsYouType, store.state().replaceAsYouType, (on) =>
    store.update({ replaceAsYouType: on }),
  );

  const field = (id: string, label: string): HTMLInputElement => {
    const wrap = document.createElement('label');
    wrap.className = 'fc-autocorrect__field';
    wrap.htmlFor = id;
    const text = document.createElement('span');
    text.textContent = label;
    const input = document.createElement('input');
    input.id = id;
    input.type = 'text';
    input.className = 'fc-fmtdlg__input';
    input.autocomplete = 'off';
    input.spellcheck = false;
    wrap.append(text, input);
    fields.appendChild(wrap);
    return input;
  };
  const fields = document.createElement('div');
  fields.className = 'fc-autocorrect__fields';
  body.appendChild(fields);
  const from = field('fc-autocorrect-replace', t.replace);
  const to = field('fc-autocorrect-with', t.with);

  const list = document.createElement('div');
  list.className = 'fc-autocorrect__list';
  list.setAttribute('role', 'listbox');
  list.setAttribute('aria-label', title);
  list.title = t.description;
  list.tabIndex = 0;
  body.appendChild(list);

  const actions = document.createElement('div');
  actions.className = 'fc-autocorrect__actions';
  body.appendChild(actions);
  const add = appendDialogButton(actions, { label: t.add, variant: 'secondary' });
  const remove = appendDialogButton(actions, { label: t.delete, variant: 'secondary' });

  const closeButton = appendDialogButton(footer, { label: t.close, variant: 'primary' });

  const entryOf = (key: string): AutoCorrectEntry | undefined =>
    store.state().entries.find((entry) => entry.from === key);
  const render = (): void => {
    const selected = from.value;
    list.replaceChildren(
      ...store.state().entries.map((entry, index) => {
        const option = document.createElement('div');
        option.className = 'fc-autocorrect__entry';
        option.setAttribute('role', 'option');
        option.id = `fc-autocorrect-entry-${index}`;
        option.dataset.from = entry.from;
        option.setAttribute('aria-selected', entry.from === selected ? 'true' : 'false');
        const left = document.createElement('span');
        left.textContent = entry.from;
        const right = document.createElement('span');
        right.textContent = entry.to;
        option.append(left, right);
        return option;
      }),
    );
    const existing = entryOf(selected);
    projectDisabledState(add, selected === '' || existing?.to === to.value, null);
    projectDisabledState(remove, existing === undefined, null);
  };

  shell.on(from, 'input', () => {
    const existing = entryOf(from.value);
    if (existing) to.value = existing.to;
    render();
  });
  shell.on(to, 'input', render);
  shell.on(list, 'click', (event) => {
    const option = (event.target as HTMLElement).closest<HTMLElement>('[role="option"]');
    const entry = option?.dataset.from === undefined ? undefined : entryOf(option.dataset.from);
    if (!entry) return;
    from.value = entry.from;
    to.value = entry.to;
    render();
  });
  shell.on(add, 'click', () => {
    if (from.value === '') return;
    store.add(from.value, to.value);
    render();
  });
  shell.on(remove, 'click', () => {
    store.remove(from.value);
    from.value = '';
    to.value = '';
    render();
  });
  shell.on(closeButton, 'click', () => close());

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
  from.focus({ preventScroll: true });
}

/** Closes open AutoCorrect options, as unmounting does. */
export function closeAutoCorrectDialog(instance: SpreadsheetInstance): void {
  openDialogs.get(instance)?.();
}
