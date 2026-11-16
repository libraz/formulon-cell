import { projectDisabledState } from '../menu-a11y.js';
import { createDialogSelect } from './form-controls.js';
import {
  appendDialogActions,
  appendDialogButton,
  createDialogShell,
  installDialogLifecycle,
  mountDialog,
} from './shell.js';

export interface CustomListsDialogLabels {
  title: string;
  newList: string;
  lists: string;
  entries: string;
  add: string;
  delete: string;
  hint: string;
  ok: string;
  cancel: string;
}

export interface CustomListsDialogOptions {
  /** The spreadsheet host; anchors the dialog and picks its platform layout. */
  host?: HTMLElement;
  labels: CustomListsDialogLabels;
  /** The locale's built-in lists, shown first and never edited or deleted. */
  builtIns: readonly (readonly string[])[];
  /** The user's lists, which Add and Delete change. */
  userLists: readonly (readonly string[])[];
  /** Receives the user's lists after Add or Delete changes them. */
  onChange: (next: readonly (readonly string[])[]) => void;
  /** The list to select on open, matched by its entries. */
  initial?: readonly string[];
}

type ListRef = { kind: 'new' } | { kind: 'builtIn' | 'user'; index: number };

const sameList = (a: readonly string[], b: readonly string[]): boolean =>
  a.length === b.length && a.every((entry, i) => entry === b[i]);

/** Entries typed one per line, or separated by commas, as the desktop accepts them. */
const parseEntries = (text: string): string[] =>
  text
    .split(/[\n,]/)
    .map((entry) => entry.trim())
    .filter((entry) => entry !== '');

/** The Custom Lists dialog: NEW LIST, the built-in lists and the user's lists, with List entries,
 *  Add and Delete. Resolves with the selected list's entries on OK, and `null` on Cancel. */
export const showCustomListsDialog = (opts: CustomListsDialogOptions): Promise<string[] | null> =>
  new Promise<string[] | null>((resolve) => {
    const t = opts.labels;
    const shell = createDialogShell({ host: opts.host, title: t.title, bodyVariant: 'app' });
    shell.panel.classList.add('fc-customlists');
    const doc = shell.body.ownerDocument;
    let userLists = opts.userLists.map((list) => [...list]);

    const listsLabel = doc.createElement('span');
    listsLabel.className = 'fc-customlists__label';
    listsLabel.textContent = t.lists;
    const lists = createDialogSelect([], '', { ariaLabel: t.lists });
    lists.classList.add('fc-customlists__lists');
    lists.size = 8;
    const entriesLabel = doc.createElement('label');
    entriesLabel.className = 'fc-customlists__label';
    entriesLabel.textContent = t.entries;
    const entries = doc.createElement('textarea');
    entries.className = 'fc-customlists__entries';
    entries.rows = 8;
    entries.setAttribute('aria-label', t.entries);
    const hint = doc.createElement('p');
    hint.className = 'fc-customlists__hint';
    hint.textContent = t.hint;
    const buttons = doc.createElement('div');
    buttons.className = 'fc-customlists__buttons';
    const addBtn = appendDialogButton(buttons, { label: t.add });
    addBtn.classList.add('fc-customlists__add');
    const deleteBtn = appendDialogButton(buttons, { label: t.delete });
    deleteBtn.classList.add('fc-customlists__delete');
    shell.body.append(listsLabel, lists, entriesLabel, entries, buttons, hint);

    const refOf = (value: string): ListRef => {
      if (value === 'new') return { kind: 'new' };
      const [kind, index] = value.split(':');
      return { kind: kind === 'b' ? 'builtIn' : 'user', index: Number(index) };
    };
    const entriesOf = (ref: ListRef): readonly string[] =>
      ref.kind === 'new'
        ? []
        : ((ref.kind === 'builtIn' ? opts.builtIns : userLists)[ref.index] ?? []);

    const render = (selected: string): void => {
      lists.replaceChildren();
      const add = (value: string, text: string): void => {
        const option = doc.createElement('option');
        option.value = value;
        option.textContent = text;
        lists.appendChild(option);
      };
      add('new', t.newList);
      for (const [i, list] of opts.builtIns.entries()) add(`b:${i}`, list.join(', '));
      for (const [i, list] of userLists.entries()) add(`u:${i}`, list.join(', '));
      lists.value = [...lists.options].some((o) => o.value === selected) ? selected : 'new';
      sync();
    };
    // A built-in list shows its entries read-only and cannot be deleted, as on the desktop.
    const sync = (): void => {
      const ref = refOf(lists.value);
      entries.value = entriesOf(ref).join('\n');
      entries.readOnly = ref.kind === 'builtIn';
      projectDisabledState(addBtn, ref.kind === 'builtIn', null);
      projectDisabledState(deleteBtn, ref.kind !== 'user', null);
    };
    lists.addEventListener('change', sync);

    const commit = (next: string[][], selected: string): void => {
      userLists = next;
      opts.onChange(next.map((list) => [...list]));
      render(selected);
    };
    /** Add: a new list from NEW LIST, or the edited entries of the selected user list. */
    const addEntries = (): void => {
      const ref = refOf(lists.value);
      const typed = parseEntries(entries.value);
      if (ref.kind === 'builtIn' || typed.length === 0) return;
      if (ref.kind === 'user') {
        commit(
          userLists.map((list, i) => (i === ref.index ? typed : list)),
          `u:${ref.index}`,
        );
        return;
      }
      commit([...userLists, typed], `u:${userLists.length}`);
    };
    addBtn.addEventListener('click', addEntries);
    deleteBtn.addEventListener('click', () => {
      const ref = refOf(lists.value);
      if (ref.kind !== 'user') return;
      commit(
        userLists.filter((_, i) => i !== ref.index),
        'new',
      );
    });

    const initial = opts.initial;
    const builtInIndex = initial ? opts.builtIns.findIndex((list) => sameList(list, initial)) : -1;
    const userIndex = initial ? userLists.findIndex((list) => sameList(list, initial)) : -1;
    render(builtInIndex >= 0 ? `b:${builtInIndex}` : userIndex >= 0 ? `u:${userIndex}` : 'new');

    const { cancelBtn, okBtn } = appendDialogActions(shell.footer, {
      cancelLabel: t.cancel,
      okLabel: t.ok,
    });
    const choose = (): string[] | null => {
      if (refOf(lists.value).kind === 'new') addEntries();
      const ref = refOf(lists.value);
      return ref.kind === 'new' ? null : [...entriesOf(ref)];
    };
    const lifecycle = installDialogLifecycle<string[] | null>({
      shell,
      resolve,
      onCancel: () => null,
      onSubmit: () => lifecycle.finish(choose()),
      buttons: { okBtn, cancelBtn },
    });
    mountDialog(shell, lists, { immediate: true });
  });
