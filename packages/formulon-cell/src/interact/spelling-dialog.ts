import { commentAt, recordCommentChange, setComment } from '../commands/comment.js';
import { localizedDecisionReason } from '../commands/permission-reasons.js';
import {
  firstSpellingIssue,
  proofingLocale,
  replaceSpellingWord,
  replaceWholeWord,
  type SpellingIssueKind,
  type SpellingWord,
} from '../commands/spelling.js';
import { parseAddrKey } from '../engine/address.js';
import type { Addr } from '../engine/types.js';
import type { SpreadsheetInstance } from '../mount/types.js';
import { mutators } from '../store/store.js';
import { projectDisabledState } from '../toolbar/menu-a11y.js';
import { appendDialogButton, appendDialogFrame, createDialogShell } from './dialog-shell.js';
import { isSubmitEnter } from './mac-dialog-keys.js';

/** A cell's constant text or its note: what the check reads. */
interface SpellingTarget {
  readonly addr: Addr;
  readonly source: 'cell' | 'note';
}

interface Cursor {
  index: number;
  offset: number;
}

interface Issue {
  readonly target: SpellingTarget;
  readonly kind: SpellingIssueKind;
  readonly word: SpellingWord;
}

/** What one instance remembers between checks, for its lifetime. */
interface SpellingSession {
  /** Words ignored with Ignore All or added with Add, lower-cased. */
  readonly accepted: Set<string>;
  // The desktop app starts with Always suggest on and Ignore UPPERCASE off.
  alwaysSuggest: boolean;
  ignoreUppercase: boolean;
}

const sessions = new WeakMap<SpreadsheetInstance, SpellingSession>();
const openDialogs = new WeakMap<SpreadsheetInstance, () => void>();

const sessionOf = (instance: SpreadsheetInstance): SpellingSession => {
  let session = sessions.get(instance);
  if (!session) {
    session = { accepted: new Set(), alwaysSuggest: true, ignoreUppercase: false };
    sessions.set(instance, session);
  }
  return session;
};

const COMMAND_ID = 'spelling';

/** The cells and notes the check walks, by rows from the active cell, wrapping to the top. One
 *  selected cell checks the whole sheet; a selected range checks only that range. */
export function spellingTargets(instance: SpreadsheetInstance): SpellingTarget[] {
  const state = instance.store.getState();
  const { active, range, extraRanges } = state.selection;
  const sheet = state.data.sheetIndex;
  const wholeSheet =
    range.r0 === range.r1 && range.c0 === range.c1 && (extraRanges?.length ?? 0) === 0;
  const inScope = (addr: Addr): boolean =>
    addr.sheet === sheet &&
    (wholeSheet ||
      (addr.row >= range.r0 &&
        addr.row <= range.r1 &&
        addr.col >= range.c0 &&
        addr.col <= range.c1));
  const targets: SpellingTarget[] = [];
  for (const [key, cell] of state.data.cells) {
    const addr = parseAddrKey(key);
    if (addr && inScope(addr) && cell.value.kind === 'text' && !cell.formula) {
      targets.push({ addr, source: 'cell' });
    }
  }
  for (const [key, format] of state.format.formats) {
    const addr = parseAddrKey(key);
    if (addr && inScope(addr) && typeof format.comment === 'string' && format.comment !== '') {
      targets.push({ addr, source: 'note' });
    }
  }
  const wrapped = (addr: Addr): number =>
    addr.row < active.row || (addr.row === active.row && addr.col < active.col) ? 1 : 0;
  return targets.sort(
    (a, b) =>
      wrapped(a.addr) - wrapped(b.addr) ||
      a.addr.row - b.addr.row ||
      a.addr.col - b.addr.col ||
      (a.source === 'note' ? 1 : 0) - (b.source === 'note' ? 1 : 0),
  );
}

const textOf = (instance: SpreadsheetInstance, target: SpellingTarget): string | null => {
  if (target.source === 'note') return commentAt(instance.store.getState(), target.addr);
  if (instance.workbook.cellFormula(target.addr) !== null) return null;
  const value = instance.workbook.getValue(target.addr);
  return value.kind === 'text' ? value.value : null;
};

const languageName = (locale: string): string => {
  try {
    return new Intl.DisplayNames([locale], { type: 'language' }).of(locale) ?? locale;
  } catch {
    return locale;
  }
};

/** Writes new texts as one undo step: cells through the same authorization as a typed entry,
 *  notes through the note path. Returns the refusal reason, or null once written. A refusal
 *  rolls back whatever the group already applied. */
function writeTexts(
  instance: SpreadsheetInstance,
  edits: readonly { target: SpellingTarget; text: string }[],
): string | null {
  const { history, store, workbook } = instance;
  const cells = edits.filter((edit) => edit.target.source === 'cell');
  const notes = edits.filter((edit) => edit.target.source === 'note');
  const token = history.begin();
  let written = false;
  try {
    if (cells.length > 0) {
      const result = instance.commands.execute({
        type: 'cellBatch',
        operation: 'valueEdit',
        origin: 'ribbon',
        commandId: COMMAND_ID,
        changes: cells.map((edit) => ({
          addr: edit.target.addr,
          value: { kind: 'text', value: edit.text },
          formula: null,
        })),
      });
      if (result.status !== 'applied') {
        const rejection = result.rejections[0];
        return rejection
          ? localizedDecisionReason(rejection, instance.i18n.strings)
          : localizedDecisionReason({ code: 'operationDenied' }, instance.i18n.strings);
      }
    }
    const authorization = { origin: 'ribbon', commandId: COMMAND_ID } as const;
    for (const edit of notes) {
      const noted = recordCommentChange(
        history,
        store,
        workbook,
        [edit.target.addr],
        () => setComment(store, edit.target.addr, edit.text, workbook, authorization),
        authorization,
      );
      if (!noted) return '';
    }
    written = true;
    return null;
  } finally {
    if (written) history.end(token);
    else history.abort(token);
  }
}

/** Opens Review > Spelling at the active cell. Nothing opens when there is nothing to flag. */
export function openSpellingDialog(instance: SpreadsheetInstance): void {
  openDialogs.get(instance)?.();
  const session = sessionOf(instance);
  const targets = spellingTargets(instance);
  const locale = proofingLocale(instance.i18n.locale);
  const mac = instance.host.dataset.fcPlatform === 'mac';
  const t = instance.i18n.strings.spellingDialog;
  const rules = () => ({
    service: instance.spellingService(),
    locale,
    ignoreUppercase: session.ignoreUppercase,
    accepted: session.accepted,
  });
  const cursor: Cursor = { index: 0, offset: 0 };
  let issue: Issue | null = null;
  /** Every action on a word, for Undo Last: where it started, and whether it wrote the text or
   *  accepted the word for the session. */
  const actions: { at: Cursor; wrote: boolean; accepted?: string }[] = [];

  const findNext = (): Issue | null => {
    while (cursor.index < targets.length) {
      const target = targets[cursor.index] as SpellingTarget;
      const text = textOf(instance, target);
      const found = text === null ? null : firstSpellingIssue(text, cursor.offset, rules());
      if (found) return { target, kind: found.kind, word: found.word };
      cursor.index += 1;
      cursor.offset = 0;
    }
    return null;
  };

  issue = findNext();
  if (!issue) {
    instance.host.focus();
    return;
  }

  const shell = createDialogShell({
    host: instance.host,
    className: 'fc-spelling',
    ariaLabel: t.title.replace('{language}', languageName(locale)),
    onDismiss: () => close(),
  });
  shell.overlay.classList.add('fc-fmtdlg');
  const title = t.title.replace('{language}', languageName(locale));
  const { body, footer } = appendDialogFrame(shell, {
    title,
    panelClasses: ['fc-fmtdlg__panel', 'fc-spelling__panel'],
    bodyClass: 'fc-fmtdlg__body fc-spelling__body',
    footerClass: 'fc-fmtdlg__footer fc-spelling__footer',
  });

  const field = (id: string): { label: HTMLLabelElement; input: HTMLInputElement } => {
    const label = document.createElement('label');
    label.className = 'fc-spelling__label';
    label.htmlFor = id;
    const input = document.createElement('input');
    input.id = id;
    input.type = 'text';
    input.className = 'fc-fmtdlg__input fc-spelling__input';
    input.autocomplete = 'off';
    input.spellcheck = false;
    return { label, input };
  };
  const word = field('fc-spelling-word');
  word.input.readOnly = true;
  // The Mac dialog names the word in its label and has no field for it.
  word.input.hidden = mac;
  const changeTo = field('fc-spelling-change-to');
  const suggestionsLabel = document.createElement('div');
  suggestionsLabel.className = 'fc-spelling__label';
  suggestionsLabel.id = 'fc-spelling-suggestions-label';
  const suggestions = document.createElement('div');
  suggestions.className = 'fc-spelling__suggestions';
  suggestions.setAttribute('role', 'listbox');
  suggestions.setAttribute('aria-labelledby', suggestionsLabel.id);
  suggestions.tabIndex = 0;

  const fields = document.createElement('div');
  fields.className = 'fc-spelling__fields';
  fields.append(
    word.label,
    word.input,
    changeTo.label,
    changeTo.input,
    suggestionsLabel,
    suggestions,
  );
  if (mac) {
    const dictionary = document.createElement('select');
    dictionary.id = 'fc-spelling-dictionary';
    dictionary.className = 'fc-fmtdlg__select fc-spelling__dictionary';
    const custom = document.createElement('option');
    custom.textContent = t.customDictionaries;
    dictionary.append(custom);
    const dictionaryLabel = document.createElement('label');
    dictionaryLabel.className = 'fc-spelling__label';
    dictionaryLabel.htmlFor = dictionary.id;
    dictionaryLabel.textContent = t.addWordsTo;
    fields.append(dictionaryLabel, dictionary);
  }
  const buttons = document.createElement('div');
  buttons.className = 'fc-spelling__buttons';
  const button = (label: string, primary = false): HTMLButtonElement =>
    appendDialogButton(buttons, {
      label,
      variant: primary ? 'primary' : 'secondary',
    });
  // Ignore Once is the default button: Return leaves the word as it is.
  const ignoreOnce = button(t.ignoreOnce, true);
  const ignoreAll = button(t.ignoreAll);
  const add = button(t.add);
  const change = button(t.change);
  const changeAll = button(t.changeAll);
  const autoCorrect = button(t.autoCorrect);
  const suggest = button(t.suggest);
  const undoLast = button(t.undoLast);
  const cancel = button(t.cancel);
  body.append(fields, buttons);

  const check = (label: string, checked: boolean, set: (on: boolean) => void): HTMLInputElement => {
    const row = document.createElement('label');
    row.className = 'fc-spelling__check';
    const box = document.createElement('input');
    box.type = 'checkbox';
    box.checked = checked;
    shell.on(box, 'change', () => set(box.checked));
    row.append(box, document.createTextNode(label));
    footer.appendChild(row);
    return box;
  };
  check(t.alwaysSuggest, session.alwaysSuggest, (on) => {
    session.alwaysSuggest = on;
    if (issue) render(issue);
  });
  check(t.ignoreUppercase, session.ignoreUppercase, (on) => {
    session.ignoreUppercase = on;
  });

  const status = document.createElement('div');
  status.className = 'fc-spelling__status';
  status.setAttribute('role', 'status');
  footer.appendChild(status);

  const selectSuggestion = (option: HTMLElement): void => {
    for (const other of suggestions.children) other.setAttribute('aria-selected', 'false');
    option.setAttribute('aria-selected', 'true');
    changeTo.input.value = option.textContent ?? '';
  };
  const showSuggestions = (of: string): void => {
    const list = of === '' ? [] : instance.spellingService().suggest(of, locale);
    suggestions.replaceChildren(
      ...list.map((text, index) => {
        const option = document.createElement('div');
        option.className = 'fc-spelling__suggestion';
        option.setAttribute('role', 'option');
        option.id = `fc-spelling-suggestion-${index}`;
        option.setAttribute('aria-selected', 'false');
        option.textContent = text;
        return option;
      }),
    );
    const first = suggestions.firstElementChild as HTMLElement | null;
    if (first) selectSuggestion(first);
  };

  const render = (current: Issue): void => {
    const repeated = current.kind === 'repeated';
    const wordLabel = repeated ? t.repeatedWord : t.notInDictionary;
    word.label.textContent = mac ? `${wordLabel} ${current.word.word}` : wordLabel;
    word.input.value = current.word.word;
    changeTo.label.textContent = t.changeTo;
    suggestionsLabel.textContent = t.suggestions;
    status.textContent = '';
    delete status.dataset.state;
    suggestions.replaceChildren();
    changeTo.input.value = repeated ? '' : current.word.word;
    if (!repeated && session.alwaysSuggest) showSuggestions(current.word.word);
    // A repeated word is deleted, not changed, and never deleted everywhere at once.
    change.textContent = repeated ? t.delete : t.change;
    changeAll.textContent = repeated ? t.deleteAll : t.changeAll;
    projectDisabledState(ignoreAll, repeated, null);
    projectDisabledState(changeAll, repeated, null);
    // Always suggest lists the suggestions already, so Suggest has nothing to add.
    projectDisabledState(suggest, repeated || session.alwaysSuggest, null);
    projectDisabledState(undoLast, actions.length === 0, null);
    cancel.textContent = actions.some((action) => action.wrote) ? t.close : t.cancel;
    // A selected range stays selected while the active cell moves through it.
    const { range } = instance.store.getState().selection;
    const { row, col } = current.target.addr;
    if (row >= range.r0 && row <= range.r1 && col >= range.c0 && col <= range.c1) {
      mutators.setActivePreservingSelection(instance.store, current.target.addr);
    } else {
      mutators.setActive(instance.store, current.target.addr);
    }
  };

  /** Moves past the current word, then shows the next one or closes after the last. */
  const advance = (offset: number): void => {
    cursor.offset = offset;
    issue = findNext();
    if (issue) render(issue);
    else close();
  };

  const applyChange = (all: boolean, onWritten?: () => void): void => {
    if (!issue) return;
    const current = issue;
    const replacement = changeTo.input.value;
    const text = textOf(instance, current.target) ?? '';
    const edits: { target: SpellingTarget; text: string }[] = [];
    if (all) {
      for (const target of targets) {
        const before = textOf(instance, target);
        if (before === null) continue;
        const after = replaceWholeWord(before, current.word.word, replacement);
        if (after !== before) edits.push({ target, text: after });
      }
    } else {
      edits.push({
        target: current.target,
        text: replaceSpellingWord(text, current.word, replacement),
      });
    }
    const before: Cursor = { index: cursor.index, offset: current.word.start };
    const refused = writeTexts(instance, edits);
    if (refused !== null) {
      status.dataset.state = 'error';
      status.textContent =
        refused || localizedDecisionReason({ code: 'operationDenied' }, instance.i18n.strings);
      return;
    }
    onWritten?.();
    actions.push({ at: before, wrote: true });
    // Continue after the replacement: the text up to the word, changed the same way.
    const head = text.slice(0, current.word.end);
    advance(
      all
        ? replaceWholeWord(head, current.word.word, replacement).length
        : replaceSpellingWord(head, current.word, replacement).length,
    );
  };

  /** Leaves the current word as it is, remembering where it was for Undo Last. */
  const skip = (accepted?: string): void => {
    if (!issue) return;
    actions.push({ at: { index: cursor.index, offset: issue.word.start }, wrote: false, accepted });
    if (accepted !== undefined) session.accepted.add(accepted);
    advance(issue.word.end);
  };
  shell.on(ignoreOnce, 'click', () => skip());
  shell.on(ignoreAll, 'click', () => {
    if (issue) skip(issue.word.word.toLowerCase());
  });
  shell.on(add, 'click', () => {
    if (!issue) return;
    instance.spellingService().addWord?.(issue.word.word, locale);
    skip(issue.word.word.toLowerCase());
  });
  // Return presses the default button; the overlay hears it even where a clicked button keeps no
  // focus (WebKit).
  shell.on(shell.overlay, 'keydown', (rawEvent) => {
    const event = rawEvent as KeyboardEvent;
    if (!isSubmitEnter(event)) return;
    event.preventDefault();
    ignoreOnce.click();
  });
  shell.on(change, 'click', () => applyChange(false));
  shell.on(changeAll, 'click', () => applyChange(true));
  // AutoCorrect keeps the word and its correction in the AutoCorrect list, then changes it.
  shell.on(autoCorrect, 'click', () => {
    if (!issue || changeTo.input.value === '') return;
    const { word: typo } = issue.word;
    const correction = changeTo.input.value;
    applyChange(false, () => instance.autoCorrect.add(typo, correction));
  });
  shell.on(suggest, 'click', () => showSuggestions(changeTo.input.value));
  shell.on(undoLast, 'click', () => {
    const previous = actions.at(-1);
    if (!previous || (previous.wrote && !instance.history.undo())) return;
    actions.pop();
    if (previous.accepted !== undefined) session.accepted.delete(previous.accepted);
    cursor.index = previous.at.index;
    advance(previous.at.offset);
  });
  shell.on(cancel, 'click', () => close());
  shell.on(suggestions, 'click', (event) => {
    const option = (event.target as HTMLElement).closest<HTMLElement>('[role="option"]');
    if (option) selectSuggestion(option);
  });
  shell.on(suggestions, 'dblclick', (event) => {
    const option = (event.target as HTMLElement).closest<HTMLElement>('[role="option"]');
    if (!option) return;
    selectSuggestion(option);
    applyChange(false);
  });
  shell.on(suggestions, 'keydown', (event) => {
    const key = (event as KeyboardEvent).key;
    if (key !== 'ArrowDown' && key !== 'ArrowUp') return;
    event.preventDefault();
    const options = Array.from(suggestions.children) as HTMLElement[];
    const at = options.findIndex((option) => option.getAttribute('aria-selected') === 'true');
    const next =
      options[Math.max(0, Math.min(options.length - 1, at + (key === 'ArrowDown' ? 1 : -1)))];
    if (next) {
      selectSuggestion(next);
      suggestions.setAttribute('aria-activedescendant', next.id);
    }
  });

  let closed = false;
  function close(): void {
    if (closed) return;
    closed = true;
    openDialogs.delete(instance);
    shell.close();
    shell.dispose();
    instance.host.focus();
  }
  openDialogs.set(instance, close);

  render(issue);
  shell.open();
  ignoreOnce.focus({ preventScroll: true });
}

/** Closes an open spelling dialog, as unmounting does. */
export function closeSpellingDialog(instance: SpreadsheetInstance): void {
  openDialogs.get(instance)?.();
}
