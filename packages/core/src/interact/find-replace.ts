import {
  type FindMatch,
  type FindOptions,
  applySubstitution,
  findAll,
  findNext,
  replaceAll,
  replaceOne,
} from '../commands/find.js';
import { formatCell } from '../engine/value.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { type Strings, defaultStrings } from '../i18n/strings.js';
import { type SpreadsheetStore, mutators } from '../store/store.js';

export interface FindReplaceDeps {
  host: HTMLElement;
  store: SpreadsheetStore;
  wb: WorkbookHandle;
  strings?: Strings;
  onAfterCommit?: () => void;
}

export interface FindReplaceHandle {
  open(): void;
  close(): void;
  detach(): void;
}

export function attachFindReplace(deps: FindReplaceDeps): FindReplaceHandle {
  const { host, store, wb } = deps;
  const strings = deps.strings ?? defaultStrings;
  const t = strings.findReplace;

  const overlay = document.createElement('div');
  overlay.className = 'fc-find';
  overlay.setAttribute('role', 'dialog');
  overlay.setAttribute('aria-label', t.title);
  overlay.hidden = true;

  const findRow = document.createElement('div');
  findRow.className = 'fc-find__row';
  const findInput = document.createElement('input');
  findInput.type = 'text';
  findInput.placeholder = t.findLabel;
  findInput.setAttribute('aria-label', t.findLabel);
  findInput.spellcheck = false;
  findInput.autocomplete = 'off';
  const pill = document.createElement('span');
  pill.className = 'fc-find__pill';
  pill.textContent = '0 / 0';
  findRow.append(findInput, pill);

  const replaceRow = document.createElement('div');
  replaceRow.className = 'fc-find__row';
  const replaceInput = document.createElement('input');
  replaceInput.type = 'text';
  replaceInput.placeholder = t.replaceLabel;
  replaceInput.setAttribute('aria-label', t.replaceLabel);
  replaceInput.spellcheck = false;
  replaceInput.autocomplete = 'off';
  replaceRow.append(replaceInput);

  const buttonRow = document.createElement('div');
  buttonRow.className = 'fc-find__row';
  const prevBtn = makeBtn(t.prev, t.prev);
  const nextBtn = makeBtn(t.next, t.next);
  const replaceBtn = makeBtn(t.replaceOne, t.replaceOne);
  const replaceAllBtn = makeBtn(t.replaceAll, t.replaceAll);
  const caseLabel = document.createElement('label');
  caseLabel.className = 'fc-find__row';
  const caseToggle = document.createElement('input');
  caseToggle.type = 'checkbox';
  caseToggle.id = `fc-find-case-${Math.random().toString(36).slice(2, 8)}`;
  const caseText = document.createElement('span');
  caseText.textContent = t.matchCase;
  caseLabel.append(caseToggle, caseText);
  const closeBtn = makeBtn('×', t.close);
  buttonRow.append(prevBtn, nextBtn, replaceBtn, replaceAllBtn, caseLabel, closeBtn);

  overlay.append(findRow, replaceRow, buttonRow);
  host.appendChild(overlay);

  let lastQuery = '';
  let currentMatch: FindMatch | null = null;

  const opts = (): FindOptions => ({
    query: findInput.value,
    caseSensitive: caseToggle.checked,
  });

  const updatePill = (text?: string): void => {
    if (text !== undefined) {
      pill.textContent = text;
      return;
    }
    const o = opts();
    if (!o.query) {
      pill.textContent = '0 / 0';
      return;
    }
    const all = findAll(store.getState(), o);
    if (all.length === 0) {
      pill.textContent = '0 / 0';
      return;
    }
    let idx = -1;
    if (currentMatch) {
      const cur = currentMatch.addr;
      idx = all.findIndex(
        (m) => m.addr.row === cur.row && m.addr.col === cur.col && m.addr.sheet === cur.sheet,
      );
    }
    pill.textContent = `${idx >= 0 ? idx + 1 : 0} / ${all.length}`;
  };

  const step = (direction: 'next' | 'prev'): void => {
    const o = opts();
    if (!o.query) {
      currentMatch = null;
      updatePill();
      return;
    }
    const state = store.getState();
    const from = currentMatch ? currentMatch.addr : state.selection.active;
    const m = findNext(state, o, from, direction);
    currentMatch = m;
    if (m) mutators.setActive(store, m.addr);
    updatePill();
  };

  const doReplace = (): void => {
    const o = opts();
    if (!o.query || !currentMatch) {
      step('next');
      return;
    }
    if (wb.cellFormula(currentMatch.addr) === null) {
      const cur = formatCell(wb.getValue(currentMatch.addr));
      const next = applySubstitution(cur, o, replaceInput.value);
      if (next !== cur) {
        replaceOne(wb, currentMatch, next);
        deps.onAfterCommit?.();
      }
    }
    step('next');
  };

  const doReplaceAll = (): void => {
    const o = opts();
    if (!o.query) {
      updatePill();
      return;
    }
    const n = replaceAll(store.getState(), wb, o, replaceInput.value);
    if (n > 0) deps.onAfterCommit?.();
    currentMatch = null;
    updatePill(`${n} replaced`);
  };

  const onFindKey = (e: KeyboardEvent): void => {
    if (e.key === 'Enter') {
      e.preventDefault();
      step(e.shiftKey ? 'prev' : 'next');
    } else if (e.key === 'Escape') {
      e.preventDefault();
      api.close();
    }
  };

  const onFindInput = (): void => {
    // Reset the active match so the next step starts from the active selection
    // rather than a stale match that no longer matches.
    currentMatch = null;
    updatePill();
  };

  const onReplaceKey = (e: KeyboardEvent): void => {
    if (e.key === 'Escape') {
      e.preventDefault();
      api.close();
    }
  };

  const onCaseChange = (): void => {
    currentMatch = null;
    updatePill();
  };

  const onOverlayKey = (e: KeyboardEvent): void => {
    // Keep every keystroke in the overlay isolated from the grid's keyboard
    // handler, which would otherwise treat printable keys as begin-edit and
    // Enter as a new edit on the active cell.
    e.stopPropagation();
    if (e.key === 'Escape') {
      e.preventDefault();
      api.close();
    }
  };

  const onPrevClick = (): void => step('prev');
  const onNextClick = (): void => step('next');
  const onCloseClick = (): void => api.close();

  prevBtn.addEventListener('click', onPrevClick);
  nextBtn.addEventListener('click', onNextClick);
  replaceBtn.addEventListener('click', doReplace);
  replaceAllBtn.addEventListener('click', doReplaceAll);
  closeBtn.addEventListener('click', onCloseClick);
  findInput.addEventListener('keydown', onFindKey);
  findInput.addEventListener('input', onFindInput);
  replaceInput.addEventListener('keydown', onReplaceKey);
  caseToggle.addEventListener('change', onCaseChange);
  overlay.addEventListener('keydown', onOverlayKey);

  const api: FindReplaceHandle = {
    open(): void {
      overlay.hidden = false;
      findInput.value = lastQuery;
      currentMatch = null;
      updatePill();
      // Defer focus so the keystroke that opened us doesn't get swallowed.
      requestAnimationFrame(() => {
        findInput.focus();
        findInput.select();
      });
    },
    close(): void {
      lastQuery = findInput.value;
      overlay.hidden = true;
      currentMatch = null;
      host.focus();
    },
    detach(): void {
      prevBtn.removeEventListener('click', onPrevClick);
      nextBtn.removeEventListener('click', onNextClick);
      replaceBtn.removeEventListener('click', doReplace);
      replaceAllBtn.removeEventListener('click', doReplaceAll);
      closeBtn.removeEventListener('click', onCloseClick);
      findInput.removeEventListener('keydown', onFindKey);
      findInput.removeEventListener('input', onFindInput);
      replaceInput.removeEventListener('keydown', onReplaceKey);
      caseToggle.removeEventListener('change', onCaseChange);
      overlay.removeEventListener('keydown', onOverlayKey);
      overlay.remove();
    },
  };

  return api;
}

function makeBtn(label: string, ariaLabel: string): HTMLButtonElement {
  const b = document.createElement('button');
  b.type = 'button';
  b.className = 'fc-find__btn';
  b.textContent = label;
  b.setAttribute('aria-label', ariaLabel);
  return b;
}
