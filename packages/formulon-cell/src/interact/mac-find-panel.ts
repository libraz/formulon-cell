import type { MacFindPanelStrings } from '../i18n/strings/mac/find-replace.js';
import { createInteractionButton } from './chip-button.js';
import { markSurface, openerOf, returnFocus } from './dialog-shell.js';
import { openPopupMenu } from './popup-menu.js';

export type FindScope = 'sheet' | 'workbook';

export interface MacFindPanelDeps {
  host: HTMLElement;
  strings: () => MacFindPanelStrings;
  closeLabel: () => string;
  /** The search state the panel shares with the Find and Replace dialog. */
  query: () => string;
  setQuery: (text: string) => void;
  scope: () => FindScope;
  setScope: (scope: FindScope) => void;
  /** Runs one search step; true when it landed on a match. */
  step: (direction: 'next' | 'prev') => boolean;
  /** More Options > Advanced Search... and Replace...: the dialog on that tab. */
  openDialog: (tab: 'find' | 'replace') => void;
  /** The panel was hidden. */
  onClose: () => void;
}

export interface MacFindPanel {
  open(): void;
  /** Hides the panel; focus goes back to the grid unless `restoreFocus` is false. */
  close(restoreFocus?: boolean): void;
  isOpen(): boolean;
  relabel(): void;
  detach(): void;
}

/** The Mac ⌘F panel: a compact find field floating over the top right of the grid, with a scope
 *  menu, Previous / Next and a More Options menu leading to the full dialog. */
export const createMacFindPanel = (deps: MacFindPanelDeps): MacFindPanel => {
  const panel = document.createElement('div');
  let opener: HTMLElement | null = null;
  panel.className = 'fc-findpanel';
  panel.setAttribute('role', 'dialog');
  panel.setAttribute('aria-modal', 'false');
  markSurface(panel, 'floating');
  panel.hidden = true;

  const button = (className: string, text = ''): HTMLButtonElement =>
    createInteractionButton({ className, text });

  const titlebar = document.createElement('div');
  titlebar.className = 'fc-findpanel__titlebar';
  const title = document.createElement('span');
  title.className = 'fc-findpanel__title';
  const closeBtn = button('fc-findpanel__close', '×');
  titlebar.append(title, closeBtn);

  const field = document.createElement('div');
  field.className = 'fc-findpanel__field';
  const scopeBtn = button('fc-findpanel__scope');
  scopeBtn.setAttribute('aria-haspopup', 'menu');
  scopeBtn.setAttribute('aria-expanded', 'false');
  const input = document.createElement('input');
  input.type = 'text';
  input.className = 'fc-findpanel__input';
  input.spellcheck = false;
  input.autocomplete = 'off';
  const clearBtn = button('fc-findpanel__clear', '×');
  field.append(scopeBtn, input, clearBtn);

  const prevBtn = button('fc-findpanel__step fc-findpanel__prev', '‹');
  const nextBtn = button('fc-findpanel__step fc-findpanel__next', '›');
  const moreBtn = button('fc-findpanel__more', '•••');
  moreBtn.setAttribute('aria-haspopup', 'menu');
  moreBtn.setAttribute('aria-expanded', 'false');

  const row = document.createElement('div');
  row.className = 'fc-findpanel__row';
  row.append(field, prevBtn, nextBtn, moreBtn);
  panel.append(titlebar, row);
  deps.host.append(panel);

  let closeMenu: (() => void) | null = null;
  const dropMenu = (): void => {
    closeMenu?.();
    closeMenu = null;
  };

  const syncField = (): void => {
    const t = deps.strings();
    const scopeLabel = deps.scope() === 'workbook' ? t.findInWorkbook : t.findInSheet;
    input.placeholder = scopeLabel;
    input.setAttribute('aria-label', scopeLabel);
    scopeBtn.setAttribute('aria-label', scopeLabel);
    clearBtn.hidden = input.value === '';
  };

  const relabel = (): void => {
    const t = deps.strings();
    title.textContent = t.title;
    panel.setAttribute('aria-label', t.title);
    closeBtn.setAttribute('aria-label', deps.closeLabel());
    clearBtn.setAttribute('aria-label', t.clear);
    prevBtn.setAttribute('aria-label', t.previous);
    nextBtn.setAttribute('aria-label', t.next);
    moreBtn.setAttribute('aria-label', t.moreOptions);
    syncField();
  };

  const runStep = (direction: 'next' | 'prev'): void => {
    // A match leaves the text selected, ready for the next query.
    if (deps.step(direction)) {
      input.focus();
      input.select();
    }
  };

  const toggleMenu = (
    anchor: HTMLButtonElement,
    build: () => Parameters<typeof openPopupMenu>[0],
  ) => {
    const wasOpen = closeMenu !== null && anchor.getAttribute('aria-expanded') === 'true';
    dropMenu();
    if (!wasOpen) closeMenu = openPopupMenu(build());
  };

  const onScope = (): void =>
    toggleMenu(scopeBtn, () => {
      const t = deps.strings();
      const pick = (scope: FindScope) => () => {
        deps.setScope(scope);
        syncField();
        input.focus();
      };
      return {
        anchor: scopeBtn,
        items: [
          { label: t.findInSheet, checked: deps.scope() === 'sheet', onPick: pick('sheet') },
          {
            label: t.findInWorkbook,
            checked: deps.scope() === 'workbook',
            onPick: pick('workbook'),
          },
        ],
      };
    });

  const onMore = (): void =>
    toggleMenu(moreBtn, () => {
      const t = deps.strings();
      const toDialog = (tab: 'find' | 'replace') => () => {
        api.close(false);
        deps.openDialog(tab);
      };
      return {
        anchor: moreBtn,
        label: t.moreOptions,
        items: [
          { label: t.advancedSearch, onPick: toDialog('find') },
          { label: t.replace, onPick: toDialog('replace') },
        ],
      };
    });

  const onInput = (): void => {
    deps.setQuery(input.value);
    syncField();
  };
  const onClear = (): void => {
    input.value = '';
    onInput();
    input.focus();
  };
  // Return and Shift+Return both find the next match; Previous is the ‹ button only.
  const onInputKey = (e: KeyboardEvent): void => {
    if (e.key !== 'Enter' || e.isComposing) return;
    e.preventDefault();
    runStep('next');
  };
  // Keys stay in the panel; Escape closes it only while it has focus.
  const onPanelKey = (e: KeyboardEvent): void => {
    e.stopPropagation();
    if (e.key === 'Escape') {
      e.preventDefault();
      api.close();
    }
  };
  const onPrev = (): void => runStep('prev');
  const onNext = (): void => runStep('next');
  const onClose = (): void => api.close();

  closeBtn.addEventListener('click', onClose);
  scopeBtn.addEventListener('click', onScope);
  moreBtn.addEventListener('click', onMore);
  clearBtn.addEventListener('click', onClear);
  prevBtn.addEventListener('click', onPrev);
  nextBtn.addEventListener('click', onNext);
  input.addEventListener('input', onInput);
  input.addEventListener('keydown', onInputKey);
  panel.addEventListener('keydown', onPanelKey);

  const api: MacFindPanel = {
    open(): void {
      if (!panel.contains(panel.ownerDocument.activeElement))
        opener = openerOf(panel.ownerDocument);
      panel.hidden = false;
      input.value = deps.query();
      syncField();
      input.focus();
      input.select();
    },
    close(restoreFocus = true): void {
      if (panel.hidden) return;
      dropMenu();
      panel.hidden = true;
      deps.onClose();
      const target = opener;
      opener = null;
      if (restoreFocus) returnFocus(target, panel, deps.host);
    },
    isOpen: (): boolean => !panel.hidden,
    relabel,
    detach(): void {
      dropMenu();
      closeBtn.removeEventListener('click', onClose);
      scopeBtn.removeEventListener('click', onScope);
      moreBtn.removeEventListener('click', onMore);
      clearBtn.removeEventListener('click', onClear);
      prevBtn.removeEventListener('click', onPrev);
      nextBtn.removeEventListener('click', onNext);
      input.removeEventListener('input', onInput);
      input.removeEventListener('keydown', onInputKey);
      panel.removeEventListener('keydown', onPanelKey);
      panel.remove();
    },
  };
  relabel();
  return api;
};
