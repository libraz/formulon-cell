import { suggestFunctions } from '../commands/refs.js';

export interface AutocompleteHandle {
  /** Re-evaluate the suggestion list against the current input value/caret.
   *  Hides the popover when there's nothing to suggest. */
  refresh(): void;
  /** Programmatically close the popover. */
  close(): void;
  /** True when the popover is visible. Callers consult this on Enter/Tab/Esc
   *  to decide whether to consume the key for selection. */
  isOpen(): boolean;
  /** Move the highlight up/down. No-op when closed. */
  move(delta: 1 | -1): void;
  /** Insert the highlighted suggestion into the input at the partial token's
   *  position. Returns true when it actually inserted (so callers know to
   *  preventDefault on the originating key). */
  acceptHighlighted(): boolean;
  detach(): void;
}

export interface AutocompleteDeps {
  /** The textarea/input being edited. */
  input: HTMLInputElement | HTMLTextAreaElement;
  /** Called after the popover has rewritten input.value so consumers can
   *  re-sync mirror state (e.g. formula-bar ref highlighting). */
  onAfterInsert?: () => void;
}

/**
 * Function-name autocomplete popover. Hangs off the document body — the
 * caller calls `refresh()` from input handlers, and arrow/enter/tab/escape
 * are intercepted by checking `isOpen()` before the input's own logic runs.
 */
export function attachAutocomplete(deps: AutocompleteDeps): AutocompleteHandle {
  const { input } = deps;
  let root: HTMLDivElement | null = null;
  let matches: string[] = [];
  let tokenStart = 0;
  let highlight = 0;

  const close = (): void => {
    if (!root) return;
    root.remove();
    root = null;
    matches = [];
    tokenStart = 0;
    highlight = 0;
  };

  const positionUnderCaret = (el: HTMLDivElement): void => {
    const rect = input.getBoundingClientRect();
    el.style.left = `${rect.left}px`;
    el.style.top = `${rect.bottom + 2}px`;
    el.style.minWidth = `${Math.max(180, rect.width / 2)}px`;
  };

  const render = (): void => {
    if (!root) {
      root = document.createElement('div');
      root.className = 'fc-autocomplete';
      root.setAttribute('role', 'listbox');
      document.body.appendChild(root);
    }
    root.innerHTML = '';
    for (let i = 0; i < matches.length; i += 1) {
      const item = document.createElement('div');
      item.className = 'fc-autocomplete__item';
      if (i === highlight) item.classList.add('fc-autocomplete__item--active');
      item.setAttribute('role', 'option');
      item.textContent = matches[i] ?? '';
      // Use mousedown (not click) so the input doesn't blur first.
      item.addEventListener('mousedown', (e) => {
        e.preventDefault();
        highlight = i;
        acceptHighlighted();
      });
      root.appendChild(item);
    }
    positionUnderCaret(root);
  };

  const refresh = (): void => {
    const text = input.value;
    const caret = input.selectionStart ?? text.length;
    const sug = suggestFunctions(text, caret);
    if (!sug || sug.matches.length === 0) {
      close();
      return;
    }
    matches = sug.matches;
    tokenStart = sug.tokenStart;
    if (highlight >= matches.length) highlight = 0;
    render();
  };

  const move = (delta: 1 | -1): void => {
    if (!root || matches.length === 0) return;
    highlight = (highlight + delta + matches.length) % matches.length;
    render();
  };

  const acceptHighlighted = (): boolean => {
    if (!root || matches.length === 0) return false;
    const pick = matches[highlight];
    if (!pick) return false;
    const before = input.value.slice(0, tokenStart);
    const after = input.value.slice(input.selectionStart ?? input.value.length);
    const insert = `${pick}(`;
    input.value = before + insert + after;
    const caret = before.length + insert.length;
    input.setSelectionRange(caret, caret);
    input.focus();
    close();
    deps.onAfterInsert?.();
    return true;
  };

  return {
    refresh,
    close,
    isOpen: () => root != null,
    move,
    acceptHighlighted,
    detach() {
      close();
    },
  };
}
