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

/** Subset of `WorkbookHandle.getTables()` consumed by the structured-ref
 *  suggestion path. Only `name` and `columns` are needed. */
export interface AutocompleteTable {
  name: string;
  columns: string[];
}

export interface AutocompleteDeps {
  /** The textarea/input being edited. */
  input: HTMLInputElement | HTMLTextAreaElement;
  /** Called after the popover has rewritten input.value so consumers can
   *  re-sync mirror state (e.g. formula-bar ref highlighting). */
  onAfterInsert?: () => void;
  /** Snapshot of available tables for structured-ref completion. Returning
   *  `[]` (or omitting) disables the `Table[…` path; functions still complete. */
  getTables?: () => readonly AutocompleteTable[];
  /** Upper-cased custom function names to merge into the function-suggestion
   *  list. Pulled from `inst.formula.list()` so consumer registrations
   *  appear alongside the engine's built-ins. */
  getCustomFunctions?: () => readonly string[];
}

interface SuggestionContext {
  matches: string[];
  /** Character offset where the partial token starts (inclusive). */
  tokenStart: number;
  /** Character offset where the partial token ends (exclusive). */
  tokenEnd: number;
  /** What gets inserted when the user accepts a match — defaults to `${match}`
   *  for refs and `${match}(` for functions. */
  insertSuffix: string;
  /** The kind drives display labels (and would let us paint badges later). */
  kind: 'function' | 'column';
}

/**
 * Function-name + structured-ref autocomplete popover. Hangs off the document
 * body — the caller calls `refresh()` from input handlers, and arrow / enter /
 * tab / escape are intercepted by checking `isOpen()` before the input's own
 * logic runs.
 */
export function attachAutocomplete(deps: AutocompleteDeps): AutocompleteHandle {
  const { input } = deps;
  let root: HTMLDivElement | null = null;
  let ctx: SuggestionContext | null = null;
  let highlight = 0;

  const close = (): void => {
    if (!root) return;
    root.remove();
    root = null;
    ctx = null;
    highlight = 0;
  };

  const positionUnderCaret = (el: HTMLDivElement): void => {
    const rect = input.getBoundingClientRect();
    el.style.left = `${rect.left}px`;
    el.style.top = `${rect.bottom + 2}px`;
    el.style.minWidth = `${Math.max(180, rect.width / 2)}px`;
  };

  const render = (): void => {
    if (!ctx) return;
    if (!root) {
      root = document.createElement('div');
      root.className = 'fc-autocomplete';
      root.setAttribute('role', 'listbox');
      document.body.appendChild(root);
    }
    root.innerHTML = '';
    for (let i = 0; i < ctx.matches.length; i += 1) {
      const item = document.createElement('div');
      item.className = 'fc-autocomplete__item';
      if (i === highlight) item.classList.add('fc-autocomplete__item--active');
      item.setAttribute('role', 'option');
      item.dataset.fcKind = ctx.kind;
      item.textContent = ctx.matches[i] ?? '';
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
    const next = computeContext(text, caret, deps.getTables, deps.getCustomFunctions);
    if (!next) {
      close();
      return;
    }
    ctx = next;
    if (highlight >= ctx.matches.length) highlight = 0;
    render();
  };

  const move = (delta: 1 | -1): void => {
    if (!ctx || ctx.matches.length === 0) return;
    highlight = (highlight + delta + ctx.matches.length) % ctx.matches.length;
    render();
  };

  const acceptHighlighted = (): boolean => {
    if (!ctx || ctx.matches.length === 0) return false;
    const pick = ctx.matches[highlight];
    if (!pick) return false;
    const before = input.value.slice(0, ctx.tokenStart);
    const after = input.value.slice(ctx.tokenEnd);
    const insert = `${pick}${ctx.insertSuffix}`;
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

/** Build a suggestion context for the caret position. Structured refs win
 *  when the caret is inside a `Table[…` token; otherwise we fall back to
 *  function-name suggestions. */
function computeContext(
  text: string,
  caret: number,
  getTables?: () => readonly AutocompleteTable[],
  getCustomFunctions?: () => readonly string[],
): SuggestionContext | null {
  const struct = suggestStructuredRef(text, caret, getTables?.() ?? []);
  if (struct) return struct;
  const fn = suggestFunctions(text, caret);
  const custom = suggestCustomFunctions(text, caret, getCustomFunctions?.() ?? []);
  if (!fn && !custom) return null;
  // Merge custom names ahead of built-ins so user code surfaces first.
  const seen = new Set<string>();
  const merged: string[] = [];
  for (const n of custom?.matches ?? []) {
    if (seen.has(n)) continue;
    seen.add(n);
    merged.push(n);
  }
  for (const n of fn?.matches ?? []) {
    if (seen.has(n)) continue;
    seen.add(n);
    merged.push(n);
  }
  const tokenStart = fn?.tokenStart ?? custom?.tokenStart ?? caret;
  return {
    matches: merged.slice(0, 8),
    tokenStart,
    tokenEnd: caret,
    insertSuffix: '(',
    kind: 'function',
  };
}

/** Filter the consumer's custom function list against the partial token at
 *  `caret`. Mirrors `suggestFunctions` but against a runtime-supplied list. */
function suggestCustomFunctions(
  text: string,
  caret: number,
  names: readonly string[],
): { token: string; tokenStart: number; matches: string[] } | null {
  if (!text.startsWith('=') || names.length === 0) return null;
  let i = caret - 1;
  while (i >= 0) {
    const ch = text[i] ?? '';
    if (/[A-Za-z0-9_]/.test(ch)) i -= 1;
    else break;
  }
  const tokenStart = i + 1;
  const token = text.slice(tokenStart, caret);
  if (token.length < 1 || !/^[A-Za-z]/.test(token)) return null;
  const upper = token.toUpperCase();
  const matches = names.filter((n) => n.startsWith(upper)).slice(0, 8);
  if (matches.length === 0) return null;
  return { token, tokenStart, matches };
}

const NAME_RE = /[A-Za-z_][A-Za-z0-9_]*$/;

/** Detect a `TableName[partial` token immediately before the caret and return
 *  matching column names. The token is allowed to contain spaces inside the
 *  bracket so multi-word column names work. Returns null when the caret is
 *  outside such a token. */
export function suggestStructuredRef(
  text: string,
  caret: number,
  tables: readonly AutocompleteTable[],
): SuggestionContext | null {
  if (!text.startsWith('=') || tables.length === 0) return null;
  // Look for the most recent unmatched `[` to the left of the caret, bounded
  // to the same formula segment (no `]` between caret and the opening bracket).
  let bracket = -1;
  for (let i = caret - 1; i >= 0; i -= 1) {
    const ch = text[i];
    if (ch === ']' || ch === '"') return null;
    if (ch === '[') {
      bracket = i;
      break;
    }
  }
  if (bracket <= 0) return null;
  const before = text.slice(0, bracket);
  const nameMatch = NAME_RE.exec(before);
  if (!nameMatch) return null;
  const tableName = nameMatch[0];
  const table = tables.find((t) => t.name.toLowerCase() === tableName.toLowerCase());
  if (!table || table.columns.length === 0) return null;
  const partial = text.slice(bracket + 1, caret);
  // Reject when the partial already contains characters that close the ref.
  if (/[\[\],]/.test(partial)) return null;
  const lower = partial.toLowerCase();
  const matches = table.columns.filter((c) => c.toLowerCase().startsWith(lower)).slice(0, 12);
  if (matches.length === 0) return null;
  return {
    matches,
    tokenStart: bracket + 1,
    tokenEnd: caret,
    insertSuffix: ']',
    kind: 'column',
  };
}
