import { FUNCTION_SIGNATURES, suggestFunctions } from '../commands/refs.js';
import { inheritHostTokens } from './inherit-host-tokens.js';

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
  setLabels(labels: Partial<AutocompleteLabels>): void;
  detach(): void;
}

export interface AutocompleteLabels {
  customFunction: string;
  structuredTableColumn: string;
  pickFromList: string;
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
  /** Excel-style "pick from list" source. Called for plain-text edits to
   *  surface previous text values entered above the active cell in the same
   *  column. Implementations should return values in nearest-first order
   *  (closest row first), already deduped, blanks excluded. The popover
   *  filters by case-insensitive prefix on the typed token. Omit to disable. */
  getColumnValues?: (sheet: number, col: number, beforeRow: number) => readonly string[];
  /** Address of the cell currently being edited. Required for the
   *  `getColumnValues` path so the popover knows which sheet/column/row to
   *  scan above. Omit when only function/structured-ref suggestion is needed. */
  editingAddr?: { sheet: number; row: number; col: number };
  /** Engine-driven function catalog. When the workbook exposes
   *  `functionNames()` (5/5 build), pass it through so autocomplete
   *  surfaces the full registry — not just the 98-entry curated list in
   *  `commands/refs.ts`. Falls back to the static catalog when omitted
   *  or null. */
  getFunctionNames?: () => readonly string[] | null;
  labels?: Partial<AutocompleteLabels>;
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
  /** The kind drives display labels and badges. */
  kind: 'function' | 'tableColumn' | 'column' | 'history';
}

/**
 * Function-name + structured-ref autocomplete popover. Hangs off the document
 * body — the caller calls `refresh()` from input handlers, and arrow / enter /
 * tab / escape are intercepted by checking `isOpen()` before the input's own
 * logic runs.
 */
export function attachAutocomplete(deps: AutocompleteDeps): AutocompleteHandle {
  const { input } = deps;
  let labels: AutocompleteLabels = {
    customFunction: 'Custom function',
    structuredTableColumn: 'Structured table column',
    pickFromList: 'Pick from list',
    ...deps.labels,
  };
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
    inheritHostTokens(input, root);
    root.replaceChildren();
    for (let i = 0; i < ctx.matches.length; i += 1) {
      const pick = ctx.matches[i] ?? '';
      const item = document.createElement('button');
      item.className = 'fc-autocomplete__item';
      if (i === highlight) item.classList.add('fc-autocomplete__item--active');
      item.setAttribute('role', 'option');
      item.dataset.fcKind = ctx.kind;
      item.type = 'button';

      const badge = document.createElement('span');
      badge.className = 'fc-autocomplete__badge';
      badge.textContent =
        ctx.kind === 'function' ? 'fx' : ctx.kind === 'tableColumn' ? 'tbl' : 'list';

      const main = document.createElement('span');
      main.className = 'fc-autocomplete__main';

      const name = document.createElement('span');
      name.className = 'fc-autocomplete__name';
      name.textContent = pick;

      const detail = document.createElement('span');
      detail.className = 'fc-autocomplete__detail';
      if (ctx.kind === 'function') {
        const sig = FUNCTION_SIGNATURES[pick.toUpperCase()];
        detail.textContent = sig ? `${pick}(${sig.join(', ')})` : labels.customFunction;
      } else if (ctx.kind === 'tableColumn') {
        detail.textContent = labels.structuredTableColumn;
      } else {
        detail.textContent = labels.pickFromList;
      }

      main.append(name, detail);
      item.append(badge, main);
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
    const next = computeContext(text, caret, deps);
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
    setLabels(next) {
      labels = { ...labels, ...next };
      render();
    },
    detach() {
      close();
    },
  };
}

/** Build a suggestion context for the caret position. Plain-text edits go
 *  through the column-history source (Excel's "pick from list"). Formula edits
 *  fall through to structured-ref → function/custom-name suggestions. */
function computeContext(
  text: string,
  caret: number,
  deps: AutocompleteDeps,
): SuggestionContext | null {
  if (!text.startsWith('=')) {
    const addr = deps.editingAddr;
    if (!addr || !deps.getColumnValues) return null;
    const values = deps.getColumnValues(addr.sheet, addr.col, addr.row);
    return suggestColumnHistory(text, caret, values);
  }
  const struct = suggestStructuredRef(text, caret, deps.getTables?.() ?? []);
  if (struct) return struct;
  const engineNames = deps.getFunctionNames?.();
  const fn = suggestFunctions(text, caret, 8, engineNames ? { names: engineNames } : undefined);
  const custom = suggestCustomFunctions(text, caret, deps.getCustomFunctions?.() ?? []);
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

/** Excel's column-history autocomplete: when editing a plain-text cell, match
 *  what the user has typed so far against the values already entered above in
 *  the same column. `values` must already be deduped, blanks excluded, in
 *  nearest-first order (caller's responsibility). The whole input acts as the
 *  partial token — Excel replaces the entire cell text on accept, never just
 *  a fragment. Returns null when the input is empty or no value prefix-matches. */
export function suggestColumnHistory(
  text: string,
  caret: number,
  values: readonly string[],
): SuggestionContext | null {
  // Only suggest when caret is at end and the input is a non-empty token. A
  // shorter caret (mid-edit) means the user is correcting earlier characters,
  // not extending the tail — Excel doesn't pop the list there either.
  if (caret !== text.length) return null;
  if (text.length === 0) return null;
  const lower = text.toLowerCase();
  const matches: string[] = [];
  for (const v of values) {
    if (v.length === text.length) continue; // skip exact-length (would be no-op)
    if (v.toLowerCase().startsWith(lower)) matches.push(v);
    if (matches.length >= 10) break;
  }
  if (matches.length === 0) return null;
  return {
    matches,
    tokenStart: 0,
    tokenEnd: text.length,
    insertSuffix: '',
    kind: 'column',
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
  if (/[[\],]/.test(partial)) return null;
  const lower = partial.toLowerCase();
  const matches = table.columns.filter((c) => c.toLowerCase().startsWith(lower)).slice(0, 12);
  if (matches.length === 0) return null;
  return {
    matches,
    tokenStart: bracket + 1,
    tokenEnd: caret,
    insertSuffix: ']',
    kind: 'tableColumn',
  };
}
