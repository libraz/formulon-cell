import { formulaCaretContext } from '../commands/formula-edit-context.js';
import { FUNCTION_SIGNATURES, suggestFunctions } from '../commands/refs.js';
import type { Addr, CellValue } from '../engine/types.js';
import { createExcelRibbonSvg } from '../toolbar/excel-ribbon-icons.js';
import { createInteractionButton } from './chip-button.js';
import { createFormulaCaretPositioner } from './formula-helper-position.js';
import { overlayPortalFor } from './overlay-portal.js';

let autocompleteSeq = 0;

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
  /** Move to an edge or by one visible page. No-op when closed. */
  navigate?(action: 'first' | 'last' | 'pageUp' | 'pageDown'): void;
  /** Close and suppress the same text/caret until it changes. */
  dismiss?(): void;
  /** Toggle suggestions for the current text/caret, overriding dismissal. */
  toggle?(): void;
  /** Hide while an IME composition is in progress. */
  setComposing?(composing: boolean): void;
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

/** A defined or table name offered next to function names. `detail` is the
 *  secondary line (the name's refers-to text). */
export interface AutocompleteName {
  name: string;
  detail: string;
}

/** Names usable from `sheetIndex`: workbook-scoped and that sheet's own names,
 *  plus table names. Names scoped to another sheet are excluded. */
export function collectAutocompleteNames(
  definedNames: Iterable<{ name: string; formula: string; localSheetId: number; hidden?: boolean }>,
  tables: readonly AutocompleteTable[],
  sheetIndex: number,
): AutocompleteName[] {
  const out: AutocompleteName[] = [];
  for (const dn of definedNames) {
    if (!dn.name.trim() || dn.hidden) continue;
    if (dn.localSheetId !== -1 && dn.localSheetId !== sheetIndex) continue;
    out.push({ name: dn.name, detail: dn.formula.startsWith('=') ? dn.formula : `=${dn.formula}` });
  }
  for (const t of tables) out.push({ name: t.name, detail: t.columns.join(', ') });
  return out;
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
  /** Defined and table names suggested alongside functions in formulas; build
   *  the list with `collectAutocompleteNames`. Omit to disable. */
  getNames?: () => readonly AutocompleteName[];
  /** Upper-cased custom function names to merge into the function-suggestion
   *  list. Pulled from `inst.formula.list()` so consumer registrations
   *  appear alongside the engine's built-ins. */
  getCustomFunctions?: () => readonly string[];
  /** Argument labels for a custom function (`CustomFunctionMeta.args`), shown
   *  as its signature. Returning `undefined` falls back to the generic label. */
  getCustomFunctionArgs?: (name: string) => readonly string[] | undefined;
  /** Spreadsheet-style "pick from list" source. Called for plain-text edits to
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
  /** Display names for the active function locale. When omitted, the
   * canonical names above are used. The values are intentionally kept as the
   * insertion text so aliases and Unicode names remain visible to the user. */
  getFunctionDisplayNames?: () => readonly string[] | null;
  /** Resolve a display alias before looking up its canonical signature. */
  resolveFunctionName?: (rawName: string) => string | null;
  /** Localized function ScreenTip details, keyed by canonical function name. */
  getFunctionDetails?: (
    canonicalName: string,
  ) => { signature?: string; description?: string } | undefined;
  /** Read the workbook's live formula list separator. */
  getListSeparator?: () => string;
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
  kind: 'function' | 'tableColumn' | 'column' | 'history' | 'name';
  /** Per-match kind and detail when the list mixes functions and names. */
  kinds?: SuggestionContext['kind'][];
  details?: string[];
}

interface AutocompleteItemButtonInput {
  id: string;
  kind: SuggestionContext['kind'];
  label: string;
}

function autocompleteBadgeText(kind: SuggestionContext['kind']): string {
  if (kind === 'name') return 'nm';
  return kind === 'function' ? 'fx' : kind === 'tableColumn' ? 'tbl' : 'list';
}

function createAutocompleteItemButton(input: AutocompleteItemButtonInput): HTMLButtonElement {
  const item = createInteractionButton({
    className: ['fc-autocomplete__item', input.kind === 'name' ? 'fc-autocomplete__item--name' : '']
      .filter(Boolean)
      .join(' '),
    role: 'option',
    selected: false,
    tabIndex: -1,
    dataset: { fcKind: input.kind },
  });
  item.id = input.id;

  const glyph =
    input.kind === 'function' ? createExcelRibbonSvg('function', 'fc-autocomplete__glyph') : null;
  if (glyph) {
    item.appendChild(glyph);
  } else {
    const badge = document.createElement('span');
    badge.className = 'fc-autocomplete__badge';
    badge.textContent = autocompleteBadgeText(input.kind);
    item.appendChild(badge);
  }

  const name = document.createElement('span');
  name.className = 'fc-autocomplete__name';
  name.textContent = input.label;
  item.appendChild(name);
  return item;
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
  let ctxValue: string | null = null;
  let ctxCaret = -1;
  let highlight = 0;
  let highlightKey: string | null = null;
  let rootId = '';
  let listbox: HTMLDivElement | null = null;
  let screenTip: HTMLDivElement | null = null;
  let isComposing = false;
  let dismissed: { value: string; caret: number } | null = null;
  const optionsByKey = new Map<string, HTMLButtonElement>();
  let nextOptionIndex = 0;
  const positioner = createFormulaCaretPositioner(input);

  const close = (): void => {
    root?.remove();
    root = null;
    listbox = null;
    screenTip = null;
    ctx = null;
    ctxValue = null;
    ctxCaret = -1;
    highlight = 0;
    highlightKey = null;
    optionsByKey.clear();
    input.setAttribute('aria-expanded', 'false');
    input.removeAttribute('aria-controls');
    input.removeAttribute('aria-activedescendant');
  };

  const candidateKey = (kind: SuggestionContext['kind'], label: string): string =>
    `${kind}:${label.toLocaleLowerCase()}`;

  const positionHelpNearActiveRow = (): void => {
    if (!root || !listbox || !screenTip) return;
    const width = input.ownerDocument.defaultView?.innerWidth ?? 1024;
    const rootRect = root.getBoundingClientRect();
    const scaleX =
      root.offsetWidth > 0 && rootRect.width > 0 ? rootRect.width / root.offsetWidth : 1;
    const effectiveWidth = width / scaleX;
    const isNarrow = effectiveWidth <= 520;
    root.classList.toggle('fc-autocomplete--narrow', isNarrow);
    screenTip.style.marginTop = isNarrow ? '4px' : '0px';
    if (!screenTip.textContent?.trim()) return;
    if (isNarrow) {
      return;
    }
    const active = listbox.querySelector<HTMLElement>('.fc-autocomplete__item--active');
    const listHeight = listbox.clientHeight || 200;
    const helpHeight = screenTip.offsetHeight || 56;
    const activeTop = (active?.offsetTop ?? 0) - listbox.scrollTop;
    const top = Math.max(0, Math.min(activeTop, listHeight - helpHeight));
    screenTip.style.marginTop = `${top}px`;
  };

  const updateActive = (): void => {
    if (!ctx || !listbox || !screenTip) return;
    const options = listbox.querySelectorAll<HTMLButtonElement>('[role="option"]');
    const active = options[highlight];
    options.forEach((option, index) => {
      const isActive = index === highlight;
      option.classList.toggle('fc-autocomplete__item--active', isActive);
      option.setAttribute('aria-selected', String(isActive));
      option.removeAttribute('aria-describedby');
    });
    if (active) {
      input.setAttribute('aria-activedescendant', active.id);
      if (screenTip.textContent?.trim()) active.setAttribute('aria-describedby', screenTip.id);
      if (typeof active.scrollIntoView === 'function') active.scrollIntoView({ block: 'nearest' });
    } else {
      input.removeAttribute('aria-activedescendant');
    }
    highlightKey = active
      ? candidateKey(ctx.kinds?.[highlight] ?? ctx.kind, ctx.matches[highlight] ?? '')
      : null;
    positionHelpNearActiveRow();
    if (root) {
      positioner.position(
        root,
        'below',
        overlayPortalFor(input).querySelector<HTMLElement>('.fc-arghelper'),
      );
    }
  };

  const detailsFor = (
    pick: string,
    kind: SuggestionContext['kind'],
    index: number,
  ): { signature?: string; description?: string } => {
    if (kind === 'name') return { description: ctx?.details?.[index] ?? '' };
    if (kind === 'tableColumn') return { description: labels.structuredTableColumn };
    if (kind === 'history' || kind === 'column') return { description: labels.pickFromList };
    if (!ctx || kind !== 'function') return {};

    const canonical = deps.resolveFunctionName?.(pick) ?? pick.toLocaleUpperCase();
    const supplied = deps.getFunctionDetails?.(canonical);
    const separator = deps.getListSeparator?.() === ';' ? ';' : ',';
    const signature =
      supplied?.signature ??
      (() => {
        const args = FUNCTION_SIGNATURES[canonical] ?? deps.getCustomFunctionArgs?.(canonical);
        return args ? `${pick}(${args.join(`${separator} `)})` : undefined;
      })();
    const isCustom = (deps.getCustomFunctions?.() ?? []).some(
      (name) => name.toLocaleLowerCase() === canonical.toLocaleLowerCase(),
    );
    return {
      signature: signature || (isCustom ? labels.customFunction : undefined),
      description: supplied?.description,
    };
  };

  const updateScreenTip = (): void => {
    if (!ctx || !screenTip) return;
    const pick = ctx.matches[highlight] ?? '';
    const kind = ctx.kinds?.[highlight] ?? ctx.kind;
    const details = detailsFor(pick, kind, highlight);
    screenTip.replaceChildren();
    if (details.signature) {
      const signature = document.createElement('div');
      signature.className = 'fc-autocomplete__signature';
      signature.textContent = details.signature;
      screenTip.appendChild(signature);
    }
    if (details.description) {
      const description = document.createElement('div');
      description.className = 'fc-autocomplete__description';
      description.textContent = details.description;
      screenTip.appendChild(description);
    }
    root?.classList.toggle(
      'fc-autocomplete--has-help',
      Boolean(details.signature || details.description),
    );
  };

  const render = (): void => {
    if (!ctx) return;
    if (!root) {
      root = document.createElement('div');
      autocompleteSeq += 1;
      rootId = `fc-autocomplete-${autocompleteSeq}`;
      root.id = `${rootId}-panel`;
      root.className = 'fc-autocomplete';
      listbox = document.createElement('div');
      listbox.id = rootId;
      listbox.className = 'fc-autocomplete__list';
      listbox.setAttribute('role', 'listbox');
      listbox.addEventListener('scroll', () => {
        positionHelpNearActiveRow();
        if (root) {
          positioner.position(
            root,
            'below',
            overlayPortalFor(input).querySelector<HTMLElement>('.fc-arghelper'),
          );
        }
      });
      screenTip = document.createElement('div');
      screenTip.id = `${rootId}-screen-tip`;
      screenTip.className = 'fc-autocomplete__help';
      screenTip.setAttribute('role', 'tooltip');
      root.append(listbox, screenTip);
      overlayPortalFor(input).appendChild(root);
    }
    input.setAttribute('aria-expanded', 'true');
    input.setAttribute('aria-controls', listbox?.id ?? rootId);
    const present = new Set<string>();
    for (let i = 0; i < ctx.matches.length; i += 1) {
      const pick = ctx.matches[i] ?? '';
      const kind = ctx.kinds?.[i] ?? ctx.kind;
      const key = candidateKey(kind, pick);
      present.add(key);
      let item = optionsByKey.get(key);
      if (!item) {
        item = createAutocompleteItemButton({
          id: `${rootId}-option-${nextOptionIndex}`,
          kind,
          label: pick,
        });
        nextOptionIndex += 1;
        optionsByKey.set(key, item);
        item.addEventListener('mousedown', (e) => {
          e.preventDefault();
          const option = e.currentTarget as HTMLButtonElement;
          const index = Number(option.dataset.fcIndex ?? -1);
          if (index < 0 || !ctx || !ctx.matches[index]) return;
          highlight = index;
          updateScreenTip();
          updateActive();
        });
        item.addEventListener('mousemove', (e) => {
          const option = e.currentTarget as HTMLButtonElement;
          const index = Number(option.dataset.fcIndex ?? -1);
          if (index < 0 || !ctx || !ctx.matches[index] || highlight === index) return;
          highlight = index;
          updateScreenTip();
          updateActive();
        });
        item.addEventListener('dblclick', (e) => {
          e.preventDefault();
          const option = e.currentTarget as HTMLButtonElement;
          const index = Number(option.dataset.fcIndex ?? -1);
          if (index < 0 || !ctx || !ctx.matches[index]) return;
          highlight = index;
          updateScreenTip();
          updateActive();
          acceptHighlighted();
        });
      }
      item.dataset.fcIndex = String(i);
      const currentAtIndex = listbox?.children[i] ?? null;
      if (item !== currentAtIndex) listbox?.insertBefore(item, currentAtIndex);
    }
    for (const [key, item] of optionsByKey) {
      if (present.has(key)) continue;
      item.remove();
      optionsByKey.delete(key);
    }
    updateScreenTip();
    updateActive();
    positioner.position(
      root,
      'below',
      overlayPortalFor(input).querySelector<HTMLElement>('.fc-arghelper'),
    );
  };

  const refresh = (): void => {
    if (isComposing || !positioner.isAnchorVisible()) {
      close();
      return;
    }
    const start = input.selectionStart ?? input.value.length;
    const end = input.selectionEnd ?? start;
    if (start !== end) {
      dismissed = null;
      close();
      return;
    }
    const text = input.value;
    const caret = start;
    if (dismissed) {
      if (dismissed.value === text && dismissed.caret === caret) {
        close();
        return;
      }
      dismissed = null;
    }
    const next = computeContext(text, caret, deps);
    if (!next) {
      close();
      return;
    }
    const previousKey = highlightKey;
    ctx = next;
    ctxValue = text;
    ctxCaret = caret;
    const preserved = previousKey
      ? next.matches.findIndex(
          (match, index) => candidateKey(next.kinds?.[index] ?? next.kind, match) === previousKey,
        )
      : -1;
    highlight = Math.max(preserved, 0);
    render();
  };

  const move = (delta: 1 | -1): void => {
    if (!ctx || ctx.matches.length === 0) return;
    highlight = (highlight + delta + ctx.matches.length) % ctx.matches.length;
    updateScreenTip();
    updateActive();
  };

  const navigate = (action: 'first' | 'last' | 'pageUp' | 'pageDown'): void => {
    if (!ctx || !listbox || ctx.matches.length === 0) return;
    if (action === 'first') highlight = 0;
    else if (action === 'last') highlight = ctx.matches.length - 1;
    else {
      const rowHeight = listbox.querySelector<HTMLElement>('[role="option"]')?.offsetHeight || 32;
      const pageSize = Math.max(1, Math.floor((listbox.clientHeight || rowHeight * 8) / rowHeight));
      const pageDelta = action === 'pageDown' ? pageSize : -pageSize;
      highlight = Math.max(0, Math.min(ctx.matches.length - 1, highlight + pageDelta));
    }
    updateScreenTip();
    updateActive();
  };

  const dismiss = (): void => {
    dismissed = {
      value: input.value,
      caret: input.selectionStart ?? input.value.length,
    };
    close();
  };

  const setComposing = (composing: boolean): void => {
    if (isComposing === composing) return;
    isComposing = composing;
    if (composing) close();
    else refresh();
  };

  const acceptHighlighted = (): boolean => {
    const start = input.selectionStart ?? input.value.length;
    const end = input.selectionEnd ?? start;
    if (start !== end) {
      close();
      return false;
    }
    if (ctxValue !== input.value || ctxCaret !== start) {
      // Input/caret movement may outlive the popover's last refresh (for
      // example a selection made before Enter). Refresh the list and let the
      // next Enter accept it; never apply old token offsets to new text.
      refresh();
      return false;
    }
    if (!ctx || ctx.matches.length === 0) return false;
    const pick = ctx.matches[highlight];
    if (!pick) return false;
    const before = input.value.slice(0, ctx.tokenStart);
    const after = input.value.slice(ctx.tokenEnd);
    const insert = ctx.kinds?.[highlight] === 'name' ? pick : `${pick}${ctx.insertSuffix}`;
    input.value = before + insert + after;
    const caret = before.length + insert.length;
    input.setSelectionRange(caret, caret);
    input.focus();
    dismissed = { value: input.value, caret };
    close();
    deps.onAfterInsert?.();
    return true;
  };

  const onCompositionStart = (): void => setComposing(true);
  const onCompositionEnd = (): void => setComposing(false);
  input.addEventListener('compositionstart', onCompositionStart);
  input.addEventListener('compositionend', onCompositionEnd);
  positioner.subscribe(() => {
    if (root) refresh();
  });

  return {
    refresh,
    close() {
      // A caller closing the editor ends the current edit session. Clear the
      // Escape/acceptance suppression so a later edit with the same draft can
      // suggest again; internal hides keep that suppression intact.
      dismissed = null;
      close();
    },
    isOpen: () => root != null,
    move,
    navigate,
    dismiss,
    toggle() {
      if (root) dismiss();
      else {
        dismissed = null;
        refresh();
      }
    },
    setComposing,
    acceptHighlighted,
    setLabels(next) {
      labels = { ...labels, ...next };
      render();
    },
    detach() {
      close();
      input.removeEventListener('compositionstart', onCompositionStart);
      input.removeEventListener('compositionend', onCompositionEnd);
      positioner.dispose();
    },
  };
}

/** Build a suggestion context for the caret position. Plain-text edits go
 *  through the column-history source (the "pick from list"). Formula edits
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
  // Suggestions replace the token prefix but leave the suffix in place. If
  // the caret splits a token, accepting would duplicate its remaining text.
  const next = text[caret] ?? '';
  if (next && /[\p{L}\p{N}_.\\]/u.test(next)) return null;
  const struct = suggestStructuredRef(text, caret, deps.getTables?.() ?? []);
  if (struct) return struct;
  const engineNames = deps.getFunctionNames?.();
  const displayNames = deps.getFunctionDisplayNames?.() ?? null;
  const fnNames = displayNames ?? engineNames;
  const fn = suggestFunctions(
    text,
    caret,
    Number.MAX_SAFE_INTEGER,
    fnNames ? { names: fnNames } : undefined,
  );
  const custom = suggestCustomFunctions(text, caret, deps.getCustomFunctions?.() ?? []);
  const names = suggestNames(text, caret, deps.getNames?.() ?? []);
  if (!fn && !custom && !names) return null;
  // Merge runtime and built-in names, then present the complete prefix match
  // set alphabetically. Keep the first spelling on a case-insensitive clash.
  const seen = new Set<string>();
  const merged: string[] = [];
  for (const n of custom?.matches ?? []) {
    const key = n.toLocaleLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    merged.push(n);
  }
  for (const n of fn?.matches ?? []) {
    const key = n.toLocaleLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    merged.push(n);
  }
  const tokenStart = fn?.tokenStart ?? custom?.tokenStart ?? names?.tokenStart ?? caret;
  if (names) {
    // Functions and names share one alphabetical list, as the desktop app does.
    const entries = [
      ...merged.map((label) => ({ label, kind: 'function' as const, detail: '' })),
      ...names.matches.map((n) => ({ label: n.name, kind: 'name' as const, detail: n.detail })),
    ];
    const unique: typeof entries = [];
    const entryNames = new Set<string>();
    for (const entry of entries) {
      const key = entry.label.toLocaleLowerCase();
      if (entryNames.has(key)) continue;
      entryNames.add(key);
      unique.push(entry);
    }
    unique.sort((a, b) => compareCaseless(a.label, b.label));
    return {
      matches: unique.map((e) => e.label),
      tokenStart,
      tokenEnd: caret,
      insertSuffix: '(',
      kind: 'function',
      kinds: unique.map((e) => e.kind),
      details: unique.map((e) => e.detail),
    };
  }
  merged.sort(compareCaseless);
  return {
    matches: merged,
    tokenStart,
    tokenEnd: caret,
    insertSuffix: '(',
    kind: 'function',
  };
}

/** Plain-text values above `beforeRow` in column `col` for the "pick from
 *  list" popover: text-only (formulas, numbers, blanks all skip), deduped,
 *  nearest-first, at most ten. Takes the sheet's populated cells in one pass
 *  rather than probing each row — per-row `cellFormula` lookups are O(n) and
 *  would run on every keystroke. */
export function pickListValues(
  cells: Iterable<{ addr: Addr; value: CellValue; formula: string | null }>,
  col: number,
  beforeRow: number,
): string[] {
  const hits: { row: number; text: string }[] = [];
  for (const e of cells) {
    if (e.addr.col !== col) continue;
    if (e.addr.row >= beforeRow) continue;
    // Formulas don't contribute — the pick-list is verbatim text only.
    if (e.formula !== null) continue;
    if (e.value.kind !== 'text') continue;
    const text = e.value.value;
    if (text.length === 0) continue;
    hits.push({ row: e.addr.row, text });
  }
  // Nearest-first: highest row index wins.
  hits.sort((a, b) => b.row - a.row);
  const out: string[] = [];
  const seen = new Set<string>();
  for (const h of hits) {
    if (seen.has(h.text)) continue;
    seen.add(h.text);
    out.push(h.text);
    if (out.length >= 10) break;
  }
  return out;
}

/** column-history autocomplete: when editing a plain-text cell, match
 *  what the user has typed so far against the values already entered above in
 *  the same column. `values` must already be deduped, blanks excluded, in
 *  nearest-first order (caller's responsibility). The whole input acts as the
 *  partial token — desktop spreadsheets replaces the entire cell text on accept, never just
 *  a fragment. Returns null when the input is empty or no value prefix-matches. */
export function suggestColumnHistory(
  text: string,
  caret: number,
  values: readonly string[],
): SuggestionContext | null {
  // Only suggest when caret is at end and the input is a non-empty token. A
  // shorter caret (mid-edit) means the user is correcting earlier characters,
  // not extending the tail — spreadsheets don't pop the list there either.
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
  if (formulaCaretContext(text, caret) !== 'expression') return null;
  let i = caret - 1;
  while (i >= 0) {
    const ch = text[i] ?? '';
    if (/[\p{L}\p{N}_]/u.test(ch)) i -= 1;
    else break;
  }
  const tokenStart = i + 1;
  const token = text.slice(tokenStart, caret);
  if (token.length < 1 || !/^[\p{L}_]/u.test(token)) return null;
  const upper = token.toLocaleUpperCase();
  const matches = names.filter((n) => n.toLocaleUpperCase().startsWith(upper));
  if (matches.length === 0) return null;
  return { token, tokenStart, matches };
}

const compareCaseless = (a: string, b: string): number => {
  const x = a.toLocaleLowerCase();
  const y = b.toLocaleLowerCase();
  return x < y ? -1 : x > y ? 1 : 0;
};

/** Filter defined/table names against the partial token at `caret`. The match
 *  is a case-insensitive prefix, so Japanese names work the same way. */
function suggestNames(
  text: string,
  caret: number,
  names: readonly AutocompleteName[],
): { tokenStart: number; matches: AutocompleteName[] } | null {
  if (!text.startsWith('=') || names.length === 0) return null;
  if (formulaCaretContext(text, caret) !== 'expression') return null;
  let i = caret - 1;
  while (i >= 0 && /[\p{L}\p{N}_.\\]/u.test(text[i] ?? '')) i -= 1;
  const tokenStart = i + 1;
  const token = text.slice(tokenStart, caret);
  if (token.length < 1 || !/^[\p{L}_\\]/u.test(token)) return null;
  const upper = token.toLocaleUpperCase();
  const seen = new Set<string>();
  const matches = names.filter((n) => {
    const key = n.name.toLocaleUpperCase();
    if (!key.startsWith(upper) || seen.has(key)) return false;
    seen.add(key);
    return true;
  });
  return matches.length === 0 ? null : { tokenStart, matches };
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
  const caretContext = formulaCaretContext(text, caret);
  if (caretContext === 'string' || caretContext === 'sheet') return null;
  const next = text[caret] ?? '';
  if (next && /[\p{L}\p{N}_.\\\]]/u.test(next)) return null;
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
  const seen = new Set<string>();
  const matches = table.columns
    .filter((c) => {
      const key = c.toLocaleLowerCase();
      if (!key.startsWith(lower) || seen.has(key)) return false;
      seen.add(key);
      return true;
    })
    .sort(compareCaseless);
  if (matches.length === 0) return null;
  return {
    matches,
    tokenStart: bracket + 1,
    tokenEnd: caret,
    insertSuffix: ']',
    kind: 'tableColumn',
  };
}
