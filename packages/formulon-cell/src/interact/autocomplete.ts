import { FUNCTION_SIGNATURES, suggestFunctions } from '../commands/refs.js';
import { createExcelRibbonSvg } from '../toolbar/excel-ribbon-icons.js';
import type { AutocompleteName, AutocompleteTable } from './autocomplete-names.js';
import {
  type AutocompleteGroup,
  type AutocompleteSuggestionKind,
  caretSplitsToken,
  existingFunctionParenLength,
  partitionByPrefix,
  type SuggestionContext,
  sortCaseless,
  suggestColumnHistory,
  suggestCustomFunctions,
  suggestNames,
  suggestStructuredRef,
} from './autocomplete-suggestions.js';
import { createInteractionButton } from './chip-button.js';
import { createFormulaCaretPositioner } from './formula-helper-position.js';
import { overlayPortalFor } from './overlay-portal.js';
import { viewportSize } from './overlay-position.js';

let autocompleteSeq = 0;

export interface AutocompleteInsertOutcome {
  kind: AutocompleteSuggestionKind;
}

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
  /** Header of the function group in a grouped list. */
  functions: string;
  /** Header of the name group in a grouped list. */
  namedRanges: string;
  /** Header of the table name group in a grouped list. */
  structuredReferences: string;
}

export interface AutocompleteDeps {
  /** The textarea/input being edited. */
  input: HTMLInputElement | HTMLTextAreaElement;
  /** Called immediately before a successful candidate rewrites the input. */
  onBeforeInsert?: () => void;
  /** Called after the popover has rewritten input.value so consumers can
   *  re-sync mirror state (e.g. formula-bar ref highlighting). Receives the
   *  accepted candidate kind; existing zero-argument callbacks remain valid. */
  onAfterInsert?: (outcome?: AutocompleteInsertOutcome) => void;
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
  /** True to list names in their own headed group above the headed functions, as the Mac
   *  desktop does; otherwise names and functions share one alphabetical list. */
  groupNames?: () => boolean;
  labels?: Partial<AutocompleteLabels>;
}

const GROUP_LABEL: Readonly<Record<AutocompleteGroup['kind'], (l: AutocompleteLabels) => string>> =
  {
    name: (l) => l.namedRanges,
    table: (l) => l.structuredReferences,
    function: (l) => l.functions,
  };

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
    functions: 'Functions',
    namedRanges: 'Named Ranges',
    structuredReferences: 'Structured References',
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
  const headersByKind = new Map<AutocompleteGroup['kind'], HTMLDivElement>();
  const groupHeader = (kind: AutocompleteGroup['kind']): HTMLDivElement => {
    let header = headersByKind.get(kind);
    if (!header) {
      header = document.createElement('div');
      header.className = 'fc-autocomplete__group';
      header.setAttribute('role', 'presentation');
      headersByKind.set(kind, header);
    }
    header.textContent = GROUP_LABEL[kind](labels);
    return header;
  };
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
    headersByKind.clear();
    input.setAttribute('aria-expanded', 'false');
    input.removeAttribute('aria-controls');
    input.removeAttribute('aria-activedescendant');
  };

  const candidateKey = (kind: SuggestionContext['kind'], label: string): string =>
    `${kind}:${label.toLocaleLowerCase()}`;

  const positionHelpNearActiveRow = (): void => {
    if (!root || !listbox || !screenTip) return;
    const { width } = viewportSize(input.ownerDocument);
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

  /** `scroll` is for keyboard-driven changes; a pointer-driven one is already under the pointer. */
  const updateActive = (scroll = true): void => {
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
      if (scroll && typeof active.scrollIntoView === 'function')
        active.scrollIntoView({ block: 'nearest' });
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
    const groupStarts = new Map((ctx.groups ?? []).map((group) => [group.start, group.kind]));
    let child = 0;
    for (let i = 0; i < ctx.matches.length; i += 1) {
      const groupKind = groupStarts.get(i);
      if (groupKind) {
        const header = groupHeader(groupKind);
        const currentAtChild = listbox?.children[child] ?? null;
        if (header !== currentAtChild) listbox?.insertBefore(header, currentAtChild);
        child += 1;
      }
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
          updateActive(false);
        });
        item.addEventListener('mousemove', (e) => {
          const option = e.currentTarget as HTMLButtonElement;
          const index = Number(option.dataset.fcIndex ?? -1);
          if (index < 0 || !ctx || !ctx.matches[index] || highlight === index) return;
          highlight = index;
          updateScreenTip();
          updateActive(false);
        });
        item.addEventListener('dblclick', (e) => {
          e.preventDefault();
          const option = e.currentTarget as HTMLButtonElement;
          const index = Number(option.dataset.fcIndex ?? -1);
          if (index < 0 || !ctx || !ctx.matches[index]) return;
          highlight = index;
          updateScreenTip();
          updateActive(false);
          acceptHighlighted();
        });
      }
      item.dataset.fcIndex = String(i);
      const currentAtChild = listbox?.children[child] ?? null;
      if (item !== currentAtChild) listbox?.insertBefore(item, currentAtChild);
      child += 1;
    }
    for (const [key, item] of optionsByKey) {
      if (present.has(key)) continue;
      item.remove();
      optionsByKey.delete(key);
    }
    for (const [kind, header] of headersByKind) {
      if ([...groupStarts.values()].includes(kind)) continue;
      header.remove();
      headersByKind.delete(kind);
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
    const acceptedKind = ctx.kinds?.[highlight] ?? ctx.kind;
    const before = input.value.slice(0, ctx.tokenStart);
    const after = input.value.slice(ctx.tokenEnd);
    const insert =
      ctx.insertions?.[highlight] ??
      (acceptedKind === 'name' ? pick : `${pick}${ctx.insertSuffix}`);
    deps.onBeforeInsert?.();
    input.value = before + insert + after;
    const caret =
      before.length +
      insert.length +
      (acceptedKind === 'function' ? (ctx.existingCallParenLength ?? 0) : 0);
    input.setSelectionRange(caret, caret);
    input.focus();
    dismissed = { value: input.value, caret };
    close();
    deps.onAfterInsert?.({ kind: acceptedKind });
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
    if (caret !== text.length || text.length === 0) return null;
    const values = deps.getColumnValues(addr.sheet, addr.col, addr.row);
    return suggestColumnHistory(text, caret, values);
  }
  if (caretSplitsToken(text, caret)) return null;
  const existingCallParenLength = existingFunctionParenLength(text, caret);
  const struct = existingCallParenLength
    ? null
    : suggestStructuredRef(text, caret, deps.getTables?.() ?? []);
  if (struct) return struct;
  const completion = completeFunctionNames(text, caret, deps);
  const names = existingCallParenLength ? null : suggestNames(text, caret, deps.getNames?.() ?? []);
  if (!completion && !names) return null;
  const functions = completion?.names ?? [];
  const seen = new Set(functions.map((label) => label.toLocaleLowerCase()));
  const tokenStart = completion?.tokenStart ?? names?.tokenStart ?? caret;
  const token = text.slice(tokenStart, caret).toLocaleUpperCase();
  const grouped = deps.groupNames?.() ?? false;
  if (!names) {
    return {
      matches: functions,
      tokenStart,
      tokenEnd: caret,
      insertSuffix: existingCallParenLength ? '' : '(',
      kind: 'function',
      existingCallParenLength: existingCallParenLength ?? undefined,
      groups: grouped && functions.length > 0 ? [{ start: 0, kind: 'function' }] : undefined,
    };
  }
  const nameEntries = sortCaseless(
    names.matches.filter((n) => !seen.has(n.name.toLocaleLowerCase())),
    (n) => n.name,
  ).map((n) => ({ label: n.name, kind: 'name' as const, detail: n.detail, table: n.table }));
  const functionEntries = functions.map((label) => ({
    label,
    kind: 'function' as const,
    detail: '',
  }));
  // The Mac heads defined names, then table names, then functions; Windows mixes them
  // alphabetically.
  type Entry = { label: string; kind: 'name' | 'function'; detail: string };
  const runs: [AutocompleteGroup['kind'], Entry[]][] = [
    ['name', nameEntries.filter((e) => !e.table)],
    ['table', nameEntries.filter((e) => e.table)],
    ['function', functionEntries],
  ];
  const groups: AutocompleteGroup[] = [];
  const entries: Entry[] = [];
  if (grouped) {
    for (const [kind, run] of runs) {
      if (run.length === 0) continue;
      groups.push({ start: entries.length, kind });
      entries.push(...run);
    }
  } else {
    const mixed = sortCaseless([...functionEntries, ...nameEntries], (e) => e.label);
    entries.push(...partitionByPrefix(mixed, (e) => e.label, token));
  }
  return {
    matches: entries.map((e) => e.label),
    tokenStart,
    tokenEnd: caret,
    insertSuffix: '(',
    kind: 'function',
    kinds: entries.map((e) => e.kind),
    details: entries.map((e) => e.detail),
    groups: groups.length > 0 ? groups : undefined,
  };
}

/** The functions a formula can call matching the identifier before `caret`: the custom
 *  functions first, then the built-ins, keeping the first spelling on a case-insensitive clash;
 *  names starting with the token come first, then those only containing it, each run
 *  alphabetical. Null when the caret is not at a function name. */
export function completeFunctionNames(
  text: string,
  caret: number,
  source: Pick<
    AutocompleteDeps,
    'getFunctionNames' | 'getFunctionDisplayNames' | 'getCustomFunctions'
  >,
): { token: string; tokenStart: number; names: string[] } | null {
  const fnNames = source.getFunctionDisplayNames?.() ?? source.getFunctionNames?.() ?? null;
  const fn = suggestFunctions(
    text,
    caret,
    Number.MAX_SAFE_INTEGER,
    fnNames ? { names: fnNames } : undefined,
  );
  const custom = suggestCustomFunctions(text, caret, source.getCustomFunctions?.() ?? []);
  if (!fn && !custom) return null;
  const seen = new Set<string>();
  const merged: string[] = [];
  for (const n of [...(custom?.matches ?? []), ...(fn?.matches ?? [])]) {
    const key = n.toLocaleLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    merged.push(n);
  }
  const tokenStart = fn?.tokenStart ?? custom?.tokenStart ?? caret;
  const token = text.slice(tokenStart, caret);
  const upper = token.toLocaleUpperCase();
  const sorted = sortCaseless(merged, (label) => label);
  return {
    token,
    tokenStart,
    names: partitionByPrefix(sorted, (label) => label, upper),
  };
}
