import { addrKey } from '../engine/address.js';
import type { Addr, Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { SpreadsheetStore } from '../store/store.js';
import type { SelectionSlice } from '../store/types.js';
import { listComments } from './comment.js';

/** Desktop spreadsheets "Go To Special" categories supported in v1. Mirrors the radio list
 *  on the dialog. Future expansion can add `row-differences`,
 *  `column-differences`, `precedents`, `dependents`. */
export type GoToSpecialKind =
  | 'blanks'
  | 'non-blanks'
  | 'formulas'
  | 'constants'
  | 'numbers'
  | 'text'
  | 'errors'
  | 'data-validation'
  | 'conditional-format';

/** Whether the predicate runs across the active sheet or just the current
 *  selection rectangle. Spreadsheets use the selection scope automatically when the
 *  current selection covers more than one cell. */
export type GoToScope = 'sheet' | 'selection';

/** Excel-style subfilters used by the Formulas and Constants choices in
 *  Go To Special. Omitted options mean "all value kinds". */
export interface GoToSpecialValueFilters {
  numbers?: boolean;
  text?: boolean;
  logical?: boolean;
  errors?: boolean;
}

/** Spreadsheet error sentinels — the strings that can appear in a `text` value via
 *  user input or pasted data. The engine returns `kind: 'error'` for live
 *  formula errors, so we match both. Anything outside this list (e.g. a
 *  literal `#TODO!`) is treated as plain text. */
const ERROR_SENTINELS: ReadonlySet<string> = new Set([
  '#NULL!',
  '#DIV/0!',
  '#VALUE!',
  '#REF!',
  '#NAME?',
  '#NUM!',
  '#N/A',
  '#GETTING_DATA',
  '#SPILL!',
  '#CALC!',
]);

const inRange = (addr: Addr, range: Range): boolean =>
  addr.sheet === range.sheet &&
  addr.row >= range.r0 &&
  addr.row <= range.r1 &&
  addr.col >= range.c0 &&
  addr.col <= range.c1;

/** Walk every populated cell on the sheet (or only those inside the current
 *  selection) and return matches for the given kind. Output is row-major
 *  ordered — the dialog uses the first match as the new active cell.
 *
 *  Predicates rely on three sources:
 *  - `wb.getValue(addr)` — current evaluated kind/value.
 *  - `wb.cellFormula(addr)` — formula text (truthy ≡ formula cell).
 *  - `store` — format slice (data validation entries) and conditional rules.
 *
 *  For `blanks` we additionally walk every cell inside the selection
 *  rectangle so an explicit blank in the middle of an empty range still
 *  surfaces — the spreadsheet behavior. For other kinds we only iterate populated
 *  cells; an empty cell can't satisfy a "formulas" predicate. */
export function findMatchingCells(
  wb: WorkbookHandle,
  store: SpreadsheetStore,
  scope: GoToScope,
  kind: GoToSpecialKind,
  filters?: GoToSpecialValueFilters,
): Addr[] {
  const state = store.getState();
  const sheet = state.data.sheetIndex;
  const selection = state.selection.range;
  const useSelection = scope === 'selection';

  // `blanks` is special — we sweep every cell in the selection rect (or in a
  // bounded sweep around populated cells when scope = sheet) and return the
  // ones that read as blank.
  if (kind === 'blanks') {
    return findBlanks(wb, useSelection ? selection : null, sheet);
  }

  const matches: Addr[] = [];
  for (const entry of wb.cells(sheet)) {
    if (useSelection && !inRange(entry.addr, selection)) continue;
    if (matchesKind(entry.addr, entry.value, entry.formula, kind, store, filters)) {
      matches.push(entry.addr);
    }
  }

  // Sort row-major for predictable bounding-rect anchoring.
  matches.sort((a, b) => a.row - b.row || a.col - b.col);
  return matches;
}

interface CellEntry {
  addr: Addr;
  value: ReturnType<WorkbookHandle['getValue']>;
  formula: string | null;
}

const matchesKind = (
  addr: Addr,
  value: CellEntry['value'],
  formula: string | null,
  kind: GoToSpecialKind,
  store: SpreadsheetStore,
  filters?: GoToSpecialValueFilters,
): boolean => {
  switch (kind) {
    case 'non-blanks':
      return value.kind !== 'blank';
    case 'formulas':
      return !!formula && valueMatchesFilters(value, filters);
    case 'constants':
      return !formula && value.kind !== 'blank' && valueMatchesFilters(value, filters);
    case 'numbers':
      return value.kind === 'number';
    case 'text':
      // Plain text cells, but not spreadsheet error sentinels masquerading as text.
      return value.kind === 'text' && !ERROR_SENTINELS.has(value.value);
    case 'errors':
      if (value.kind === 'error') return true;
      // Some engines surface error sentinels as text. Treat as error too.
      if (value.kind === 'text' && ERROR_SENTINELS.has(value.value)) return true;
      return false;
    case 'data-validation': {
      const fmt = store.getState().format.formats.get(addrKey(addr));
      return !!fmt?.validation;
    }
    case 'conditional-format': {
      const rules = store.getState().conditional.rules;
      for (const rule of rules) if (inRange(addr, rule.range)) return true;
      return false;
    }
    case 'blanks':
      // Handled separately above — populated cells don't match `blanks`.
      return false;
    default:
      return false;
  }
};

const valueMatchesFilters = (
  value: CellEntry['value'],
  filters: GoToSpecialValueFilters | undefined,
): boolean => {
  if (!filters) return true;
  const any =
    filters.numbers === true ||
    filters.text === true ||
    filters.logical === true ||
    filters.errors === true;
  if (!any) return false;
  switch (value.kind) {
    case 'number':
      return filters.numbers === true;
    case 'bool':
      return filters.logical === true;
    case 'text':
      return ERROR_SENTINELS.has(value.value) ? filters.errors === true : filters.text === true;
    case 'error':
      return filters.errors === true;
    default:
      return false;
  }
};

/** Sweep every cell inside the search area and return the ones that read as
 *  blank. Within the selection rectangle (scope === 'selection') we walk the
 *  whole rect; for the sheet scope we walk the bounding rect of populated
 *  cells and yield gaps inside it. Bounding the sheet sweep keeps the cost
 *  proportional to actual content rather than 16M empty cells. */
const findBlanks = (wb: WorkbookHandle, selection: Range | null, sheet: number): Addr[] => {
  const populated = new Set<string>();
  let minRow = Number.POSITIVE_INFINITY;
  let maxRow = Number.NEGATIVE_INFINITY;
  let minCol = Number.POSITIVE_INFINITY;
  let maxCol = Number.NEGATIVE_INFINITY;
  for (const entry of wb.cells(sheet)) {
    populated.add(addrKey(entry.addr));
    if (entry.addr.row < minRow) minRow = entry.addr.row;
    if (entry.addr.row > maxRow) maxRow = entry.addr.row;
    if (entry.addr.col < minCol) minCol = entry.addr.col;
    if (entry.addr.col > maxCol) maxCol = entry.addr.col;
  }

  const r0 = selection ? selection.r0 : Number.isFinite(minRow) ? minRow : 0;
  const r1 = selection ? selection.r1 : Number.isFinite(maxRow) ? maxRow : -1;
  const c0 = selection ? selection.c0 : Number.isFinite(minCol) ? minCol : 0;
  const c1 = selection ? selection.c1 : Number.isFinite(maxCol) ? maxCol : -1;
  if (r1 < r0 || c1 < c0) return [];

  const matches: Addr[] = [];
  for (let r = r0; r <= r1; r += 1) {
    for (let c = c0; c <= c1; c += 1) {
      const addr: Addr = { sheet, row: r, col: c };
      if (populated.has(addrKey(addr))) continue;
      const v = wb.getValue(addr);
      if (v.kind === 'blank') matches.push(addr);
    }
  }
  return matches;
};

/** Inclusive bounding rectangle of a non-empty addr list. Throws when the
 *  list is empty — callers should check `matches.length > 0` first. */
export const boundingRange = (matches: readonly Addr[]): Range => {
  if (matches.length === 0) {
    throw new Error('boundingRange: empty match list');
  }
  const first = matches[0] as Addr;
  let r0 = first.row;
  let r1 = first.row;
  let c0 = first.col;
  let c1 = first.col;
  for (const a of matches) {
    if (a.row < r0) r0 = a.row;
    if (a.row > r1) r1 = a.row;
    if (a.col < c0) c0 = a.col;
    if (a.col > c1) c1 = a.col;
  }
  return { sheet: first.sheet, r0, c0, r1, c1 };
};

/** Build an exact multi-range selection for a list of row-major matches.
 *  The first matched cell becomes the primary range and the rest are stored as
 *  single-cell `extraRanges`, avoiding the false positives introduced by an
 *  enclosing rectangle. */
export const selectionFromMatches = (matches: readonly Addr[]): SelectionSlice => {
  if (matches.length === 0) {
    throw new Error('selectionFromMatches: empty match list');
  }
  const first = matches[0] as Addr;
  return {
    active: first,
    anchor: first,
    range: { sheet: first.sheet, r0: first.row, c0: first.col, r1: first.row, c1: first.col },
    extraRanges: matches.slice(1).map((addr) => ({
      sheet: addr.sheet,
      r0: addr.row,
      c0: addr.col,
      r1: addr.row,
      c1: addr.col,
    })),
  };
};

export type RibbonFindAction =
  | 'find'
  | 'replace'
  | 'go-to'
  | 'go-to-special'
  | 'formulas'
  | 'constants'
  | 'numbers'
  | 'text'
  | 'errors'
  | 'conditional-format'
  | 'data-validation'
  | 'comments';

export interface RibbonFindActionReportItem {
  severity: 'info' | 'warning';
  label: string;
  detail: string;
}

export interface RibbonFindActionReport {
  title: string;
  items: RibbonFindActionReportItem[];
}

export type RibbonFindActionResult =
  | { kind: 'open-find'; mode: 'find' | 'replace' }
  | { kind: 'open-go-to' }
  | { kind: 'open-go-to-special' }
  | { kind: 'selected' }
  | { kind: 'report'; report: RibbonFindActionReport }
  | { kind: 'noop' };

export interface ExecuteRibbonFindActionDeps {
  store: SpreadsheetStore;
  workbook: WorkbookHandle;
  action: RibbonFindAction;
  strings: {
    findSelect: string;
    findNoMatches: string;
    commentNone: string;
  };
}

const SPECIAL_FIND_KINDS = new Set<RibbonFindAction>([
  'formulas',
  'constants',
  'numbers',
  'text',
  'errors',
  'conditional-format',
  'data-validation',
]);

/** Shared "Find & Select" ribbon split-button. Resolves the action into a
 *  host-routable verdict — dialog opener, selection update applied in place,
 *  or a report dialog payload to surface "no matches"/"no comments".
 *
 *  This keeps React/Vue wrappers free of the "which dialog → which selector →
 *  what message" branching; each host just plugs the result into its dialog
 *  setter and `openFindReplace` / `openGoTo` / `openGoToSpecial` methods. */
export const executeRibbonFindAction = (
  deps: ExecuteRibbonFindActionDeps,
): RibbonFindActionResult => {
  const { store, workbook, action, strings } = deps;
  if (action === 'find') return { kind: 'open-find', mode: 'find' };
  if (action === 'replace') return { kind: 'open-find', mode: 'replace' };
  if (action === 'go-to') return { kind: 'open-go-to' };
  if (action === 'go-to-special') return { kind: 'open-go-to-special' };
  if (SPECIAL_FIND_KINDS.has(action)) {
    const matches = findMatchingCells(
      workbook,
      store,
      'sheet',
      action as Exclude<
        RibbonFindAction,
        'find' | 'replace' | 'go-to' | 'go-to-special' | 'comments'
      >,
    );
    if (!matches[0]) {
      return {
        kind: 'report',
        report: {
          title: strings.findSelect,
          items: [{ severity: 'info', label: strings.findNoMatches, detail: '' }],
        },
      };
    }
    store.setState((state) => ({ ...state, selection: selectionFromMatches(matches) }));
    return { kind: 'selected' };
  }
  if (action === 'comments') {
    const comments = listComments(store.getState());
    const first = comments[0]?.addr;
    if (!first) {
      return {
        kind: 'report',
        report: {
          title: strings.findSelect,
          items: [{ severity: 'info', label: strings.commentNone, detail: '' }],
        },
      };
    }
    const selection = selectionFromMatches(comments.map((entry) => entry.addr));
    store.setState((state) => ({ ...state, selection }));
    return { kind: 'selected' };
  }
  return { kind: 'noop' };
};
