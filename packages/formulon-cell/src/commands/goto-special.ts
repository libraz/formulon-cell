import type { Addr, Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { addrKey } from '../engine/workbook-handle.js';
import type { SpreadsheetStore } from '../store/store.js';

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
    if (matchesKind(entry.addr, entry.value, entry.formula, kind, store)) {
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
): boolean => {
  switch (kind) {
    case 'non-blanks':
      return value.kind !== 'blank';
    case 'formulas':
      return !!formula;
    case 'constants':
      return !formula && value.kind !== 'blank';
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
