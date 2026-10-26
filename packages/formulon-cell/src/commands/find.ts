import { addrKey, parseAddrKey } from '../engine/address.js';
import type { Addr, CellValue } from '../engine/types.js';
import { formatCell } from '../engine/value.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { SpreadsheetStore, State } from '../store/store.js';
import { inputLocaleForWorkbook, writeInput } from './coerce-input.js';
import { blockDataTableCells } from './data-table-guard.js';
import { isCellWritable, warnProtected } from './protection.js';

export interface FindOptions {
  query: string;
  caseSensitive?: boolean;
  /** Entire-cell match. */
  matchWhole?: boolean;
  /** Excel's "Within" option. Defaults to the active sheet. */
  within?: 'sheet' | 'workbook';
  /** Excel's "Search" option. Defaults to row-major traversal. */
  searchBy?: 'rows' | 'columns';
  /** Excel's "Look in" option. Defaults to displayed values. */
  lookIn?: 'formulas' | 'values' | 'comments' | 'notes';
  /** Workbook used for locale-aware display text and threaded comments.
   *  Optional so the command remains usable against a store-only snapshot. */
  workbook?: WorkbookHandle;
  /** Explicit display/input locale used when the workbook cannot supply one. */
  locale?: string;
}

export interface FindMatch {
  addr: Addr;
}

interface CellEntry {
  addr: Addr;
  display: string;
}

interface ThreadEntry {
  id: string;
  parentId: string;
  row: number;
  col: number;
  text: string;
}

interface SearchCell {
  addr: Addr;
  value: CellValue;
  formula: string | null;
}

const localeForSearch = (opts: FindOptions): string =>
  opts.locale ?? (opts.workbook ? inputLocaleForWorkbook(opts.workbook) : undefined) ?? 'en-US';

const valueDisplay = (value: Parameters<typeof formatCell>[0], opts: FindOptions): string =>
  formatCell(value, localeForSearch(opts));

const workbookDisplay = (wb: WorkbookHandle | undefined, addr: Addr, fallback: string): string => {
  if (!wb || typeof wb.getDisplayText !== 'function') return fallback;
  try {
    return wb.getDisplayText(addr.sheet, addr.row, addr.col) ?? fallback;
  } catch {
    return fallback;
  }
};

const formulaDisplay = (formula: string, opts: FindOptions): string => {
  if (!opts.workbook || typeof opts.workbook.getFormulaForDisplay !== 'function') return formula;
  try {
    return opts.workbook.getFormulaForDisplay(formula);
  } catch {
    return formula;
  }
};

const entriesForThreadedSheet = (
  state: State,
  sheet: number,
  wb: WorkbookHandle | undefined,
): readonly ThreadEntry[] => {
  if (wb?.capabilities?.threadedComments) {
    try {
      return wb.getThreadedComments(sheet);
    } catch {
      // Use the hydrated store snapshot when an engine read is unavailable.
    }
  }
  return state.threadedComments.bySheet.get(sheet) ?? [];
};

const threadedCommentCells = (state: State, opts: FindOptions): CellEntry[] => {
  const wb = opts.workbook;
  let sheets: number[];
  if ((opts.within ?? 'sheet') === 'sheet') {
    sheets = [state.data.sheetIndex];
  } else if (wb) {
    try {
      sheets = Array.from({ length: wb.sheetCount }, (_, sheet) => sheet);
    } catch {
      sheets = [...state.threadedComments.bySheet.keys()];
    }
  } else {
    sheets = [...state.threadedComments.bySheet.keys()];
  }
  const entries = new Map<string, CellEntry>();
  for (const sheet of sheets) {
    const comments = entriesForThreadedSheet(state, sheet, wb);
    const byId = new Map(comments.map((entry) => [entry.id, entry]));
    const orderedComments = [...comments].sort(
      (left, right) => Number(left.parentId !== '') - Number(right.parentId !== ''),
    );
    for (const comment of orderedComments) {
      const opening = comment.parentId ? (byId.get(comment.parentId) ?? comment) : comment;
      const addr = { sheet, row: opening.row, col: opening.col };
      const key = addrKey(addr);
      const existing = entries.get(key);
      entries.set(key, {
        addr,
        display: existing ? `${existing.display}\n${comment.text}` : comment.text,
      });
    }
  }
  return sortCellEntries([...entries.values()], opts);
};

const searchSheets = (state: State, opts: FindOptions): number[] => {
  if ((opts.within ?? 'sheet') === 'sheet') return [state.data.sheetIndex];
  if (opts.workbook) {
    try {
      const count = opts.workbook.sheetCount;
      if (Number.isInteger(count) && count > 0) {
        return Array.from({ length: count }, (_, sheet) => sheet);
      }
    } catch {
      // Fall through to the sheets represented by the store snapshot.
    }
  }
  const sheets = new Set<number>();
  for (const key of state.data.cells.keys()) {
    const addr = parseAddrKey(key);
    if (addr) sheets.add(addr.sheet);
  }
  for (const key of state.format.formats.keys()) {
    const addr = parseAddrKey(key);
    if (addr) sheets.add(addr.sheet);
  }
  for (const sheet of state.threadedComments.bySheet.keys()) sheets.add(sheet);
  sheets.add(state.data.sheetIndex);
  return [...sheets].sort((left, right) => left - right);
};

const cellsForSearch = (state: State, opts: FindOptions): CellEntry[] => {
  const lookIn = opts.lookIn ?? 'values';
  if (lookIn === 'comments') return threadedCommentCells(state, opts);
  if (lookIn === 'notes') return notesForSearch(state, opts);
  const selected = new Set(searchSheets(state, opts));
  const cells = new Map<string, SearchCell>();
  if (opts.workbook) {
    for (const sheet of selected) {
      try {
        for (const entry of opts.workbook.cells(sheet)) {
          cells.set(addrKey(entry.addr), entry);
        }
      } catch {
        // A store-only snapshot remains a valid fallback for older adapters.
      }
    }
  }
  // The store is authoritative for the active sheet's pending edits, and may
  // also contain explicitly hydrated cells from other sheets in low-level
  // hosts. Overlay it after the workbook enumeration to retain those edits.
  for (const [key, cell] of state.data.cells) {
    const addr = parseAddrKey(key);
    if (!addr || !selected.has(addr.sheet)) continue;
    cells.set(key, { addr, value: cell.value, formula: cell.formula });
  }
  return sortCellEntries(
    [...cells.values()].map((cell) => ({
      addr: cell.addr,
      display: cellDisplayForSearch(cell, cell.addr, opts),
    })),
    opts,
  );
};

const notesForSearch = (state: State, opts: FindOptions): CellEntry[] => {
  const selected = new Set(searchSheets(state, opts));
  const out = new Map<string, CellEntry>();
  const stateHasComment = new Set<string>();
  for (const [key, fmt] of state.format.formats) {
    const addr = parseAddrKey(key);
    if (!addr) continue;
    if (!selected.has(addr.sheet)) continue;
    if (Object.hasOwn(fmt, 'comment')) stateHasComment.add(key);
    if (!fmt.comment) continue;
    out.set(key, { addr, display: fmt.comment });
  }
  if (opts.workbook?.capabilities?.comments) {
    for (const sheet of selected) {
      try {
        const native =
          opts.workbook.capabilities.commentsEnumerable &&
          typeof opts.workbook.getComments === 'function'
            ? opts.workbook.getComments(sheet)
            : [...opts.workbook.cells(sheet)].flatMap((entry) => {
                const comment = opts.workbook?.getComment(sheet, entry.addr.row, entry.addr.col);
                return comment ? [{ ...entry.addr, ...comment }] : [];
              });
        for (const comment of native) {
          const addr = { sheet, row: comment.row, col: comment.col };
          const key = addrKey(addr);
          if (stateHasComment.has(key)) continue;
          if (comment.text) out.set(key, { addr, display: comment.text });
        }
      } catch {
        // The hydrated store remains the compatibility path for old engines.
      }
    }
  }
  return sortCellEntries([...out.values()], opts);
};

const cellDisplayForSearch = (
  cell: { value: Parameters<typeof formatCell>[0]; formula: string | null } | undefined,
  addr: Addr,
  opts: FindOptions,
): string => {
  if (!cell) return '';
  if (opts.lookIn === 'formulas') {
    let formula = cell.formula;
    if (opts.workbook && typeof opts.workbook.cellFormula === 'function') {
      try {
        formula = opts.workbook.cellFormula(addr) ?? formula;
      } catch {
        // Keep the hydrated store formula when the engine read is unavailable.
      }
    }
    return formula ? formulaDisplay(formula, opts) : '';
  }
  return workbookDisplay(opts.workbook, addr, valueDisplay(cell.value, opts));
};

function sortCellEntries(entries: CellEntry[], opts: FindOptions): CellEntry[] {
  const byRows = opts.searchBy !== 'columns';
  entries.sort((a, b) => {
    if (a.addr.sheet !== b.addr.sheet) return a.addr.sheet - b.addr.sheet;
    return byRows
      ? a.addr.row - b.addr.row || a.addr.col - b.addr.col
      : a.addr.col - b.addr.col || a.addr.row - b.addr.row;
  });
  return entries;
}

function isMatch(display: string, opts: FindOptions): boolean {
  if (opts.query === '') return false;
  const haystack = opts.caseSensitive ? display : display.toLowerCase();
  const needle = opts.caseSensitive ? opts.query : opts.query.toLowerCase();
  if (opts.matchWhole) return haystack === needle;
  return haystack.includes(needle);
}

export function findAll(state: State, opts: FindOptions): FindMatch[] {
  if (opts.query === '') return [];
  const out: FindMatch[] = [];
  for (const entry of cellsForSearch(state, opts)) {
    if (isMatch(entry.display, opts)) out.push({ addr: entry.addr });
  }
  return out;
}

export function findNext(
  state: State,
  opts: FindOptions,
  from: Addr | null,
  direction: 'next' | 'prev',
): FindMatch | null {
  if (opts.query === '') return null;
  const cells = cellsForSearch(state, opts);
  if (cells.length === 0) return null;

  const matches = cells.filter((e) => isMatch(e.display, opts));
  if (matches.length === 0) return null;

  // Locate the cursor: index of the first match strictly after / before `from`.
  if (!from) {
    const pick = direction === 'next' ? matches[0] : matches[matches.length - 1];
    return pick ? { addr: pick.addr } : null;
  }
  const byRows = opts.searchBy !== 'columns';
  const fromKey =
    from.sheet * 1_000_000_000_000 +
    (byRows ? from.row * 1_000_000 + from.col : from.col * 1_000_000 + from.row);
  const matchKey = (a: Addr): number =>
    a.sheet * 1_000_000_000_000 + (byRows ? a.row * 1_000_000 + a.col : a.col * 1_000_000 + a.row);

  if (direction === 'next') {
    const ahead = matches.find((m) => matchKey(m.addr) > fromKey);
    const pick = ahead ?? matches[0];
    return pick ? { addr: pick.addr } : null;
  }
  // prev
  let behind: CellEntry | null = null;
  for (const m of matches) {
    if (matchKey(m.addr) < fromKey) behind = m;
    else break;
  }
  const pick = behind ?? matches[matches.length - 1];
  return pick ? { addr: pick.addr } : null;
}

/** Resolve the same displayed text used by findAll for a result row or a
 * replacement. Keeping this projection in the command module prevents the
 * dialog from falling back to the browser's English formatter. */
export function displayTextForFindMatch(state: State, addr: Addr, opts: FindOptions): string {
  const lookIn = opts.lookIn ?? 'values';
  if (lookIn === 'comments' || lookIn === 'notes') {
    const entries =
      lookIn === 'comments' ? threadedCommentCells(state, opts) : notesForSearch(state, opts);
    return entries.find((entry) => addrKey(entry.addr) === addrKey(addr))?.display ?? '';
  }
  const cell = state.data.cells.get(addrKey(addr));
  if (cell) return cellDisplayForSearch(cell, addr, opts);

  // Workbook-wide search can return a cell from a sheet that the store has
  // not hydrated.  Result rows and restricted Replace All still need the
  // exact projection used by the command search, so read the engine directly
  // when there is no pending store overlay for this address.
  if (!opts.workbook) return '';
  if (lookIn === 'formulas') {
    try {
      const formula = opts.workbook.cellFormula(addr);
      return formula ? formulaDisplay(formula, opts) : '';
    } catch {
      return '';
    }
  }
  try {
    const value = opts.workbook.getValue(addr);
    return workbookDisplay(opts.workbook, addr, valueDisplay(value, opts));
  } catch {
    return '';
  }
}

function substituteCaseAware(
  source: string,
  query: string,
  replacement: string,
  caseSensitive: boolean,
  matchWhole: boolean,
): string {
  if (matchWhole) {
    const eq = caseSensitive ? source === query : source.toLowerCase() === query.toLowerCase();
    return eq ? replacement : source;
  }
  if (caseSensitive) return source.split(query).join(replacement);
  // Case-insensitive scan that preserves untouched segments verbatim.
  const lcSource = source.toLowerCase();
  const lcQuery = query.toLowerCase();
  if (!lcQuery) return source;
  let out = '';
  let i = 0;
  while (i <= source.length - lcQuery.length) {
    if (lcSource.startsWith(lcQuery, i)) {
      out += replacement;
      i += lcQuery.length;
    } else {
      out += source[i];
      i += 1;
    }
  }
  out += source.slice(i);
  return out;
}

export function replaceOne(
  wb: WorkbookHandle,
  match: FindMatch,
  replacement: string,
  store?: SpreadsheetStore,
  locale?: string,
  lookIn: FindOptions['lookIn'] = 'values',
): boolean {
  if (lookIn === 'comments' || lookIn === 'notes') return false;
  const formula = wb.cellFormula(match.addr);
  // Formula replacement is an explicit mode of Find/Replace. Values mode
  // still refuses formulas so a displayed result can never destroy its
  // source expression.
  if (lookIn === 'formulas' ? formula === null : formula !== null) return false;
  if (store && !isCellWritable(store.getState(), match.addr)) {
    warnProtected(match.addr);
    return false;
  }
  if (blockDataTableCells(wb, [match.addr])) return false;
  // `writeInput` canonicalizes a localized formula before it reaches the
  // engine, and applies the selected input locale to displayed values.
  writeInput(wb, match.addr, replacement, store, locale);
  return true;
}

export function replaceAll(
  state: State,
  wb: WorkbookHandle,
  opts: FindOptions,
  replacement: string,
  store?: SpreadsheetStore,
): number {
  if (opts.query === '' || (opts.lookIn ?? 'values') === 'comments' || opts.lookIn === 'notes') {
    return 0;
  }
  const matches = findAll(state, opts).filter((m) => {
    const formula = wb.cellFormula(m.addr);
    return opts.lookIn === 'formulas' ? formula !== null : formula === null;
  });
  // Refuse the whole replacement up front rather than stopping partway.
  if (
    blockDataTableCells(
      wb,
      matches.map((m) => m.addr),
    )
  )
    return 0;
  let count = 0;
  // The mode gate below keeps formula source replacement and displayed-value
  // replacement separate, so a batched recalc cannot turn one into the other.
  wb.withBatchedRecalc(() => {
    for (const m of matches) {
      const formula = wb.cellFormula(m.addr);
      if (!isCellWritable(state, m.addr)) {
        warnProtected(m.addr);
        continue;
      }
      const cur =
        opts.lookIn === 'formulas'
          ? formulaDisplay(formula ?? '', opts)
          : workbookDisplay(wb, m.addr, valueDisplay(wb.getValue(m.addr), opts));
      const next = substituteCaseAware(
        cur,
        opts.query,
        replacement,
        opts.caseSensitive ?? false,
        opts.matchWhole ?? false,
      );
      if (next === cur) continue;
      writeInput(wb, m.addr, next, store, opts.locale);
      count += 1;
    }
  });
  return count;
}

/** Compute the substituted display string for a single match. Exposed for the
 *  overlay so it can apply per-match replacements with consistent semantics. */
export function applySubstitution(source: string, opts: FindOptions, replacement: string): string {
  return substituteCaseAware(
    source,
    opts.query,
    replacement,
    opts.caseSensitive ?? false,
    opts.matchWhole ?? false,
  );
}
