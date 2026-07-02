import type { Addr } from '../engine/types.js';
import { formatCell } from '../engine/value.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { SpreadsheetStore, State } from '../store/store.js';
import { writeInput } from './coerce-input.js';
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
}

export interface FindMatch {
  addr: Addr;
}

interface CellEntry {
  addr: Addr;
  display: string;
}

function addrFromKey(key: string): Addr | null {
  const parts = key.split(':');
  if (parts.length !== 3) return null;
  const sheet = Number(parts[0]);
  const row = Number(parts[1]);
  const col = Number(parts[2]);
  if (!Number.isFinite(sheet) || !Number.isFinite(row) || !Number.isFinite(col)) return null;
  return { sheet, row, col };
}

function cellsForSearch(state: State, opts: FindOptions): CellEntry[] {
  const sheet = state.data.sheetIndex;
  const out: CellEntry[] = [];
  const lookIn = opts.lookIn ?? 'values';
  if (lookIn === 'comments' || lookIn === 'notes') {
    for (const [key, fmt] of state.format.formats) {
      const addr = addrFromKey(key);
      if (!addr) continue;
      if ((opts.within ?? 'sheet') === 'sheet' && addr.sheet !== sheet) continue;
      if (!fmt.comment) continue;
      out.push({ addr, display: fmt.comment });
    }
    return sortCellEntries(out, opts);
  }

  for (const [key, cell] of state.data.cells) {
    const addr = addrFromKey(key);
    if (!addr) continue;
    if ((opts.within ?? 'sheet') === 'sheet' && addr.sheet !== sheet) continue;
    const display =
      lookIn === 'formulas' ? (cell.formula ?? formatCell(cell.value)) : formatCell(cell.value);
    out.push({ addr, display });
  }
  return sortCellEntries(out, opts);
}

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
): boolean {
  // Don't mutate formula cells — the search runs on the displayed value, so
  // overwriting would silently destroy the formula.
  if (wb.cellFormula(match.addr) !== null) return false;
  if (store && !isCellWritable(store.getState(), match.addr)) {
    warnProtected(match.addr);
    return false;
  }
  // `replacement` is the new raw cell content (callers compute substitution
  // upstream when they need partial replace). Run through writeInput so type
  // is preserved.
  writeInput(wb, match.addr, replacement, store);
  return true;
}

export function replaceAll(
  state: State,
  wb: WorkbookHandle,
  opts: FindOptions,
  replacement: string,
  store?: SpreadsheetStore,
): number {
  if (opts.query === '') return 0;
  const matches = findAll(state, opts);
  let count = 0;
  for (const m of matches) {
    if (wb.cellFormula(m.addr) !== null) continue;
    if (!isCellWritable(state, m.addr)) {
      warnProtected(m.addr);
      continue;
    }
    const cur = formatCell(wb.getValue(m.addr));
    const next = substituteCaseAware(
      cur,
      opts.query,
      replacement,
      opts.caseSensitive ?? false,
      opts.matchWhole ?? false,
    );
    if (next === cur) continue;
    writeInput(wb, m.addr, next, store);
    count += 1;
  }
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
