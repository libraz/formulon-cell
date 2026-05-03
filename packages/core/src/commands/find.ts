import type { Addr } from '../engine/types.js';
import { formatCell } from '../engine/value.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { State } from '../store/store.js';
import { writeInput } from './coerce-input.js';

export interface FindOptions {
  query: string;
  caseSensitive?: boolean;
  /** Entire-cell match. */
  matchWhole?: boolean;
}

export interface FindMatch {
  addr: Addr;
}

interface CellEntry {
  addr: Addr;
  display: string;
}

function activeSheetCells(state: State): CellEntry[] {
  const sheet = state.data.sheetIndex;
  const out: CellEntry[] = [];
  for (const [key, cell] of state.data.cells) {
    const parts = key.split(':');
    if (parts.length !== 3) continue;
    const sh = Number(parts[0]);
    if (sh !== sheet) continue;
    const row = Number(parts[1]);
    const col = Number(parts[2]);
    if (!Number.isFinite(row) || !Number.isFinite(col)) continue;
    out.push({ addr: { sheet: sh, row, col }, display: formatCell(cell.value) });
  }
  // Row-major order so traversal mirrors visual reading order.
  out.sort((a, b) => a.addr.row - b.addr.row || a.addr.col - b.addr.col);
  return out;
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
  for (const entry of activeSheetCells(state)) {
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
  const cells = activeSheetCells(state);
  if (cells.length === 0) return null;

  const matches = cells.filter((e) => isMatch(e.display, opts));
  if (matches.length === 0) return null;

  // Locate the cursor: index of the first match strictly after / before `from`.
  if (!from) {
    const pick = direction === 'next' ? matches[0] : matches[matches.length - 1];
    return pick ? { addr: pick.addr } : null;
  }
  const fromKey = from.row * 1_000_000 + from.col;
  const matchKey = (a: Addr): number => a.row * 1_000_000 + a.col;

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

export function replaceOne(wb: WorkbookHandle, match: FindMatch, replacement: string): void {
  // Don't mutate formula cells — the search runs on the displayed value, so
  // overwriting would silently destroy the formula.
  if (wb.cellFormula(match.addr) !== null) return;
  // `replacement` is the new raw cell content (callers compute substitution
  // upstream when they need partial replace). Run through writeInput so type
  // is preserved.
  writeInput(wb, match.addr, replacement);
}

export function replaceAll(
  state: State,
  wb: WorkbookHandle,
  opts: FindOptions,
  replacement: string,
): number {
  if (opts.query === '') return 0;
  const matches = findAll(state, opts);
  let count = 0;
  for (const m of matches) {
    if (wb.cellFormula(m.addr) !== null) continue;
    const cur = formatCell(wb.getValue(m.addr));
    const next = substituteCaseAware(
      cur,
      opts.query,
      replacement,
      opts.caseSensitive ?? false,
      opts.matchWhole ?? false,
    );
    if (next === cur) continue;
    writeInput(wb, m.addr, next);
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
