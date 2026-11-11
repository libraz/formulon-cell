import { addrKey, parseAddrKey } from '../engine/address.js';
import type { Addr, CellValue, Range } from '../engine/types.js';
import { formatCell } from '../engine/value.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { SpreadsheetStore, State } from '../store/store.js';
import { escapeRegExp } from '../text-escape.js';
import { inputLocaleForWorkbook, writeInput } from './coerce-input.js';
import { blockDataTableCells } from './data-table-guard.js';
import { isCellWritable, warnProtected } from './protection.js';

export interface FindOptions {
  query: string;
  caseSensitive?: boolean;
  /** Entire-cell match. */
  matchWhole?: boolean;
  /** Excel's "Match byte": `true` tells a half-width voiced kana pair (ｶﾞ)
   *  from its composed form (ガ). Half-width and full-width letters, digits,
   *  symbols and katakana, and hiragana against katakana, always compare
   *  equal, as they do in the desktop app. */
  matchWidth?: boolean;
  /** Excel's "Within" option. Defaults to the active sheet. */
  within?: 'sheet' | 'workbook';
  /** Excel's "Search" option. Defaults to row-major traversal. */
  searchBy?: 'rows' | 'columns';
  /** Excel's "Look in" option. Defaults to displayed values. */
  lookIn?: 'formulas' | 'values' | 'comments' | 'notes';
  /** List matches in the desktop app's Find All order: sheets from the active
   *  one, and per sheet the cells after the active cell first, wrapping to it
   *  last. Off keeps plain sheet/row order, which stepping through matches
   *  relies on. */
  listFromActive?: boolean;
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
  parentId: string | null;
  cell: { row: number; col: number };
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
      (left, right) => Number(left.parentId !== null) - Number(right.parentId !== null),
    );
    for (const comment of orderedComments) {
      const opening = comment.parentId ? (byId.get(comment.parentId) ?? comment) : comment;
      const addr = { sheet, row: opening.cell.row, col: opening.cell.col };
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
    // A constant is its own content in the formula bar, as in Excel's "Formulas".
    return formula ? formulaDisplay(formula, opts) : valueDisplay(cell.value, opts);
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

const HALFWIDTH_KATAKANA =
  '。「」、・ヲァィゥェォャュョッーアイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワン゛゜';

const VOICED_MARK = '゛';
const SEMI_VOICED_MARK = '゜';
/** Katakana that take the voiced mark (one code point up) and the semi-voiced one (two up). */
const VOICEABLE = 'カキクケコサシスセソタチツテトハヒフヘホ';
const SEMI_VOICEABLE = 'ハヒフヘホ';

/** The composed kana a base and a following mark spell, if any. */
const composeKana = (base: string, mark: string | undefined): string | null => {
  if (mark === VOICED_MARK) {
    if (base === 'ウ') return 'ヴ';
    if (VOICEABLE.includes(base)) return String.fromCharCode(base.charCodeAt(0) + 1);
  }
  if (mark === SEMI_VOICED_MARK && SEMI_VOICEABLE.includes(base)) {
    return String.fromCharCode(base.charCodeAt(0) + 2);
  }
  return null;
};

/** A composed voiced kana split into its base and mark, if it is one. */
const decomposeKana = (ch: string): string | null => {
  if (ch === 'ヴ') return `ウ${VOICED_MARK}`;
  const code = ch.charCodeAt(0);
  const voiced = String.fromCharCode(code - 1);
  if (VOICEABLE.includes(voiced) && composeKana(voiced, VOICED_MARK) === ch) {
    return voiced + VOICED_MARK;
  }
  const semi = String.fromCharCode(code - 2);
  if (SEMI_VOICEABLE.includes(semi) && composeKana(semi, SEMI_VOICED_MARK) === ch) {
    return semi + SEMI_VOICED_MARK;
  }
  return null;
};

/** One code unit as the desktop app compares it: full-width ASCII as
 *  half-width, half-width katakana as full-width, hiragana as katakana, and
 *  lower case unless the case must match. Always one code unit long, so a
 *  folded string indexes like its source. */
const foldChar = (ch: string, caseSensitive: boolean): string => {
  const code = ch.charCodeAt(0);
  let out = ch;
  if (code >= 0xff01 && code <= 0xff5e) out = String.fromCharCode(code - 0xfee0);
  else if (code === 0x3000) out = ' ';
  else if (code >= 0xff61 && code <= 0xff9f) out = HALFWIDTH_KATAKANA[code - 0xff61] ?? ch;
  else if (code >= 0x3041 && code <= 0x3096) out = String.fromCharCode(code + 0x60);
  if (!caseSensitive) {
    const lower = out.toLowerCase();
    if (lower.length === 1) out = lower;
  }
  return out;
};

const foldText = (text: string, caseSensitive: boolean): string => {
  let out = '';
  for (let i = 0; i < text.length; i += 1) out += foldChar(text[i] ?? '', caseSensitive);
  return out;
};

/** A voiced kana that matches both its composed and its base-plus-mark
 *  spelling, or only the query's own spelling under Match byte. */
const kanaAlternatives = (composed: string, split: string, own: string, matchByte: boolean) =>
  matchByte ? escapeRegExp(own) : `(?:${escapeRegExp(composed)}|${escapeRegExp(split)})`;

/** The query as a pattern over folded text: `?` is one character, `*` the
 *  shortest run that lets the rest match, and `~` makes the next character
 *  literal (a tilde before any other character is dropped). */
function queryPattern(opts: FindOptions): RegExp {
  const caseSensitive = opts.caseSensitive ?? false;
  const matchByte = opts.matchWidth === true;
  const query = opts.query;
  let source = '';
  for (let i = 0; i < query.length; i += 1) {
    let ch = query[i] ?? '';
    if (ch === '*' || ch === '?') {
      source += ch === '*' ? '[\\s\\S]*?' : '[\\s\\S]';
      continue;
    }
    if (ch === '~' && i + 1 < query.length) {
      i += 1;
      ch = query[i] ?? '';
    }
    const folded = foldChar(ch, caseSensitive);
    const next = query[i + 1];
    const composed =
      next !== undefined && next !== '~'
        ? composeKana(folded, foldChar(next, caseSensitive))
        : null;
    if (composed) {
      const split = folded + foldChar(next ?? '', caseSensitive);
      source += kanaAlternatives(composed, split, split, matchByte);
      i += 1;
      continue;
    }
    const split = decomposeKana(folded);
    source += split ? kanaAlternatives(folded, split, folded, matchByte) : escapeRegExp(folded);
  }
  return new RegExp(opts.matchWhole ? `^(?:${source})$` : source, opts.matchWhole ? '' : 'g');
}

/** A predicate for cell texts the query finds. */
export function matcherFor(opts: FindOptions): (display: string) => boolean {
  if (opts.query === '') return () => false;
  const pattern = queryPattern(opts);
  const caseSensitive = opts.caseSensitive ?? false;
  return (display) => {
    pattern.lastIndex = 0;
    return pattern.test(foldText(display, caseSensitive));
  };
}

const inRanges = (ranges: readonly Range[], addr: Addr): boolean =>
  ranges.some(
    (r) =>
      r.sheet === addr.sheet &&
      addr.row >= r.r0 &&
      addr.row <= r.r1 &&
      addr.col >= r.c0 &&
      addr.col <= r.c1,
  );

/** Matches in the order the desktop app lists them: sheets from the active
 *  one, and within each sheet the cells after the active cell first, then
 *  the cells up to and including it. */
const fromActiveCell = (state: State, opts: FindOptions, matches: CellEntry[]): FindMatch[] => {
  const byRows = opts.searchBy !== 'columns';
  const active = state.selection.active;
  const key = (a: Addr): number[] => (byRows ? [a.row, a.col] : [a.col, a.row]);
  const after = (a: Addr): boolean => {
    const [k0 = 0, k1 = 0] = key(a);
    const [a0 = 0, a1 = 0] = key(active);
    return k0 > a0 || (k0 === a0 && k1 > a1);
  };
  const sheets = [...new Set(matches.map((m) => m.addr.sheet))];
  const first = sheets.findIndex((sheet) => sheet >= state.data.sheetIndex);
  const order = first < 0 ? sheets : [...sheets.slice(first), ...sheets.slice(0, first)];
  const out: FindMatch[] = [];
  for (const sheet of order) {
    const own = matches.filter((m) => m.addr.sheet === sheet);
    for (const m of own) if (after(m.addr)) out.push({ addr: m.addr });
    for (const m of own) if (!after(m.addr)) out.push({ addr: m.addr });
  }
  return out;
};

/** Every match, in sheet order unless `listFromActive`. A multi-cell selection narrows a sheet search to itself, as
 *  in the desktop app. */
export function findAll(state: State, opts: FindOptions): FindMatch[] {
  if (opts.query === '') return [];
  const isMatch = matcherFor(opts);
  const { range, extraRanges } = state.selection;
  const scoped =
    (opts.within ?? 'sheet') === 'sheet' && (range.r0 !== range.r1 || range.c0 !== range.c1)
      ? [range, ...(extraRanges ?? [])]
      : null;
  const hits = cellsForSearch(state, opts).filter(
    (entry) => isMatch(entry.display) && (!scoped || inRanges(scoped, entry.addr)),
  );
  return opts.listFromActive
    ? fromActiveCell(state, opts, hits)
    : hits.map((m) => ({ addr: m.addr }));
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

  const isMatch = matcherFor(opts);
  const matches = cells.filter((e) => isMatch(e.display));
  if (matches.length === 0) return null;

  // Locate the cursor: index of the first match strictly after / before `from`.
  if (!from) {
    const pick = direction === 'next' ? matches[0] : matches[matches.length - 1];
    return pick ? { addr: pick.addr } : null;
  }
  const byRows = opts.searchBy !== 'columns';
  const compare = (a: Addr, b: Addr): number =>
    a.sheet - b.sheet || (byRows ? a.row - b.row || a.col - b.col : a.col - b.col || a.row - b.row);

  if (direction === 'next') {
    const ahead = matches.find((m) => compare(m.addr, from) > 0);
    const pick = ahead ?? matches[0];
    return pick ? { addr: pick.addr } : null;
  }
  // prev
  let behind: CellEntry | null = null;
  for (const m of matches) {
    if (compare(m.addr, from) < 0) behind = m;
    else break;
  }
  const pick = behind ?? matches[matches.length - 1];
  return pick ? { addr: pick.addr } : null;
}

/** Resolve the same displayed text used by findAll for a result row or a
 * replacement. Keeping this projection in the command module prevents the
 * dialog from falling back to the browser's English formatter. */
export function displayTextForFindMatch(state: State, addr: Addr, opts: FindOptions): string {
  return findDisplayResolver(state, opts)(addr);
}

/** `displayTextForFindMatch` for many addresses: the comment/note lookup is
 *  built once, not per address. */
export function findDisplayResolver(state: State, opts: FindOptions): (addr: Addr) => string {
  const lookIn = opts.lookIn ?? 'values';
  if (lookIn === 'comments' || lookIn === 'notes') {
    const entries =
      lookIn === 'comments' ? threadedCommentCells(state, opts) : notesForSearch(state, opts);
    const byKey = new Map(entries.map((entry) => [addrKey(entry.addr), entry.display]));
    return (addr) => byKey.get(addrKey(addr)) ?? '';
  }
  return (addr) => cellDisplayText(state, addr, opts, lookIn);
}

function cellDisplayText(
  state: State,
  addr: Addr,
  opts: FindOptions,
  lookIn: NonNullable<FindOptions['lookIn']>,
): string {
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

/** `source` with every text the query finds replaced, or the whole text on
 *  an entire-cell match. Text outside the matches stays verbatim. */
export function applySubstitution(source: string, opts: FindOptions, replacement: string): string {
  if (opts.query === '') return source;
  const pattern = queryPattern(opts);
  const folded = foldText(source, opts.caseSensitive ?? false);
  if (opts.matchWhole) return pattern.test(folded) ? replacement : source;
  let out = '';
  let last = 0;
  for (const match of folded.matchAll(pattern)) {
    // An empty match (a bare `*`) replaces nothing on its own.
    if (match[0] === '') continue;
    out += source.slice(last, match.index) + replacement;
    last = match.index + match[0].length;
  }
  return out + source.slice(last);
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
  // Formulas mode replaces cell contents, constants included, as the desktop
  // app does. Values mode refuses formulas so a displayed result can never
  // destroy its source expression.
  if (lookIn !== 'formulas' && formula !== null) return false;
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
  const matches = findAll(state, opts).flatMap((m) => {
    const formula = wb.cellFormula(m.addr);
    return opts.lookIn === 'formulas' || formula === null ? [{ addr: m.addr, formula }] : [];
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
      const { formula } = m;
      if (!isCellWritable(state, m.addr)) {
        warnProtected(m.addr);
        continue;
      }
      const value = wb.getValue(m.addr);
      const cur =
        opts.lookIn === 'formulas'
          ? formula !== null
            ? formulaDisplay(formula, opts)
            : valueDisplay(value, opts)
          : workbookDisplay(wb, m.addr, valueDisplay(value, opts));
      const next = applySubstitution(cur, opts, replacement);
      if (next === cur) continue;
      // A replacement that is not a valid formula leaves the cell as it was.
      try {
        writeInput(wb, m.addr, next, store, opts.locale);
      } catch {
        continue;
      }
      count += 1;
    }
  });
  return count;
}
