import { addrKey, parseAddrKey } from '../engine/address.js';
import { findDependents, findPrecedents } from '../engine/refs-graph.js';
import type { Addr, Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { rangeValidationCells, validationAt } from '../store/range-validations.js';
import { rangeContainsAddr } from '../store/selection-geometry.js';
import type { SpreadsheetStore, State } from '../store/store.js';
import type { SelectionSlice } from '../store/types.js';
import { listComments } from './comment.js';
import { currentRegion, formatHasContent } from './current-region.js';

/** Desktop spreadsheet "Go To Special" categories. `numbers`, `text` and
 *  `errors` are not dialog options; the Find & Select ribbon menu uses them. */
export type GoToSpecialKind =
  | 'comments'
  | 'constants'
  | 'formulas'
  | 'numbers'
  | 'text'
  | 'errors'
  | 'blanks'
  | 'current-region'
  | 'current-array'
  | 'row-differences'
  | 'column-differences'
  | 'precedents'
  | 'dependents'
  | 'last-cell'
  | 'visible-cells'
  | 'conditional-format'
  | 'data-validation';

/** Excel-style subfilters used by the Formulas and Constants choices in
 *  Go To Special. Omitted options mean "all value kinds". */
export interface GoToSpecialValueFilters {
  numbers?: boolean;
  text?: boolean;
  logical?: boolean;
  errors?: boolean;
}

export interface GoToSpecialOptions {
  /** Value-kind checkboxes for `formulas` / `constants`. */
  filters?: GoToSpecialValueFilters;
  /** `precedents` / `dependents`: follow the chain transitively instead of one level. */
  allLevels?: boolean;
  /** `conditional-format` / `data-validation`: only cells sharing the active cell's rules. */
  same?: boolean;
}

export type GoToSpecialResult =
  | { status: 'found'; ranges: Range[] }
  | { status: 'none' }
  | { status: 'too-large' };

/** Largest number of cells a search may select. */
export const MAX_GOTO_SPECIAL_CELLS = 100_000;

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

type CellValue = ReturnType<WorkbookHandle['getValue']>;

interface CellEntry {
  addr: Addr;
  value: CellValue;
  formula: string | null;
}

/** A rectangle's rows `r0..r1` x columns `c0..c1`, without the sheet. */
interface Run {
  r0: number;
  r1: number;
  c0: number;
  c1: number;
}

interface SearchContext {
  wb: WorkbookHandle;
  state: State;
  sheet: number;
  entries: CellEntry[];
  /** Selection rectangles on the active sheet. */
  selected: Range[];
  /** True when the selection is one cell, so the search widens to the sheet. */
  single: boolean;
  options: GoToSpecialOptions;
}

const inRange = (r: Range, row: number, col: number): boolean =>
  row >= r.r0 && row <= r.r1 && col >= r.c0 && col <= r.c1;

const inAny = (ranges: readonly Range[], row: number, col: number): boolean =>
  ranges.some((r) => inRange(r, row, col));

const runArea = (r: Run): number => (r.r1 - r.r0 + 1) * (r.c1 - r.c0 + 1);

const addrsToRuns = (addrs: readonly Addr[]): Run[] => {
  const sorted = [...addrs].sort((a, b) => a.row - b.row || a.col - b.col);
  const runs: Run[] = [];
  let open: Run | null = null;
  for (const a of sorted) {
    if (open && open.r0 === a.row && a.col <= open.c1 + 1) {
      open.c1 = Math.max(open.c1, a.col);
    } else {
      open = { r0: a.row, r1: a.row, c0: a.col, c1: a.col };
      runs.push(open);
    }
  }
  return runs;
};

/** Merge runs into rectangles: vertically touching runs with identical column
 *  spans become one rectangle. Output is ordered by top-left cell. */
const mergeRuns = (sheet: number, runs: readonly Run[]): Range[] => {
  const bySpan = new Map<string, Run[]>();
  for (const run of runs) {
    const key = `${run.c0}:${run.c1}`;
    const list = bySpan.get(key);
    if (list) list.push(run);
    else bySpan.set(key, [run]);
  }
  const merged: Run[] = [];
  for (const list of bySpan.values()) {
    list.sort((a, b) => a.r0 - b.r0);
    let open: Run | null = null;
    for (const run of list) {
      if (open && run.r0 <= open.r1 + 1) {
        open.r1 = Math.max(open.r1, run.r1);
      } else {
        open = { ...run };
        merged.push(open);
      }
    }
  }
  merged.sort((a, b) => a.r0 - b.r0 || a.c0 - b.c0);
  return merged.map((r) => ({ sheet, r0: r.r0, c0: r.c0, r1: r.r1, c1: r.c1 }));
};

const valueKey = (value: CellValue | undefined): string =>
  JSON.stringify(value ?? { kind: 'blank' });

const usedRange = (state: State, sheet: number, entries: readonly CellEntry[]): Range | null => {
  let r0 = Number.POSITIVE_INFINITY;
  let c0 = Number.POSITIVE_INFINITY;
  let r1 = -1;
  let c1 = -1;
  const grow = (row: number, col: number): void => {
    if (row < r0) r0 = row;
    if (row > r1) r1 = row;
    if (col < c0) c0 = col;
    if (col > c1) c1 = col;
  };
  for (const e of entries) grow(e.addr.row, e.addr.col);
  for (const [key, format] of state.format.formats) {
    if (!formatHasContent(format)) continue;
    const addr = parseAddrKey(key);
    if (addr && addr.sheet === sheet) grow(addr.row, addr.col);
  }
  return r1 < 0 ? null : { sheet, r0, c0, r1, c1 };
};

/** Where a search looks: the used range for a single selected cell, otherwise
 *  only the selected rectangles. */
const searchAreas = (ctx: SearchContext): Range[] => {
  if (!ctx.single) return ctx.selected;
  const used = usedRange(ctx.state, ctx.sheet, ctx.entries);
  return used ? [used] : [];
};

const matchesKind = (
  entry: CellEntry,
  kind: GoToSpecialKind,
  filters?: GoToSpecialValueFilters,
): boolean => {
  const { value, formula } = entry;
  switch (kind) {
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
      return value.kind === 'text' && ERROR_SENTINELS.has(value.value);
    default:
      return false;
  }
};

const valueMatchesFilters = (
  value: CellValue,
  filters: GoToSpecialValueFilters | undefined,
): boolean => {
  if (!filters) return true;
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

/** Empty cells inside the areas as row runs, without visiting each cell: rows
 *  holding no value become one full-width run. */
const blankRuns = (areas: readonly Range[], entries: readonly CellEntry[]): Run[] => {
  const runs: Run[] = [];
  for (const area of areas) {
    const colsByRow = new Map<number, number[]>();
    for (const { addr, value } of entries) {
      if (value.kind === 'blank' || !inRange(area, addr.row, addr.col)) continue;
      const cols = colsByRow.get(addr.row);
      if (cols) cols.push(addr.col);
      else colsByRow.set(addr.row, [addr.col]);
    }
    let cursor = area.r0;
    for (const row of [...colsByRow.keys()].sort((a, b) => a - b)) {
      if (cursor < row) runs.push({ r0: cursor, r1: row - 1, c0: area.c0, c1: area.c1 });
      let col = area.c0;
      for (const filled of (colsByRow.get(row) as number[]).sort((a, b) => a - b)) {
        if (col < filled) runs.push({ r0: row, r1: row, c0: col, c1: filled - 1 });
        col = filled + 1;
      }
      if (col <= area.c1) runs.push({ r0: row, r1: row, c0: col, c1: area.c1 });
      cursor = row + 1;
    }
    if (cursor <= area.r1) runs.push({ r0: cursor, r1: area.r1, c0: area.c0, c1: area.c1 });
  }
  return runs;
};

/** Split `lo..hi` into segments that skip every index in `hidden`. */
const visibleSegments = (
  lo: number,
  hi: number,
  hidden: ReadonlySet<number>,
): [number, number][] => {
  const skipped = [...hidden].filter((i) => i >= lo && i <= hi).sort((a, b) => a - b);
  const out: [number, number][] = [];
  let start = lo;
  for (const i of skipped) {
    if (i > start) out.push([start, i - 1]);
    start = i + 1;
  }
  if (start <= hi) out.push([start, hi]);
  return out;
};

const visibleRuns = (state: State, areas: readonly Range[]): Run[] => {
  const runs: Run[] = [];
  for (const area of areas) {
    const rows = visibleSegments(area.r0, area.r1, state.layout.hiddenRows);
    const cols = visibleSegments(area.c0, area.c1, state.layout.hiddenCols);
    for (const [r0, r1] of rows) for (const [c0, c1] of cols) runs.push({ r0, r1, c0, c1 });
  }
  return runs;
};

/** Cells whose value differs from the active cell's row (`row`) or column
 *  (`col`) counterpart on the same line, inside the areas. */
const differenceRuns = (
  ctx: SearchContext,
  areas: readonly Range[],
  axis: 'row' | 'col',
): Run[] => {
  const { sheet, entries } = ctx;
  const active = ctx.state.selection.active;
  const keys = new Map<string, string>();
  for (const e of entries)
    if (e.value.kind !== 'blank') keys.set(addrKey(e.addr), valueKey(e.value));
  const blank = valueKey(undefined);
  const refPos = axis === 'row' ? active.row : active.col;
  const runs: Run[] = [];
  for (const area of areas) {
    const [l0, l1, p0, p1] =
      axis === 'row' ? [area.c0, area.c1, area.r0, area.r1] : [area.r0, area.r1, area.c0, area.c1];
    const posByLine = new Map<number, Map<number, string>>();
    for (const { addr, value } of entries) {
      if (value.kind === 'blank' || !inRange(area, addr.row, addr.col)) continue;
      const line = axis === 'row' ? addr.col : addr.row;
      const pos = axis === 'row' ? addr.row : addr.col;
      const map = posByLine.get(line) ?? new Map<number, string>();
      map.set(pos, valueKey(value));
      posByLine.set(line, map);
    }
    const emit = (line: number, a: number, b: number): void => {
      runs.push(
        axis === 'row'
          ? { r0: a, r1: b, c0: line, c1: line }
          : { r0: line, r1: line, c0: a, c1: b },
      );
    };
    for (let line = l0; line <= l1; line += 1) {
      const refAddr = {
        sheet,
        row: axis === 'row' ? active.row : line,
        col: axis === 'row' ? line : active.col,
      };
      const ref = keys.get(addrKey(refAddr)) ?? blank;
      const positions = posByLine.get(line);
      if (ref === blank) {
        for (const [pos, key] of positions ?? [])
          if (key !== blank && pos !== refPos) emit(line, pos, pos);
        continue;
      }
      const same = new Set<number>([refPos]);
      for (const [pos, key] of positions ?? []) if (key === ref) same.add(pos);
      let cursor = p0;
      for (const pos of [...same].filter((p) => p >= p0 && p <= p1).sort((a, b) => a - b)) {
        if (cursor < pos) emit(line, cursor, pos - 1);
        cursor = pos + 1;
      }
      if (cursor <= p1) emit(line, cursor, p1);
    }
  }
  return runs;
};

/** Cells reached from the selection through precedent / dependent links on the
 *  active sheet. Returns null once the result outgrows the selection cap. */
const traceAddrs = (ctx: SearchContext, kind: 'precedents' | 'dependents'): Addr[] | null => {
  const next = kind === 'precedents' ? findPrecedents : findDependents;
  const seeds: Addr[] = [];
  for (const r of ctx.selected) {
    for (let row = r.r0; row <= r.r1; row += 1) {
      for (let col = r.c0; col <= r.c1; col += 1) {
        seeds.push({ sheet: ctx.sheet, row, col });
        if (seeds.length > MAX_GOTO_SPECIAL_CELLS) return null;
      }
    }
  }
  const seen = new Set<string>();
  const found: Addr[] = [];
  let frontier = seeds;
  while (frontier.length > 0) {
    const following: Addr[] = [];
    for (const addr of frontier) {
      for (const n of next(ctx.wb, addr)) {
        if (n.sheet !== ctx.sheet) continue;
        const key = addrKey(n);
        if (seen.has(key)) continue;
        seen.add(key);
        if (seen.size > MAX_GOTO_SPECIAL_CELLS) return null;
        found.push(n);
        following.push(n);
      }
    }
    if (!ctx.options.allLevels) break;
    frontier = following;
  }
  return found;
};

/** Cells covered by conditional-format rules, optionally only those covered by
 *  exactly the rules covering the active cell. Null when it outgrows the cap. */
const conditionalAddrs = (ctx: SearchContext): Addr[] | null => {
  const { state, sheet } = ctx;
  const rules = state.conditional.rules;
  const signature = (row: number, col: number): string => {
    const hit: number[] = [];
    rules.forEach((rule, i) => {
      if (rangeContainsAddr(rule.range, { sheet, row, col })) hit.push(i);
    });
    return hit.join(',');
  };
  const active = state.selection.active;
  const wanted = ctx.options.same ? signature(active.row, active.col) : null;
  if (wanted === '') return [];
  const areas = ctx.single ? null : ctx.selected;
  const out = new Map<string, Addr>();
  let budget = 0;
  for (const rule of rules) {
    if (rule.range.sheet !== sheet) continue;
    for (const area of areas ?? [rule.range]) {
      const r0 = Math.max(rule.range.r0, area.r0);
      const r1 = Math.min(rule.range.r1, area.r1);
      const c0 = Math.max(rule.range.c0, area.c0);
      const c1 = Math.min(rule.range.c1, area.c1);
      if (r0 > r1 || c0 > c1) continue;
      budget += (r1 - r0 + 1) * (c1 - c0 + 1);
      if (budget > MAX_GOTO_SPECIAL_CELLS) return null;
      for (let row = r0; row <= r1; row += 1) {
        for (let col = c0; col <= c1; col += 1) {
          const addr = { sheet, row, col };
          if (wanted !== null && signature(row, col) !== wanted) continue;
          out.set(addrKey(addr), addr);
        }
      }
    }
  }
  return [...out.values()];
};

/** Cells carrying data validation, optionally only those whose rule equals the
 *  active cell's. Blank cells count: validation lives on the format slice. */
const validationAddrs = (ctx: SearchContext): Addr[] => {
  const { state, sheet } = ctx;
  const areas = ctx.single ? null : ctx.selected;
  const active = state.selection.active;
  const activeRule = validationAt(state.format, active);
  if (ctx.options.same && !activeRule) return [];
  const wanted = ctx.options.same ? JSON.stringify(activeRule) : null;
  const out: Addr[] = [];
  for (const [key, format] of state.format.formats) {
    if (!format.validation) continue;
    const addr = parseAddrKey(key);
    if (!addr || addr.sheet !== sheet) continue;
    if (areas && !inAny(areas, addr.row, addr.col)) continue;
    if (wanted !== null && JSON.stringify(format.validation) !== wanted) continue;
    out.push(addr);
  }
  for (const { addr, validation } of rangeValidationCells(state, sheet)) {
    if (areas && !inAny(areas, addr.row, addr.col)) continue;
    if (wanted !== null && JSON.stringify(validation) !== wanted) continue;
    out.push(addr);
  }
  return out;
};

const finish = (sheet: number, runs: readonly Run[]): GoToSpecialResult => {
  if (runs.length === 0) return { status: 'none' };
  const ranges = mergeRuns(sheet, runs);
  let total = 0;
  for (const r of ranges) {
    total += runArea({ r0: r.r0, r1: r.r1, c0: r.c0, c1: r.c1 });
    if (total > MAX_GOTO_SPECIAL_CELLS) return { status: 'too-large' };
  }
  return { status: 'found', ranges };
};

const TOO_LARGE: GoToSpecialResult = { status: 'too-large' };

/** Resolve a Go To Special category against the active sheet. Scope follows the
 *  spreadsheet rule: a single selected cell searches the sheet's used range,
 *  a larger selection searches only itself. The result is a list of
 *  rectangles in top-left order; the first rectangle's top-left cell is the
 *  cell to activate. */
export function findMatchingRanges(
  wb: WorkbookHandle,
  store: SpreadsheetStore,
  kind: GoToSpecialKind,
  options: GoToSpecialOptions = {},
): GoToSpecialResult {
  const state = store.getState();
  const sheet = state.data.sheetIndex;
  const selected = [state.selection.range, ...(state.selection.extraRanges ?? [])].filter(
    (r) => r.sheet === sheet,
  );
  const only = selected.length === 1 ? (selected[0] as Range) : null;
  const ctx: SearchContext = {
    wb,
    state,
    sheet,
    entries: [...wb.cells(sheet)],
    selected,
    single: only !== null && only.r0 === only.r1 && only.c0 === only.c1,
    options,
  };
  const active = state.selection.active;

  switch (kind) {
    case 'comments': {
      const areas = searchAreas(ctx);
      const hits = listComments(state, sheet)
        .map((c) => c.addr)
        .filter((a) => inAny(areas, a.row, a.col));
      return finish(sheet, addrsToRuns(hits));
    }
    case 'blanks':
      return finish(sheet, blankRuns(searchAreas(ctx), ctx.entries));
    case 'current-region': {
      const r = active.sheet === sheet ? currentRegion(state, active) : null;
      return finish(sheet, r ? [r] : []);
    }
    case 'current-array': {
      const spill = wb.spillRanges(sheet)?.find((r) => rangeContainsAddr(r, active));
      return finish(sheet, spill ? [spill] : []);
    }
    case 'row-differences':
    case 'column-differences':
      return finish(
        sheet,
        differenceRuns(ctx, searchAreas(ctx), kind === 'row-differences' ? 'row' : 'col'),
      );
    case 'precedents':
    case 'dependents': {
      const addrs = traceAddrs(ctx, kind);
      return addrs ? finish(sheet, addrsToRuns(addrs)) : TOO_LARGE;
    }
    case 'last-cell': {
      const used = usedRange(state, sheet, ctx.entries);
      return finish(sheet, used ? [{ r0: used.r1, r1: used.r1, c0: used.c1, c1: used.c1 }] : []);
    }
    case 'visible-cells':
      return finish(sheet, visibleRuns(state, searchAreas(ctx)));
    case 'conditional-format': {
      const addrs = conditionalAddrs(ctx);
      return addrs ? finish(sheet, addrsToRuns(addrs)) : TOO_LARGE;
    }
    case 'data-validation':
      return finish(sheet, addrsToRuns(validationAddrs(ctx)));
    default: {
      const areas = searchAreas(ctx);
      const hits = ctx.entries
        .filter(
          (e) => inAny(areas, e.addr.row, e.addr.col) && matchesKind(e, kind, options.filters),
        )
        .map((e) => e.addr);
      return finish(sheet, addrsToRuns(hits));
    }
  }
}

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

/** Selection made of the given rectangles. The first rectangle is the primary
 *  range and its top-left cell becomes the active cell. */
export const selectionFromRanges = (ranges: readonly Range[]): SelectionSlice => {
  const [primary, ...rest] = ranges;
  if (!primary) {
    throw new Error('selectionFromRanges: empty range list');
  }
  const first: Addr = { sheet: primary.sheet, row: primary.r0, col: primary.c0 };
  return {
    active: first,
    anchor: { ...first },
    range: { ...primary },
    extraRanges: rest.map((r) => ({ ...r })),
  };
};

/** Build an exact multi-area selection for a list of matched cells, merging
 *  neighbours into rectangles. */
export const selectionFromMatches = (matches: readonly Addr[]): SelectionSlice => {
  if (matches.length === 0) {
    throw new Error('selectionFromMatches: empty match list');
  }
  return selectionFromRanges(mergeRuns((matches[0] as Addr).sheet, addrsToRuns(matches)));
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
  | 'comments'
  | 'object-select'
  | 'selection-pane';

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
  | { kind: 'open-objects' }
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
    selectionTooLarge: string;
  };
}

const RIBBON_SPECIAL_KINDS: Partial<Record<RibbonFindAction, GoToSpecialKind>> = {
  formulas: 'formulas',
  constants: 'constants',
  numbers: 'numbers',
  text: 'text',
  errors: 'errors',
  'conditional-format': 'conditional-format',
  'data-validation': 'data-validation',
  comments: 'comments',
};

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
  if (action === 'object-select' || action === 'selection-pane') return { kind: 'open-objects' };
  const kind = RIBBON_SPECIAL_KINDS[action];
  if (!kind) return { kind: 'noop' };
  const found = findMatchingRanges(workbook, store, kind);
  if (found.status !== 'found') {
    const label =
      found.status === 'too-large'
        ? strings.selectionTooLarge
        : kind === 'comments'
          ? strings.commentNone
          : strings.findNoMatches;
    return {
      kind: 'report',
      report: { title: strings.findSelect, items: [{ severity: 'info', label, detail: '' }] },
    };
  }
  const selection = selectionFromRanges(found.ranges);
  store.setState((state) => ({ ...state, selection }));
  return { kind: 'selected' };
};
