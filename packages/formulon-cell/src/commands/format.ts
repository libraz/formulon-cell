import { addrKey } from '../engine/address.js';
import type { Addr, Range } from '../engine/types.js';
import { normalizeFormatLocale } from '../format/locale.js';
import { formatWithPending, sameAddr } from '../store/pending-format.js';
import { selectionCoversRange, subtractRange } from '../store/selection-geometry.js';
import {
  type CellAlign,
  type CellBorderSide,
  type CellBorderStyle,
  type CellBorders,
  type CellFormat,
  type CellVAlign,
  mutators,
  type NumFmt,
  type SpreadsheetStore,
  type State,
} from '../store/store.js';
import { interactionControllerFor } from './interaction-controller.js';
import type { InteractionOperation, InteractionOrigin } from './interaction-policy.js';
import { gateProtection, isCellWritable } from './protection.js';

export { formatNumber } from '../format/number-format.js';

type ToggleKey = 'bold' | 'italic' | 'underline' | 'strike';

const MAX_MATERIALIZED_FORMAT_CELLS = 100_000;
const MAX_XLSX_ROW = 1_048_575;
const MAX_XLSX_COL = 16_383;

export interface SelectionFormatOptions {
  allowPending?: boolean;
  origin?: InteractionOrigin;
  commandId?: string;
}

interface SelectionFormatContext {
  origin: InteractionOrigin;
  commandId?: string;
}

const selectionFormatOrigin = new WeakMap<SpreadsheetStore, SelectionFormatContext>();

export function withSelectionFormatOrigin<T>(
  store: SpreadsheetStore,
  origin: InteractionOrigin,
  fn: () => T,
  commandId?: string,
): T {
  const previous = selectionFormatOrigin.get(store);
  const inheritedCommandId = commandId ?? previous?.commandId;
  selectionFormatOrigin.set(store, {
    origin,
    ...(inheritedCommandId !== undefined ? { commandId: inheritedCommandId } : {}),
  });
  try {
    return fn();
  } finally {
    if (previous === undefined) selectionFormatOrigin.delete(store);
    else selectionFormatOrigin.set(store, previous);
  }
}

const rangeContainsAddr = (range: Range, addr: Addr): boolean =>
  addr.sheet === range.sheet &&
  addr.row >= range.r0 &&
  addr.row <= range.r1 &&
  addr.col >= range.c0 &&
  addr.col <= range.c1;

const rangesIntersect = (left: Range, right: Range): boolean =>
  left.sheet === right.sheet &&
  left.r0 <= right.r1 &&
  left.r1 >= right.r0 &&
  left.c0 <= right.c1 &&
  left.c1 >= right.c0;

const validSelectionRange = (range: Range, sheet: number): boolean =>
  range.sheet === sheet &&
  Number.isInteger(range.sheet) &&
  Number.isInteger(range.r0) &&
  Number.isInteger(range.c0) &&
  Number.isInteger(range.r1) &&
  Number.isInteger(range.c1) &&
  range.sheet >= 0 &&
  range.r0 >= 0 &&
  range.c0 >= 0 &&
  range.r0 <= range.r1 &&
  range.c0 <= range.c1 &&
  range.r1 <= MAX_XLSX_ROW &&
  range.c1 <= MAX_XLSX_COL;

const selectionRanges = (state: State): Range[] | null => {
  const ranges = [state.selection.range, ...(state.selection.extraRanges ?? [])];
  const primary = ranges[0];
  if (!primary || !validSelectionRange(primary, primary.sheet)) return null;
  if (!ranges.every((range) => validSelectionRange(range, primary.sheet))) return null;
  return ranges.map((range) => ({ ...range }));
};

const disjointRanges = (ranges: readonly Range[]): Range[] => {
  const result: Range[] = [];
  for (const range of ranges) {
    let pieces = [{ ...range }];
    for (const cover of result) {
      pieces = pieces.flatMap((piece) => subtractRange(piece, cover));
      if (pieces.length === 0) break;
    }
    result.push(...pieces);
  }
  return result;
};

const mergeCoverageValid = (state: State, ranges: readonly Range[]): boolean => {
  const selection = { range: ranges[0] as Range, extraRanges: ranges.slice(1) as Range[] };
  for (const merge of state.merges.byAnchor.values()) {
    if (ranges.some((range) => rangesIntersect(range, merge))) {
      if (!selectionCoversRange(selection, merge)) return false;
    }
  }
  return true;
};

const materializeSelection = (ranges: readonly Range[]): Addr[] | null => {
  const cells = new Map<string, Addr>();
  for (const range of disjointRanges(ranges)) {
    for (let row = range.r0; row <= range.r1; row += 1) {
      for (let col = range.c0; col <= range.c1; col += 1) {
        const addr = { sheet: range.sheet, row, col };
        cells.set(addrKey(addr), addr);
        if (cells.size > MAX_MATERIALIZED_FORMAT_CELLS) return null;
      }
    }
  }
  return [...cells.values()];
};

const selectionPlan = (state: State): { ranges: Range[]; cells: Addr[] } | null => {
  const ranges = selectionRanges(state);
  if (!ranges || !mergeCoverageValid(state, ranges)) return null;
  const normalized = disjointRanges(ranges);
  const cells = materializeSelection(normalized);
  return cells ? { ranges: normalized, cells } : null;
};

export interface SelectionFormatPlan {
  readonly ranges: readonly Range[];
  readonly cells: readonly Addr[];
}

/** Plan the complete primary-plus-extra selection for dialog-style actions. */
export function planSelectionFormat(state: State): SelectionFormatPlan | null {
  const plan = selectionPlan(state);
  if (!plan) return null;
  return {
    ranges: plan.ranges.map((range) => ({ ...range })),
    cells: plan.cells.map((addr) => ({ ...addr })),
  };
}

const selectionIncludes = (ranges: readonly Range[], addr: Addr): boolean =>
  ranges.some((range) => rangeContainsAddr(range, addr));

const authorizeSelection = (
  store: SpreadsheetStore,
  ranges: readonly Range[],
  _cells: readonly Addr[],
  options: SelectionFormatOptions,
): boolean => {
  const controller = interactionControllerFor(store);
  if (!controller?.policy) return true;
  const ambient = selectionFormatOrigin.get(store);
  const decision = controller.canExecute({
    operation: 'format',
    origin: options.origin ?? ambient?.origin ?? 'instanceApi',
    commandId: options.commandId ?? ambient?.commandId,
    effects: ranges.map((range) => ({ kind: 'range' as const, range })),
  });
  return decision.allowed;
};

const addrFromKey = (key: string): Addr | null => {
  const parts = key.split(':').map(Number);
  const sheet = parts[0];
  const row = parts[1];
  const col = parts[2];
  if (
    typeof sheet !== 'number' ||
    typeof row !== 'number' ||
    typeof col !== 'number' ||
    !Number.isInteger(sheet) ||
    !Number.isInteger(row) ||
    !Number.isInteger(col)
  ) {
    return null;
  }
  return { sheet, row, col };
};

const singleCellAddr = (range: Range): { sheet: number; row: number; col: number } | null =>
  range.r0 === range.r1 && range.c0 === range.c1
    ? { sheet: range.sheet, row: range.r0, col: range.c0 }
    : null;

const canStagePendingFormat = (state: State, range: Range): boolean => {
  if ((state.selection.extraRanges?.length ?? 0) > 0 || state.ui.editor.kind !== 'idle') {
    return false;
  }
  const addr = singleCellAddr(range);
  if (!addr || !sameAddr(addr, state.selection.active)) return false;
  const key = addrKey(addr);
  if (state.format.formats.has(key)) return false;
  const cell = state.data.cells.get(key);
  return !cell || (cell.value.kind === 'blank' && !cell.formula);
};

const stagePendingFormat = (
  state: State,
  store: SpreadsheetStore,
  range: Range,
  patch: Partial<CellFormat>,
): void => {
  const addr = singleCellAddr(range);
  if (!addr) return;
  const previous = sameAddr(state.ui.pendingFormat?.addr ?? addr, addr)
    ? (state.ui.pendingFormat?.format ?? {})
    : {};
  const next: Partial<CellFormat> = { ...previous, ...patch };
  if (patch.borders) next.borders = { ...(previous.borders ?? {}), ...patch.borders };
  mutators.setPendingFormat(store, { addr, format: next });
};

interface SelectionFormatEntry {
  addr: Addr;
  patch: Partial<CellFormat> | null;
}

const mergeCellFormat = (
  previous: CellFormat | undefined,
  patch: Partial<CellFormat> | null,
): CellFormat | undefined => {
  if (patch === null) return undefined;
  const next: CellFormat = { ...(previous ?? {}), ...patch };
  if (patch.borders) next.borders = { ...(previous?.borders ?? {}), ...patch.borders };
  return next;
};

const commitSelectionFormatEntries = (
  state: State,
  store: SpreadsheetStore,
  ranges: readonly Range[],
  cells: readonly Addr[],
  entries: readonly SelectionFormatEntry[],
  options: SelectionFormatOptions,
): boolean => {
  if (!authorizeSelection(store, ranges, cells, options)) return false;
  const pendingEntry = entries.length === 1 ? entries[0] : undefined;
  const pendingAddr = singleCellAddr(ranges[0] as Range);
  if (
    pendingEntry &&
    pendingEntry.patch !== null &&
    options.allowPending !== false &&
    pendingAddr !== null &&
    isCellWritable(state, pendingAddr) &&
    canStagePendingFormat(state, ranges[0] as Range)
  ) {
    stagePendingFormat(state, store, ranges[0] as Range, pendingEntry.patch);
    return true;
  }

  const writable = entries.filter((entry) => isCellWritable(state, entry.addr));
  if (writable.length === 0) return false;
  store.setState((current) => {
    const formats = new Map(current.format.formats);
    for (const entry of writable) {
      const key = addrKey(entry.addr);
      const next = mergeCellFormat(formats.get(key), entry.patch);
      if (next) formats.set(key, next);
      else formats.delete(key);
    }
    return {
      ...current,
      format: { ...current.format, formats },
      ui: { ...current.ui, pendingFormat: null },
    };
  });
  return true;
};

export function applySelectionFormatPatch(
  state: State,
  store: SpreadsheetStore,
  patch: Partial<CellFormat> | null,
  options: SelectionFormatOptions = {},
): boolean {
  const plan = selectionPlan(state);
  if (!plan) return false;
  return commitSelectionFormatEntries(
    state,
    store,
    plan.ranges,
    plan.cells,
    plan.cells.map((addr) => ({ addr, patch })),
    options,
  );
}

const sparseSelectionPlan = (
  state: State,
  store: SpreadsheetStore,
  options: SelectionFormatOptions,
): { ranges: Range[]; cells: Addr[] } | null => {
  const ranges = selectionRanges(state);
  if (!ranges || !mergeCoverageValid(state, ranges)) return null;
  const normalized = disjointRanges(ranges);
  const cells: Addr[] = [];
  for (const key of state.format.formats.keys()) {
    const addr = addrFromKey(key);
    if (addr && selectionIncludes(normalized, addr)) cells.push(addr);
  }
  const pending = state.ui.pendingFormat?.addr;
  if (pending && selectionIncludes(normalized, pending)) cells.push(pending);
  if (!authorizeSelection(store, normalized, cells, options)) return null;
  return { ranges: normalized, cells };
};

const clearSelectionFormat = (
  state: State,
  store: SpreadsheetStore,
  visualOnly: boolean,
  options: SelectionFormatOptions = {},
): boolean => {
  const plan = sparseSelectionPlan(state, store, options);
  if (!plan) return false;
  const selected = new Set(plan.cells.map(addrKey));
  store.setState((current) => {
    const formats = new Map(current.format.formats);
    for (const [key, format] of current.format.formats) {
      if (!selected.has(key) || !isCellWritable(current, addrFromKey(key) as Addr)) continue;
      if (!visualOnly) formats.delete(key);
      else {
        const next = stripVisualFormat(format);
        if (next) formats.set(key, next);
        else formats.delete(key);
      }
    }
    const pending = current.ui.pendingFormat?.addr;
    const pendingSelected = pending
      ? selectionIncludes(plan.ranges, pending) && isCellWritable(current, pending)
      : false;
    return {
      ...current,
      format: { ...current.format, formats },
      ui: {
        ...current.ui,
        ...(pendingSelected ? { pendingFormat: null } : {}),
      },
    };
  });
  return true;
};

/** Apply a partial format patch to every writable cell in `range`. When
 *  the sheet is unprotected this falls back to the bulk
 *  `mutators.setRangeFormat` path (one setState call); when protected it
 *  walks the range cell-by-cell and only writes through the cells that
 *  are explicitly unlocked. Returns false when the entire range is gated
 *  so the caller can short-circuit (e.g. emit a single warning). */
export function applyFormatPatch(
  state: State,
  store: SpreadsheetStore,
  range: Range,
  patch: Partial<CellFormat> | null,
  opts: { allowPending?: boolean } = {},
): boolean {
  const allowed = gateProtection(state, range);
  if (allowed === null) return false;
  if (patch !== null && opts.allowPending !== false && canStagePendingFormat(state, range)) {
    stagePendingFormat(state, store, range, patch);
    return true;
  }
  if (state.ui.pendingFormat) mutators.setPendingFormat(store, null);
  // Fast path: sheet unprotected, or entire range writable.
  // We can still use the bulk mutator because per-cell-locked cells inside
  // a partially-unlocked range need to be skipped — fall through to the
  // per-cell loop in that case.
  if (!state.protection.protectedSheets.has(range.sheet)) {
    mutators.setRangeFormat(store, range, patch);
    return true;
  }
  // Per-cell loop respecting individual lock flags.
  const sheet = range.sheet;
  let wroteAny = false;
  for (let r = range.r0; r <= range.r1; r += 1) {
    for (let c = range.c0; c <= range.c1; c += 1) {
      const addr = { sheet, row: r, col: c };
      if (!isCellWritable(state, addr)) continue;
      mutators.setCellFormat(store, addr, patch);
      wroteAny = true;
    }
  }
  return wroteAny;
}

const selectionAllHave = (
  state: State,
  cells: readonly Addr[],
  predicate: (f: CellFormat | undefined) => boolean,
): boolean =>
  cells.every((addr) => {
    const pending = state.ui.pendingFormat;
    const key = addrKey(addr);
    const format =
      pending && sameAddr(pending.addr, addr) && !state.format.formats.has(key)
        ? formatWithPending(state, addr)
        : state.format.formats.get(key);
    return predicate(format);
  });

const selectionAnyHave = (
  state: State,
  cells: readonly Addr[],
  predicate: (f: CellFormat | undefined) => boolean,
): boolean =>
  cells.some((addr) => {
    const pending = state.ui.pendingFormat;
    const key = addrKey(addr);
    const format =
      pending && sameAddr(pending.addr, addr) && !state.format.formats.has(key)
        ? formatWithPending(state, addr)
        : state.format.formats.get(key);
    return predicate(format);
  });

function toggleFlag(state: State, store: SpreadsheetStore, key: ToggleKey): void {
  const plan = selectionPlan(state);
  if (!plan) return;
  const allOn = selectionAllHave(state, plan.cells, (f) => f?.[key] === true);
  const patch = { [key]: !allOn } as Partial<CellFormat>;
  commitSelectionFormatEntries(
    state,
    store,
    plan.ranges,
    plan.cells,
    plan.cells.map((addr) => ({ addr, patch })),
    {},
  );
}

export function toggleBold(state: State, store: SpreadsheetStore): void {
  toggleFlag(state, store, 'bold');
}

export function toggleItalic(state: State, store: SpreadsheetStore): void {
  toggleFlag(state, store, 'italic');
}

export function toggleUnderline(state: State, store: SpreadsheetStore): void {
  toggleFlag(state, store, 'underline');
}

export function toggleStrike(state: State, store: SpreadsheetStore): void {
  toggleFlag(state, store, 'strike');
}

export function setAlign(state: State, store: SpreadsheetStore, align: CellAlign): void {
  applySelectionFormatPatch(state, store, { align });
}

export function setVAlign(state: State, store: SpreadsheetStore, vAlign: CellVAlign): void {
  applySelectionFormatPatch(state, store, { vAlign });
}

export function toggleWrap(state: State, store: SpreadsheetStore): void {
  const plan = selectionPlan(state);
  if (!plan) return;
  const allOn = selectionAllHave(state, plan.cells, (f) => f?.wrap === true);
  commitSelectionFormatEntries(
    state,
    store,
    plan.ranges,
    plan.cells,
    plan.cells.map((addr) => ({ addr, patch: { wrap: !allOn } })),
    {},
  );
}

export function bumpIndent(state: State, store: SpreadsheetStore, delta: 1 | -1): void {
  const plan = selectionPlan(state);
  if (!plan) return;
  const entries = plan.cells.map((addr) => {
    const cur = formatWithPending(state, addr)?.indent ?? 0;
    return {
      addr,
      patch: { indent: Math.max(0, Math.min(15, cur + delta)) },
    };
  });
  commitSelectionFormatEntries(state, store, plan.ranges, plan.cells, entries, {});
}

export function setRotation(state: State, store: SpreadsheetStore, deg: number): void {
  const r = Math.max(-90, Math.min(90, Math.round(deg)));
  applySelectionFormatPatch(state, store, { rotation: r });
}

export function setNumFmt(state: State, store: SpreadsheetStore, fmt: NumFmt): void {
  applySelectionFormatPatch(state, store, { numFmt: fmt });
}

const defaultCurrencySymbol = (locale: string): string =>
  normalizeFormatLocale(locale).startsWith('ja') ? '¥' : '$';

export function cycleCurrency(state: State, store: SpreadsheetStore, locale = 'en-US'): void {
  const plan = selectionPlan(state);
  if (!plan) return;
  const hasCurrency = selectionAnyHave(state, plan.cells, (f) => f?.numFmt?.kind === 'currency');
  const next: NumFmt = hasCurrency
    ? { kind: 'general' }
    : { kind: 'currency', decimals: 2, symbol: defaultCurrencySymbol(locale) };
  commitSelectionFormatEntries(
    state,
    store,
    plan.ranges,
    plan.cells,
    plan.cells.map((addr) => ({ addr, patch: { numFmt: next } })),
    {},
  );
}

export function cyclePercent(state: State, store: SpreadsheetStore): void {
  const plan = selectionPlan(state);
  if (!plan) return;
  const hasPercent = selectionAnyHave(state, plan.cells, (f) => f?.numFmt?.kind === 'percent');
  const next: NumFmt = hasPercent ? { kind: 'general' } : { kind: 'percent', decimals: 0 };
  commitSelectionFormatEntries(
    state,
    store,
    plan.ranges,
    plan.cells,
    plan.cells.map((addr) => ({ addr, patch: { numFmt: next } })),
    {},
  );
}

const clampDecimals = (n: number): number => Math.max(0, Math.min(10, n));

export function bumpDecimals(state: State, store: SpreadsheetStore, delta: 1 | -1): void {
  const plan = selectionPlan(state);
  if (!plan) return;
  const entries: SelectionFormatEntry[] = [];
  for (const addr of plan.cells) {
    const cur = formatWithPending(state, addr)?.numFmt;
    let nextFmt: NumFmt | undefined;
    if (!cur || cur.kind === 'general') {
      if (delta === 1) nextFmt = { kind: 'fixed', decimals: 2 };
    } else if (cur.kind === 'fixed') {
      nextFmt = { kind: 'fixed', decimals: clampDecimals(cur.decimals + delta) };
    } else if (cur.kind === 'currency') {
      nextFmt = {
        kind: 'currency',
        decimals: clampDecimals(cur.decimals + delta),
        ...(cur.symbol !== undefined ? { symbol: cur.symbol } : {}),
      };
    } else if (cur.kind === 'percent') {
      nextFmt = { kind: 'percent', decimals: clampDecimals(cur.decimals + delta) };
    }
    if (nextFmt) entries.push({ addr, patch: { numFmt: nextFmt } });
  }
  if (entries.length > 0)
    commitSelectionFormatEntries(state, store, plan.ranges, plan.cells, entries, {});
}

export function setBorders(state: State, store: SpreadsheetStore, sides: CellBorders): void {
  applySelectionFormatPatch(state, store, { borders: sides });
}

export type BorderPreset =
  | 'none'
  | 'outline'
  | 'thickOutline'
  | 'all'
  | 'inside'
  | 'insideHorizontal'
  | 'insideVertical'
  | 'top'
  | 'bottom'
  | 'left'
  | 'right'
  | 'doubleBottom'
  | 'thickBottom'
  | 'topAndBottom'
  | 'topAndThickBottom'
  | 'topAndDoubleBottom'
  | 'diagonalDown'
  | 'diagonalUp';

export interface SelectionBorderAction {
  readonly preset: BorderPreset;
  readonly style?: CellBorderStyle;
  readonly color?: string;
}

export interface SelectionFormatAction {
  readonly patch: Partial<CellFormat>;
  readonly border?: SelectionBorderAction;
  readonly operations?: readonly InteractionOperation[];
}

type BorderSideName = keyof CellBorders;

const borderEntriesForPlan = (
  plan: { ranges: readonly Range[]; cells: readonly Addr[] },
  preset: BorderPreset,
  side: CellBorderSide,
  thickSide: CellBorderSide,
  doubleSide: CellBorderSide,
): SelectionFormatEntry[] => {
  const entries = new Map<string, SelectionFormatEntry>();
  const add = (addr: Addr, name: BorderSideName, value: CellBorderSide): void => {
    const key = addrKey(addr);
    const current = entries.get(key);
    if (!current) {
      entries.set(key, { addr, patch: { borders: { [name]: value } } });
      return;
    }
    if (current.patch?.borders) current.patch.borders[name] = value;
  };

  for (const range of plan.ranges) {
    for (let row = range.r0; row <= range.r1; row += 1) {
      for (let col = range.c0; col <= range.c1; col += 1) {
        const addr = { sheet: range.sheet, row, col };
        if (preset === 'all') {
          add(addr, 'top', side);
          add(addr, 'right', side);
          add(addr, 'bottom', side);
          add(addr, 'left', side);
        } else if (preset === 'none') {
          add(addr, 'top', false);
          add(addr, 'right', false);
          add(addr, 'bottom', false);
          add(addr, 'left', false);
          add(addr, 'diagonalDown', false);
          add(addr, 'diagonalUp', false);
        } else if (preset === 'inside') {
          if (row > range.r0) add(addr, 'top', side);
          if (col > range.c0) add(addr, 'left', side);
        } else if (preset === 'insideHorizontal') {
          if (row > range.r0) add(addr, 'top', side);
        } else if (preset === 'insideVertical') {
          if (col > range.c0) add(addr, 'left', side);
        } else if (preset === 'diagonalDown') {
          add(addr, 'diagonalDown', side);
        } else if (preset === 'diagonalUp') {
          add(addr, 'diagonalUp', side);
        } else if (preset === 'outline') {
          if (row === range.r0) add(addr, 'top', side);
          if (row === range.r1) add(addr, 'bottom', side);
          if (col === range.c0) add(addr, 'left', side);
          if (col === range.c1) add(addr, 'right', side);
        } else if (preset === 'thickOutline') {
          if (row === range.r0) add(addr, 'top', thickSide);
          if (row === range.r1) add(addr, 'bottom', thickSide);
          if (col === range.c0) add(addr, 'left', thickSide);
          if (col === range.c1) add(addr, 'right', thickSide);
        } else if (preset === 'top' && row === range.r0) {
          add(addr, 'top', side);
        } else if (preset === 'bottom' && row === range.r1) {
          add(addr, 'bottom', side);
        } else if (preset === 'left' && col === range.c0) {
          add(addr, 'left', side);
        } else if (preset === 'right' && col === range.c1) {
          add(addr, 'right', side);
        } else if (preset === 'doubleBottom' && row === range.r1) {
          add(addr, 'bottom', doubleSide);
        } else if (preset === 'thickBottom' && row === range.r1) {
          add(addr, 'bottom', thickSide);
        } else if (preset === 'topAndBottom') {
          if (row === range.r0) add(addr, 'top', side);
          if (row === range.r1) add(addr, 'bottom', side);
        } else if (preset === 'topAndThickBottom') {
          if (row === range.r0) add(addr, 'top', side);
          if (row === range.r1) add(addr, 'bottom', thickSide);
        } else if (preset === 'topAndDoubleBottom') {
          if (row === range.r0) add(addr, 'top', side);
          if (row === range.r1) add(addr, 'bottom', doubleSide);
        }
      }
    }
  }
  return [...entries.values()];
};

const operationsForFormatPatch = (patch: Partial<CellFormat>): Set<InteractionOperation> => {
  const operations = new Set<InteractionOperation>();
  for (const key of Object.keys(patch) as Array<keyof CellFormat>) {
    if (key === 'comment' || key === 'commentAuthor') operations.add('comment');
    else if (key === 'hyperlink' || key === 'hyperlinkDisplay' || key === 'hyperlinkTooltip') {
      operations.add('hyperlink');
    } else if (key === 'validation') operations.add('validation');
    else operations.add('format');
  }
  return operations;
};

const mergeActionCellFormat = (
  previous: CellFormat | undefined,
  patch: Partial<CellFormat>,
): CellFormat | undefined => {
  const next: CellFormat = { ...(previous ?? {}), ...patch };
  for (const key of Object.keys(patch) as Array<keyof CellFormat>) {
    if (patch[key] === undefined) delete next[key];
  }
  if (patch.borders) {
    const borders = { ...(previous?.borders ?? {}), ...patch.borders };
    for (const key of Object.keys(patch.borders) as Array<keyof CellBorders>) {
      if (patch.borders[key] === undefined) delete borders[key];
    }
    if (Object.keys(borders).length === 0) delete next.borders;
    else next.borders = borders;
  }

  if (Object.hasOwn(patch, 'hyperlink')) {
    const hyperlink = patch.hyperlink;
    if (hyperlink && previous?.hyperlink === hyperlink) {
      if (!Object.hasOwn(patch, 'hyperlinkDisplay')) {
        if (previous.hyperlinkDisplay === undefined) delete next.hyperlinkDisplay;
        else next.hyperlinkDisplay = previous.hyperlinkDisplay;
      }
      if (!Object.hasOwn(patch, 'hyperlinkTooltip')) {
        if (previous.hyperlinkTooltip === undefined) delete next.hyperlinkTooltip;
        else next.hyperlinkTooltip = previous.hyperlinkTooltip;
      }
    } else {
      if (!Object.hasOwn(patch, 'hyperlinkDisplay')) delete next.hyperlinkDisplay;
      if (!Object.hasOwn(patch, 'hyperlinkTooltip')) delete next.hyperlinkTooltip;
    }
  }
  return Object.keys(next).length > 0 ? next : undefined;
};

const authorizeSelectionOperations = (
  store: SpreadsheetStore,
  ranges: readonly Range[],
  operations: Iterable<InteractionOperation>,
  options: SelectionFormatOptions,
): boolean => {
  const controller = interactionControllerFor(store);
  if (!controller?.policy) return true;
  const ambient = selectionFormatOrigin.get(store);
  const origin = options.origin ?? ambient?.origin ?? 'instanceApi';
  const commandId = options.commandId ?? ambient?.commandId;
  for (const operation of operations) {
    if (
      !controller.canExecute({
        operation,
        origin,
        commandId,
        effects: ranges.map((range) => ({ kind: 'range' as const, range })),
      }).allowed
    ) {
      return false;
    }
  }
  return true;
};

/** Apply a touched dialog action to the complete planned union in one store publication. */
export function applySelectionFormatAction(
  state: State,
  store: SpreadsheetStore,
  action: SelectionFormatAction,
  options: SelectionFormatOptions = {},
): boolean {
  const plan = selectionPlan(state);
  if (!plan) return false;

  const operations = operationsForFormatPatch(action.patch);
  for (const operation of action.operations ?? []) operations.add(operation);
  if (action.border) operations.add('format');
  if (operations.size === 0) return false;
  if (!authorizeSelectionOperations(store, plan.ranges, operations, options)) return false;

  if (
    !action.border &&
    options.allowPending !== false &&
    plan.cells.length === 1 &&
    plan.ranges.length === 1 &&
    plan.cells[0] !== undefined &&
    isCellWritable(state, plan.cells[0]) &&
    canStagePendingFormat(state, plan.ranges[0] as Range)
  ) {
    stagePendingFormat(state, store, plan.ranges[0] as Range, action.patch);
    return true;
  }

  const entries = new Map<string, Partial<CellFormat>>();
  for (const addr of plan.cells) entries.set(addrKey(addr), { ...action.patch });
  if (action.border) {
    const style = action.border.style ?? 'thin';
    const side: CellBorderSide =
      action.border.color === undefined ? { style } : { style, color: action.border.color };
    const thickSide: CellBorderSide =
      action.border.color === undefined
        ? { style: 'thick' }
        : { style: 'thick', color: action.border.color };
    const doubleSide: CellBorderSide =
      action.border.color === undefined
        ? { style: 'double' }
        : { style: 'double', color: action.border.color };
    for (const entry of borderEntriesForPlan(
      plan,
      action.border.preset,
      side,
      thickSide,
      doubleSide,
    )) {
      const existing = entries.get(addrKey(entry.addr));
      if (!existing) continue;
      entries.set(addrKey(entry.addr), {
        ...existing,
        borders: { ...(existing.borders ?? {}), ...(entry.patch?.borders ?? {}) },
      });
    }
  }

  const writable = plan.cells.filter((addr) => isCellWritable(state, addr));
  if (writable.length === 0) return false;
  store.setState((current) => {
    const formats = new Map(current.format.formats);
    for (const addr of writable) {
      const key = addrKey(addr);
      const patch = entries.get(key);
      if (!patch) continue;
      const next = mergeActionCellFormat(formatWithPending(state, addr), patch);
      if (next) formats.set(key, next);
      else formats.delete(key);
    }
    return {
      ...current,
      format: { ...current.format, formats },
      ui: { ...current.ui, pendingFormat: null },
    };
  });
  return true;
}

export function setBorderPreset(
  state: State,
  store: SpreadsheetStore,
  preset: BorderPreset,
  style: CellBorderStyle = 'thin',
  color?: string,
): void {
  const plan = selectionPlan(state);
  if (!plan) return;
  const side: CellBorderSide = color === undefined ? { style } : { style, color };
  const thickSide: CellBorderSide =
    color === undefined ? { style: 'thick' } : { style: 'thick', color };
  const doubleSide: CellBorderSide =
    color === undefined ? { style: 'double' } : { style: 'double', color };
  const entries = borderEntriesForPlan(plan, preset, side, thickSide, doubleSide);
  if (entries.length === 0) return;
  commitSelectionFormatEntries(state, store, plan.ranges, plan.cells, entries, {});
}

/** Toolbar default: outline if missing, all-borders if outline present, else clear.
 *  Three-step cycle on repeated clicks. */
export function cycleBorders(state: State, store: SpreadsheetStore): void {
  const plan = selectionPlan(state);
  if (!plan) return;
  const hasAny = selectionAnyHave(state, plan.cells, (f) => {
    const b = f?.borders;
    return !!(b && (b.top || b.right || b.bottom || b.left));
  });
  if (!hasAny) {
    const entries = borderEntriesForPlan(plan, 'outline', true, true, true);
    commitSelectionFormatEntries(state, store, plan.ranges, plan.cells, entries, {});
    return;
  }
  const allFour = selectionAllHave(
    state,
    plan.cells,
    (f) => !!(f?.borders?.top && f.borders.right && f.borders.bottom && f.borders.left),
  );
  const entries = allFour
    ? plan.cells.map((addr) => ({
        addr,
        patch: {
          borders: { top: false, right: false, bottom: false, left: false },
        },
      }))
    : borderEntriesForPlan(plan, 'all', true, true, true);
  commitSelectionFormatEntries(state, store, plan.ranges, plan.cells, entries, {});
}

export function clearFormat(state: State, store: SpreadsheetStore): void {
  clearSelectionFormat(state, store, false);
}

const visualFormatKeys: readonly (keyof CellFormat)[] = [
  'cellStyle',
  'numFmt',
  'bold',
  'italic',
  'underline',
  'strike',
  'align',
  'vAlign',
  'wrap',
  'shrinkToFit',
  'indent',
  'rotation',
  'textDirection',
  'borders',
  'color',
  'fill',
  'fillPattern',
  'fillPatternColor',
  'fontFamily',
  'fontSize',
];

function stripVisualFormat(format: CellFormat): CellFormat | null {
  const next: CellFormat = { ...format };
  for (const key of visualFormatKeys) delete next[key];
  return Object.keys(next).length > 0 ? next : null;
}

/** Clear only visual cell formatting, preserving cell metadata such as
 *  comments, hyperlinks, data validation, and protection flags. Mirrors
 *  Excel's Home > Clear > Clear Formats more closely than `clearFormat`,
 *  which is the low-level full-format reset used by Clear All. */
export function clearVisualFormat(state: State, store: SpreadsheetStore): void {
  clearSelectionFormat(state, store, true);
}

/** Set or clear the font color across the selection. Pass `null` to clear. */
export function setFontColor(state: State, store: SpreadsheetStore, color: string | null): void {
  applySelectionFormatPatch(state, store, { color: color ?? undefined });
}

/** Set or clear the fill (background) color across the selection. */
export function setFillColor(state: State, store: SpreadsheetStore, color: string | null): void {
  applySelectionFormatPatch(state, store, { fill: color ?? undefined });
}

/** Update font family and/or size across the selection. */
export function setFont(
  state: State,
  store: SpreadsheetStore,
  patch: { fontFamily?: string | null; fontSize?: number | null },
): void {
  const next: Partial<CellFormat> = {};
  if (patch.fontFamily !== undefined) {
    next.fontFamily = patch.fontFamily ?? undefined;
  }
  if (patch.fontSize !== undefined) {
    next.fontSize = patch.fontSize ?? undefined;
  }
  applySelectionFormatPatch(state, store, next);
}
