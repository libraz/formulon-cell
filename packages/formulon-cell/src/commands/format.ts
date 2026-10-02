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

export function formatNumber(value: number, fmt: NumFmt | undefined, locale = 'en-US'): string {
  if (!Number.isFinite(value)) return String(value);
  if (!fmt || fmt.kind === 'general') {
    return new Intl.NumberFormat(locale, { maximumFractionDigits: 12 }).format(value);
  }
  if (fmt.kind === 'text') return String(value);
  if (fmt.kind === 'fixed') {
    const negStyle = fmt.negativeStyle ?? 'minus';
    const body = new Intl.NumberFormat(locale, {
      style: 'decimal',
      minimumFractionDigits: fmt.decimals,
      maximumFractionDigits: fmt.decimals,
      useGrouping: !!fmt.thousands,
    }).format(Math.abs(value));
    return applyNegative(value, body, '', negStyle);
  }
  if (fmt.kind === 'currency') {
    const symbol = fmt.symbol ?? '$';
    const negStyle = fmt.negativeStyle ?? 'minus';
    const body = new Intl.NumberFormat(locale, {
      style: 'decimal',
      minimumFractionDigits: fmt.decimals,
      maximumFractionDigits: fmt.decimals,
      useGrouping: true,
    }).format(Math.abs(value));
    return applyNegative(value, body, symbol, negStyle);
  }
  if (fmt.kind === 'percent') {
    return new Intl.NumberFormat(locale, {
      style: 'percent',
      minimumFractionDigits: fmt.decimals,
      maximumFractionDigits: fmt.decimals,
    }).format(value);
  }
  if (fmt.kind === 'scientific') {
    // Excel's built-in Scientific format is `0.00E+00` — the exponent is
    // zero-padded to at least two digits. `toExponential` emits a bare `e+4`,
    // so re-pad to match.
    return value
      .toExponential(fmt.decimals)
      .replace(
        /e([+-])(\d+)/i,
        (_m, sign: string, digits: string) => `E${sign}${digits.padStart(2, '0')}`,
      );
  }
  if (fmt.kind === 'accounting') {
    const symbol = fmt.symbol ?? '$';
    const body = new Intl.NumberFormat(locale, {
      style: 'decimal',
      minimumFractionDigits: fmt.decimals,
      maximumFractionDigits: fmt.decimals,
      useGrouping: true,
    }).format(Math.abs(value));
    if (value === 0) return `${symbol} -`;
    return value < 0 ? `(${symbol}${body})` : `${symbol}${body} `;
  }
  if (fmt.kind === 'date') {
    return renderDateTimePattern(value, fmt.pattern, locale);
  }
  if (fmt.kind === 'time') {
    return renderDateTimePattern(value, fmt.pattern, locale);
  }
  if (fmt.kind === 'datetime') {
    return renderDateTimePattern(value, fmt.pattern, locale);
  }
  if (fmt.kind === 'special') {
    return formatSpecialPattern(value, fmt.pattern);
  }
  if (fmt.kind === 'custom') {
    return formatCustomPattern(value, fmt.pattern, locale);
  }
  return String(value);
}

function applyNegative(
  value: number,
  body: string,
  symbol: string,
  style: 'minus' | 'parens' | 'red' | 'red-parens',
): string {
  const positive = `${symbol}${body}`;
  if (value >= 0) return positive;
  switch (style) {
    case 'parens':
      return `(${symbol}${body})`;
    case 'red':
      return `-${symbol}${body}`; // color is applied at paint time
    case 'red-parens':
      return `(${symbol}${body})`;
    default:
      return `-${symbol}${body}`;
  }
}

/** Spreadsheet serial date → JS Date. spreadsheet epoch is 1899-12-30 (with Lotus 123
 *  1900-leap-year bug compensation already baked in for serials > 60). */
function spreadsheetSerialToDate(serial: number): Date {
  const excel1900Offset = serial > 0 && serial < 60 ? 1 : 0;
  const ms = (serial + excel1900Offset - 25569) * 86_400_000;
  return new Date(ms);
}

const pad2 = (n: number): string => (n < 10 ? `0${n}` : `${n}`);

/** Spreadsheet-style custom format mini-language. Supports section splitting
 *  (pos;neg;zero;text), `0`/`#`/`?` digit placeholders, `.` decimal, `,`
 *  thousands & scaling, `%`, `\\X` escape, `"text"` literals, `[Red]`-style
 *  color tags (stripped — color is applied at paint time), and date tokens
 *  `yyyy`/`yy`/`mmmmm`/`mmmm`/`mmm`/`mm`/`m`/`dddd`/`ddd`/`dd`/`d`/`hh`/`h`/
 *  `ss`/`s` plus `am/pm`. Not exhaustive but covers the patterns spreadsheets
 *  ship in its built-in format codes. */
function formatCustomPattern(value: number, pattern: string, locale: string): string {
  // Split into up to four sections on ';' that aren't inside a quoted literal
  //  or a bracketed tag. Spreadsheets allow: positive;negative;zero;text. When any
  //  section carries a [>n]/[<n]/[=n] condition we evaluate those first and
  //  only fall back to the sign-based default when no condition matches.
  const sections = splitSections(pattern);

  let active: string | null = null;
  let useAbs = false;
  // First pass: try condition-bearing sections.
  for (const sec of sections) {
    const cond = parseCondition(sec);
    if (!cond) continue;
    if (cond.test(value)) {
      active = cond.body;
      // Condition matched → caller already wrote the comparator into the
      //  literal so we should NOT show a leading minus from `value` itself.
      useAbs = value < 0;
      break;
    }
  }
  if (active === null) {
    // Second pass: classic sign-based selection over sections that carry
    //  no explicit condition.
    const plain = sections.map((s) => (parseCondition(s) ? null : s));
    const pos = plain[0] ?? null;
    const neg = plain[1] ?? null;
    const zero = plain[2] ?? null;
    if (value < 0 && neg) {
      active = neg;
      useAbs = true;
    } else if (value === 0 && zero) {
      active = zero;
    } else {
      active = pos ?? sections[0] ?? '';
    }
  }

  // Strip style-only directives. Quoted literals must stay quoted until the
  // numeric/date renderer decides whether a token is active or literal.
  active = stripFormatDirectives(active);

  // If the section contains date/time tokens, render as date.
  if (/y|m|d|h|s/.test(stripLiterals(active))) {
    return renderDateTimePattern(value, active, locale);
  }

  return renderNumericPattern(useAbs ? Math.abs(value) : value, active, locale);
}

function formatSpecialPattern(value: number, pattern: string): string {
  const sections = splitSections(pattern);
  const active =
    sections.find((sec) => {
      const cond = parseCondition(sec);
      return cond ? cond.test(value) : false;
    }) ??
    (value < 0
      ? (sections[1] ?? sections[0])
      : value === 0
        ? (sections[2] ?? sections[0])
        : sections[0]) ??
    '';
  const withoutCondition = parseCondition(active)?.body ?? active;
  const body = normalizeFormatSection(withoutCondition);
  const digits = String(Math.trunc(Math.abs(value)));
  let cursor = digits.length - 1;
  let out = '';
  for (let i = body.length - 1; i >= 0; i -= 1) {
    const ch = body[i] ?? '';
    if (ch === '0' || ch === '#' || ch === '?') {
      if (cursor >= 0) {
        out = digits[cursor] + out;
        cursor -= 1;
      } else if (ch === '0') {
        out = `0${out}`;
      } else if (ch === '?') {
        out = ` ${out}`;
      }
    } else {
      out = ch + out;
    }
  }
  if (cursor >= 0) out = digits.slice(0, cursor + 1) + out;
  return value < 0 ? `-${out}` : out;
}

function normalizeFormatSection(section: string): string {
  return stripFormatDirectives(section).replace(/"([^"]*)"/g, '$1');
}

function stripFormatDirectives(section: string): string {
  return (
    section
      // Locale/currency tags: [$¥-411]#,##0 → ¥#,##0; [$-ja-JP] is locale-only.
      .replace(/\[\$([^\]-]+)(?:-[^\]]+)?\]/g, '$1')
      .replace(/\[\$-[^\]]+\]/g, '')
      // Color tags are a style concern; the formatter returns text only.
      .replace(/\[(?:Red|Green|Blue|Black|White|Yellow|Magenta|Cyan|Color\d+)\]/gi, '')
      // Alignment/fill directives. `_x` reserves one char width; `*x`
      // repeats a fill char. Canvas text output should not show either.
      .replace(/_.|\\ /g, '')
      .replace(/\*./g, '')
  );
}

/** Parse a leading condition tag like `[>100]"big"0` into its predicate and
 *  the remaining body. Returns null when the section has no condition. */
function parseCondition(section: string): { test: (n: number) => boolean; body: string } | null {
  const m = section.match(/^\s*\[(>=|<=|<>|=|>|<)\s*(-?\d+(?:\.\d+)?)\s*\](.*)$/s);
  if (!m) return null;
  const op = m[1] ?? '=';
  const target = Number.parseFloat(m[2] ?? '0');
  const body = m[3] ?? '';
  const test = (n: number): boolean => {
    switch (op) {
      case '>':
        return n > target;
      case '<':
        return n < target;
      case '>=':
        return n >= target;
      case '<=':
        return n <= target;
      case '<>':
        return n !== target;
      default:
        return n === target;
    }
  };
  return { test, body };
}

/** Split a format string on `;` that's not inside `"..."` or `[...]`. */
function splitSections(s: string): string[] {
  const out: string[] = [];
  let buf = '';
  let inQuote = false;
  let inBracket = false;
  for (let i = 0; i < s.length; i += 1) {
    const ch = s[i];
    if (ch === '\\' && i + 1 < s.length) {
      buf += ch + s[i + 1];
      i += 1;
      continue;
    }
    if (!inBracket && ch === '"') {
      inQuote = !inQuote;
      buf += ch;
      continue;
    }
    if (!inQuote && ch === '[') {
      inBracket = true;
      buf += ch;
      continue;
    }
    if (inBracket && ch === ']') {
      inBracket = false;
      buf += ch;
      continue;
    }
    if (!inQuote && !inBracket && ch === ';') {
      out.push(buf);
      buf = '';
      continue;
    }
    buf += ch;
  }
  out.push(buf);
  return out;
}

/** Strip quoted literals and escapes so the remaining string can be probed
 *  for unescaped tokens (e.g. detecting `m` as a month token). */
function stripLiterals(s: string): string {
  return s.replace(/"[^"]*"/g, '').replace(/\\./g, '');
}

function hasUnquotedPercent(s: string): boolean {
  let inQuote = false;
  let inBracket = false;
  for (let i = 0; i < s.length; i += 1) {
    const ch = s[i];
    if (ch === '\\' && i + 1 < s.length) {
      i += 1;
      continue;
    }
    if (!inBracket && ch === '"') {
      inQuote = !inQuote;
      continue;
    }
    if (!inQuote && ch === '[') {
      inBracket = true;
      continue;
    }
    if (inBracket && ch === ']') {
      inBracket = false;
      continue;
    }
    if (!inQuote && !inBracket && ch === '%') return true;
  }
  return false;
}

function renderDateTimePattern(serial: number, pattern: string, locale: string): string {
  const d = spreadsheetSerialToDate(serial);
  const yyyy = d.getUTCFullYear();
  const mm = d.getUTCMonth() + 1;
  const dd = d.getUTCDate();
  const dow = d.getUTCDay();
  const hh = d.getUTCHours();
  const mi = d.getUTCMinutes();
  const ss = d.getUTCSeconds();
  // Elapsed-time tokens ([h]/[m]/[s]) accumulate past the 24h/60m wrap and are
  // measured from the serial value directly rather than the wall-clock fields.
  const totalSeconds = Math.round(serial * 86_400);
  const elapsedHours = Math.floor(totalSeconds / 3600);
  const elapsedMinutes = Math.floor(totalSeconds / 60);
  const elapsedSeconds = totalSeconds;
  const has12h = /a\/?p|am\/pm/i.test(pattern);
  const hh12 = ((hh + 11) % 12) + 1;
  const ampm = hh < 12 ? 'AM' : 'PM';
  const DAYS_LONG = Array.from({ length: 7 }, (_, i) =>
    new Intl.DateTimeFormat(locale, { weekday: 'long', timeZone: 'UTC' }).format(
      new Date(Date.UTC(2023, 0, i + 1)),
    ),
  );
  const DAYS_SHORT = Array.from({ length: 7 }, (_, i) =>
    new Intl.DateTimeFormat(locale, { weekday: 'short', timeZone: 'UTC' }).format(
      new Date(Date.UTC(2023, 0, i + 1)),
    ),
  );
  const MONTHS_LONG = Array.from({ length: 12 }, (_, i) =>
    new Intl.DateTimeFormat(locale, { month: 'long', timeZone: 'UTC' }).format(
      new Date(Date.UTC(2023, i, 1)),
    ),
  );
  const MONTHS_SHORT = Array.from({ length: 12 }, (_, i) =>
    new Intl.DateTimeFormat(locale, { month: 'short', timeZone: 'UTC' }).format(
      new Date(Date.UTC(2023, i, 1)),
    ),
  );
  let out = '';
  let prevWasH = false;
  for (let i = 0; i < pattern.length; ) {
    const rest = pattern.slice(i);
    // Quoted literal — emit without the quotes.
    if (pattern[i] === '"') {
      const end = pattern.indexOf('"', i + 1);
      if (end < 0) {
        out += pattern.slice(i + 1);
        break;
      }
      out += pattern.slice(i + 1, end);
      i = end + 1;
      continue;
    }
    if (pattern[i] === '\\' && i + 1 < pattern.length) {
      out += pattern[i + 1];
      i += 2;
      continue;
    }
    // Elapsed-time tokens in square brackets: [h]/[hh] total hours, [m]/[mm]
    // total minutes, [s]/[ss] total seconds. They do not wrap at 24h/60.
    if (pattern[i] === '[') {
      const close = pattern.indexOf(']', i + 1);
      if (close > i) {
        const inner = pattern.slice(i + 1, close);
        if (/^h+$/i.test(inner)) {
          out += String(elapsedHours).padStart(inner.length, '0');
          i = close + 1;
          prevWasH = true;
          continue;
        }
        if (/^m+$/i.test(inner)) {
          out += String(elapsedMinutes).padStart(inner.length, '0');
          i = close + 1;
          prevWasH = false;
          continue;
        }
        if (/^s+$/i.test(inner)) {
          out += String(elapsedSeconds).padStart(inner.length, '0');
          i = close + 1;
          prevWasH = false;
          continue;
        }
      }
    }
    // Token matching, longest first. After hours, an "m"/"mm" token is
    //  minutes, not month — track prevWasH to disambiguate.
    let tok = '';
    if (rest.startsWith('yyyy')) tok = 'yyyy';
    else if (rest.startsWith('yy')) tok = 'yy';
    else if (rest.startsWith('mmmmm')) tok = 'mmmmm';
    else if (rest.startsWith('mmmm')) tok = 'mmmm';
    else if (rest.startsWith('mmm')) tok = 'mmm';
    else if (rest.startsWith('mm')) tok = 'mm';
    else if (rest.startsWith('dddd')) tok = 'dddd';
    else if (rest.startsWith('ddd')) tok = 'ddd';
    else if (rest.startsWith('dd')) tok = 'dd';
    else if (rest.startsWith('hh')) tok = 'hh';
    else if (rest.startsWith('ss')) tok = 'ss';
    else if (/^am\/pm/i.test(rest)) tok = 'am/pm';
    else if (/^a\/p/i.test(rest)) tok = 'a/p';
    else if (rest[0] === 'm') tok = 'm';
    else if (rest[0] === 'd') tok = 'd';
    else if (rest[0] === 'h') tok = 'h';
    else if (rest[0] === 's') tok = 's';

    if (!tok) {
      // Preserve the hour context across time separators (`:`, spaces): in
      // `hh:mm` the `mm` is minutes even though a literal sits between them.
      out += pattern[i];
      i += 1;
      continue;
    }
    switch (tok) {
      case 'yyyy':
        out += String(yyyy);
        break;
      case 'yy':
        out += String(yyyy).slice(-2);
        break;
      case 'mmmmm':
        out += Array.from(MONTHS_LONG[mm - 1] ?? '')[0] ?? '';
        break;
      case 'mmmm':
        out += MONTHS_LONG[mm - 1] ?? '';
        break;
      case 'mmm':
        out += MONTHS_SHORT[mm - 1] ?? '';
        break;
      case 'mm':
        out += prevWasH ? pad2(mi) : pad2(mm);
        break;
      case 'm':
        out += prevWasH ? String(mi) : String(mm);
        break;
      case 'dddd':
        out += DAYS_LONG[dow] ?? '';
        break;
      case 'ddd':
        out += DAYS_SHORT[dow] ?? '';
        break;
      case 'dd':
        out += pad2(dd);
        break;
      case 'd':
        out += String(dd);
        break;
      case 'hh':
        out += pad2(has12h ? hh12 : hh);
        break;
      case 'h':
        out += String(has12h ? hh12 : hh);
        break;
      case 'ss':
        out += pad2(ss);
        break;
      case 's':
        out += String(ss);
        break;
      case 'am/pm':
        out += /AM\/PM/.test(pattern.slice(i, i + 5)) ? ampm : ampm.toLowerCase();
        break;
      case 'a/p':
        out += hh < 12 ? 'A' : 'P';
        break;
    }
    prevWasH = tok === 'h' || tok === 'hh';
    i += tok.length;
  }
  return out;
}

/** Best rational approximation of `x` in [0,1) with denominator ≤ maxDen. */
function bestFraction(x: number, maxDen: number): [number, number] {
  let bestNum = 0;
  let bestDen = 1;
  let bestErr = Number.POSITIVE_INFINITY;
  for (let d = 1; d <= maxDen; d += 1) {
    const n = Math.round(x * d);
    const err = Math.abs(x - n / d);
    if (err < bestErr - 1e-12) {
      bestErr = err;
      bestNum = n;
      bestDen = d;
    }
    if (err < 1e-12) break;
  }
  return [bestNum, bestDen];
}

/** Render Excel fraction formats (`# ?/?`, `??/??`, `# ?/8`). Returns null when
 *  `pattern` is not a fraction format. */
function renderFractionPattern(value: number, pattern: string): string | null {
  const m = pattern.match(/^(.*?)([#0?]+)?\s*([#0?]+)\/([#0?]+|\d+)(.*?)$/);
  if (!m?.[3] || !m[4]) return null;
  // Reject when there's no `?`/`#`/`0` around the slash beyond a bare literal.
  const [, prefix, intPh, , denSpec, suffix] = m;
  const hasInt = !!intPh && intPh.length > 0;
  const sign = value < 0 ? '-' : '';
  let x = Math.abs(value);
  let whole = 0;
  if (hasInt) {
    whole = Math.floor(x);
    x -= whole;
  }
  let numer: number;
  let denom: number;
  if (/^\d+$/.test(denSpec)) {
    denom = Number.parseInt(denSpec, 10) || 1;
    numer = Math.round(x * denom);
  } else {
    const maxDen = 10 ** denSpec.length - 1;
    [numer, denom] = bestFraction(x, maxDen);
  }
  // Rounded up to a whole unit.
  if (numer === denom) {
    whole += 1;
    numer = 0;
  }
  const pre = normalizeFormatSection(prefix ?? '');
  const suf = normalizeFormatSection(suffix ?? '');
  if (numer === 0) {
    const wholeTxt = hasInt ? String(whole) : '0';
    return `${sign}${pre}${wholeTxt}${suf}`;
  }
  if (!hasInt) numer += whole * denom; // improper fraction
  const wholeTxt = hasInt && whole > 0 ? `${whole} ` : '';
  return `${sign}${pre}${wholeTxt}${numer}/${denom}${suf}`;
}

/** Render scientific / engineering notation (`0.00E+00`, `##0.0E+0`). The number
 *  of integer placeholders in the mantissa sets the exponent step, so `##0.0E+0`
 *  yields engineering notation (exponent stepped in multiples of 3). Returns null
 *  when `pattern` is not a scientific format. */
function renderScientificPattern(value: number, pattern: string, locale: string): string | null {
  const m = pattern.match(/^(.*?)([0#?][0#?,]*(?:\.[0#?]+)?)[eE]([+-]?)([0#?]+)(.*)$/);
  if (!m) return null;
  const prefix = normalizeFormatSection(m[1] ?? '');
  const mantissaPat = m[2] ?? '';
  const expSignSpec = m[3] ?? '';
  const expPat = m[4] ?? '';
  const suffix = normalizeFormatSection(m[5] ?? '');
  const dot = mantissaPat.indexOf('.');
  const intPat = dot >= 0 ? mantissaPat.slice(0, dot) : mantissaPat;
  const fracPat = dot >= 0 ? mantissaPat.slice(dot + 1) : '';
  const step = Math.max(1, (intPat.match(/[0#?]/g) ?? []).length);
  const minIntDigits = (intPat.match(/0/g) ?? []).length;
  const minFracDigits = (fracPat.match(/0/g) ?? []).length;
  const maxFracDigits = (fracPat.match(/[0#?]/g) ?? []).length;
  const sign = value < 0 ? '-' : '';
  const av = Math.abs(value);
  let exp = av === 0 ? 0 : Math.floor(Math.floor(Math.log10(av)) / step) * step;
  let mantissa = av === 0 ? 0 : av / 10 ** exp;
  // Rounding at the requested precision can push the mantissa past the step
  // boundary (e.g. 999.6 → 1000 for a 3-digit step); bump the exponent so the
  // integer part stays within `step` digits.
  if (av !== 0 && Number(mantissa.toFixed(maxFracDigits)) >= 10 ** step) {
    exp += step;
    mantissa = av / 10 ** exp;
  }
  const mantissaText = new Intl.NumberFormat(locale, {
    minimumIntegerDigits: Math.max(1, minIntDigits),
    minimumFractionDigits: minFracDigits,
    maximumFractionDigits: maxFracDigits,
    useGrouping: false,
  }).format(mantissa);
  const expDigits = String(Math.abs(exp)).padStart(expPat.length, '0');
  const expSign = exp < 0 ? '-' : expSignSpec === '+' ? '+' : '';
  return `${sign}${prefix}${mantissaText}E${expSign}${expDigits}${suffix}`;
}

function renderNumericPattern(value: number, pattern: string, locale: string): string {
  // Scientific / engineering notation (`0.00E+00`, `##0.0E+0`) is solved into a
  // mantissa and exponent rather than printing the `E+0` placeholders verbatim.
  if (/[eE][+-]?[0#?]/.test(pattern)) {
    const sci = renderScientificPattern(value, pattern, locale);
    if (sci !== null) return sci;
  }
  // Fraction formats (`# ?/?`) are solved into integer + numerator/denominator
  // rather than printing the placeholders verbatim.
  if (pattern.includes('/')) {
    const frac = renderFractionPattern(value, pattern);
    if (frac !== null) return frac;
  }
  // Detect trailing thousand-scaling commas (e.g. "0,," divides by 1e6).
  let scale = 1;
  let body = pattern;
  // Remove trailing commas after the last digit placeholder for scaling.
  const trailingCommas = body.match(/[0#?](,+)\s*[^0#?]*$/);
  if (trailingCommas) {
    const commas = trailingCommas[1] ?? '';
    scale = 10 ** (3 * commas.length);
    // Remove just those commas from the body.
    const i = body.lastIndexOf(commas);
    if (i >= 0) body = body.slice(0, i) + body.slice(i + commas.length);
  }
  const isPercent = hasUnquotedPercent(body);
  let scaled = value / scale;
  if (isPercent) scaled *= 100;

  // Multi-run integer patterns (phone `000-000-0000`, SSN `000-00-0000`) spread
  // the digits across each placeholder run right-to-left. The single-block path
  // below only fills the first run, so delegate whole-number distribution to the
  // special renderer.
  if (!isPercent && !/\.[#0?]/.test(body)) {
    const runs = body.match(/[#0?][#0?,]*/g) ?? [];
    if (runs.length > 1 && !runs.some((r) => r.includes(','))) {
      return formatSpecialPattern(scaled, body);
    }
  }

  // Find the digit-placeholder block surrounding (and including) the decimal.
  const placeholderMatch = body.match(/[#0?][#0?,]*(?:\.[#0?]+)?|\.[#0?]+/);
  if (!placeholderMatch) return body;

  const block = placeholderMatch[0];
  const dotIndex = block.indexOf('.');
  const intPart = dotIndex >= 0 ? block.slice(0, dotIndex) : block;
  const fracPart = dotIndex >= 0 ? block.slice(dotIndex + 1) : '';
  const grouping = intPart.includes(',');
  const minIntDigits = (intPart.match(/0/g) ?? []).length;
  const minFracDigits = (fracPart.match(/0/g) ?? []).length;
  const maxFracDigits = (fracPart.match(/[0#?]/g) ?? []).length;

  const formatted = new Intl.NumberFormat(locale, {
    minimumIntegerDigits: Math.max(1, minIntDigits),
    minimumFractionDigits: minFracDigits,
    maximumFractionDigits: maxFracDigits,
    useGrouping: grouping,
  }).format(scaled);

  return normalizeFormatSection(body.replace(block, formatted));
}
