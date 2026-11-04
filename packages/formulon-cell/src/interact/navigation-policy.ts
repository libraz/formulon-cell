import { interactionControllerFor } from '../commands/interaction-controller.js';
import { MAX_COL, MAX_ROW } from '../engine/address.js';
import {
  assertPatternSheet,
  compileRangeInput,
  patternOnSheet,
  type RangeInput,
  type RangePattern,
  resolveRangeInput,
} from '../engine/range-input.js';
import { sheetIndexByName } from '../engine/range-resolver.js';
import type { Addr, Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { sameAddr } from '../store/direct-format.js';
import { frozenColEnd, frozenRowEnd } from '../store/freeze-band.js';
import {
  cloneRange,
  rangeContainsAddr,
  rangeContainsRange,
  rangesIntersect,
  sameRange,
} from '../store/selection-geometry.js';
import { clampSplitPaneOrigin } from '../store/split-panes.js';
import type { LayoutSlice, SpreadsheetStore, State } from '../store/store.js';
import { resetTabNavigationSession } from './tab-navigation-session.js';
import { isCellVisible } from './visible-navigation.js';

/** The spreadsheet's physical coordinate limits (zero based, inclusive). */

/**
 * View and keyboard-navigation restrictions for an embedded grid.
 *
 * `range` limits the visible and selectable rectangle on the active sheet
 * while leaving workbook data and calculation untouched; A1 text without a
 * sheet name resolves on the active sheet. `selectable` is an additional
 * allow-list (or predicate) for cursor movement; A1 text without a sheet name
 * there applies to every sheet. It is deliberately kept separate from the
 * interaction policy: this module only decides where navigation may go.
 */
export interface ViewportOptions {
  range?: RangeInput;
  selectable?: readonly RangeInput[] | ((addr: Addr) => boolean);
  tabNavigation?: 'normal' | 'editable';
  tabBoundary?: 'stop' | 'leave';
  autoExpand?: boolean;
}

export interface NavigationPolicyHandle {
  readonly options: ViewportOptions | undefined;
  setOptions(next?: ViewportOptions): void;
  dispose(): void;
}

type NavigationOptions = Readonly<{
  range?: Range;
  selectable?: readonly RangePattern[] | ((addr: Addr) => boolean);
  /** The host's inputs as given, returned by the options getter. */
  rangeInput?: RangeInput;
  selectableInput?: readonly RangeInput[];
  tabNavigation: 'normal' | 'editable';
  tabBoundary: 'stop' | 'leave';
  autoExpand: boolean;
}>;

interface PolicyRecord {
  getWb: () => WorkbookHandle;
  options: NavigationOptions | undefined;
  handle: NavigationPolicyHandle;
}

const registry = new WeakMap<SpreadsheetStore, PolicyRecord>();

/** A bounded scan is enough for an embedded form and prevents a sparse
 * predicate from turning Tab into a billion-cell loop. */
const MAX_TAB_SCAN = 100_000;

const validInteger = (n: number): boolean => Number.isInteger(n) && Number.isFinite(n);

const cloneRangeInput = (input: RangeInput): RangeInput =>
  typeof input === 'string' ? input : cloneRange(input);

const addressInSheet = (addr: Addr): boolean =>
  validInteger(addr.sheet) &&
  addr.sheet >= 0 &&
  validInteger(addr.row) &&
  addr.row >= 0 &&
  addr.row <= MAX_ROW &&
  validInteger(addr.col) &&
  addr.col >= 0 &&
  addr.col <= MAX_COL;

function validateRange(range: Range, sheetCount: number, label: string): void {
  if (!validInteger(range.sheet) || range.sheet < 0 || range.sheet >= sheetCount) {
    throw new Error(`ViewportOptions.${label}.sheet is outside the workbook`);
  }
  if (
    !validInteger(range.r0) ||
    !validInteger(range.c0) ||
    !validInteger(range.r1) ||
    !validInteger(range.c1) ||
    range.r0 < 0 ||
    range.c0 < 0 ||
    range.r1 > MAX_ROW ||
    range.c1 > MAX_COL ||
    range.r0 > range.r1 ||
    range.c0 > range.c1
  ) {
    throw new Error(`ViewportOptions.${label} must be a non-empty in-sheet range`);
  }
}

function assertNoPartialMerge(state: State | undefined, range: Range, label: string): void {
  if (!state) return;
  for (const merge of state.merges.byAnchor.values()) {
    if (!rangesIntersect(merge, range)) continue;
    if (!rangeContainsRange(range, merge)) {
      throw new Error(
        `ViewportOptions.${label} partially exposes merged range ` +
          `${merge.sheet}:${merge.r0}:${merge.c0}-${merge.r1}:${merge.c1}`,
      );
    }
  }
}

/**
 * Validate an options object without registering it. Mount code uses this
 * before replacing a workbook so an invalid range cannot detach a working
 * policy. When `state` is supplied, merged-cell exposure is checked too.
 */
export function validateViewportOptions(
  options: ViewportOptions | undefined,
  wb: WorkbookHandle,
  state?: State,
): void {
  if (!options) return;
  const activeSheet = state?.data.sheetIndex ?? 0;
  if (options.range) {
    const range = resolveRangeInput(options.range, wb, activeSheet, 'ViewportOptions.range');
    validateRange(range, wb.sheetCount, 'range');
    assertNoPartialMerge(state, range, 'range');
    if (state && state.data.sheetIndex !== range.sheet) {
      throw new Error('ViewportOptions.range must target the active sheet');
    }
  }
  if (options.selectable && Array.isArray(options.selectable)) {
    for (const [index, input] of options.selectable.entries()) {
      const label = `selectable[${index}]`;
      const pattern = compileRangeInput(input, `ViewportOptions.${label}`);
      if (pattern.kind === 'range') validateRange(pattern.range, wb.sheetCount, label);
      else assertPatternSheet(pattern, wb, `ViewportOptions.${label}`);
      const onActive = patternOnSheet(pattern, activeSheet, wb);
      if (onActive) assertNoPartialMerge(state, onActive, label);
    }
  }
  if (options.autoExpand === true && options.range) {
    throw new Error('ViewportOptions.autoExpand cannot be used with a fixed range');
  }
  if (
    options.tabNavigation &&
    options.tabNavigation !== 'normal' &&
    options.tabNavigation !== 'editable'
  ) {
    throw new Error("ViewportOptions.tabNavigation must be 'normal' or 'editable'");
  }
  if (options.tabBoundary && options.tabBoundary !== 'stop' && options.tabBoundary !== 'leave') {
    throw new Error("ViewportOptions.tabBoundary must be 'stop' or 'leave'");
  }
  if (
    options.selectable &&
    !Array.isArray(options.selectable) &&
    typeof options.selectable !== 'function'
  ) {
    throw new Error('ViewportOptions.selectable must be a range list or predicate');
  }
}

function normalizeOptions(
  options: ViewportOptions,
  wb: WorkbookHandle,
  state: State,
): NavigationOptions {
  validateViewportOptions(options, wb, state);
  const list = Array.isArray(options.selectable)
    ? (options.selectable as readonly RangeInput[])
    : undefined;
  const selectable = list
    ? list.map((input) => {
        const pattern = compileRangeInput(input, 'ViewportOptions.selectable');
        return pattern.kind === 'range'
          ? { kind: 'range' as const, range: cloneRange(pattern.range) }
          : pattern;
      })
    : (options.selectable as ((addr: Addr) => boolean) | undefined);
  const range = options.range
    ? resolveRangeInput(options.range, wb, state.data.sheetIndex, 'ViewportOptions.range')
    : undefined;
  return Object.freeze({
    ...(range ? { range: cloneRange(range) } : {}),
    ...(options.range ? { rangeInput: cloneRangeInput(options.range) } : {}),
    ...(selectable ? { selectable } : {}),
    ...(list ? { selectableInput: list.map(cloneRangeInput) } : {}),
    tabNavigation: options.tabNavigation ?? 'normal',
    tabBoundary: options.tabBoundary ?? 'stop',
    autoExpand: options.autoExpand ?? false,
  });
}

/** The sheet a viewport range pins: a numeric range's sheet or a qualified
 * text range's sheet. Unqualified text (and an unknown sheet name, which
 * validation rejects) pins none, so the active sheet stays. */
export function viewportRangeSheet(
  options: ViewportOptions | undefined,
  wb: WorkbookHandle,
): number | undefined {
  const range = options?.range;
  if (range === undefined) return undefined;
  const pattern = compileRangeInput(range, 'ViewportOptions.range');
  if (pattern.kind === 'range') return pattern.range.sheet;
  if (pattern.ref.sheetName === null) return undefined;
  const sheet = sheetIndexByName(wb, pattern.ref.sheetName);
  return sheet < 0 ? undefined : sheet;
}

/** Return the registered navigation policy, if this store has one. */
export function navigationPolicyFor(store: SpreadsheetStore): NavigationPolicyHandle | undefined {
  return registry.get(store)?.handle;
}

function recordFor(store: SpreadsheetStore): PolicyRecord | undefined {
  return registry.get(store);
}

/**
 * Attach a navigation policy to a store. A missing options object is an
 * inert handle and leaves all legacy navigation untouched.
 */
export function attachNavigationPolicy(
  store: SpreadsheetStore,
  getWb: () => WorkbookHandle,
  options?: ViewportOptions,
): NavigationPolicyHandle {
  resetTabNavigationSession(store);
  const old = registry.get(store);
  old?.handle.dispose();

  const record = {} as PolicyRecord;
  const handle: NavigationPolicyHandle = {
    get options(): ViewportOptions | undefined {
      return record.options ? clonePublicOptions(record.options) : undefined;
    },
    setOptions(next?: ViewportOptions): void {
      if (!next) {
        resetTabNavigationSession(store);
        record.options = undefined;
        registry.delete(store);
        syncNavigationViewport(store);
        return;
      }
      // Validate and normalize before replacing the current policy.
      const normalized = normalizeOptions(next, getWb(), store.getState());
      resetTabNavigationSession(store);
      record.options = normalized;
      registry.set(store, record);
      syncNavigationViewport(store);
    },
    dispose(): void {
      if (registry.get(store) !== record) return;
      resetTabNavigationSession(store);
      registry.delete(store);
      record.options = undefined;
      syncNavigationViewport(store);
    },
  };
  record.getWb = getWb;
  record.options = undefined;
  record.handle = handle;
  if (options) {
    record.options = normalizeOptions(options, getWb(), store.getState());
    registry.set(store, record);
    syncNavigationViewport(store);
  }
  return handle;
}

function clonePublicOptions(options: NavigationOptions): ViewportOptions {
  return {
    ...(options.rangeInput ? { range: cloneRangeInput(options.rangeInput) } : {}),
    ...(options.selectableInput
      ? { selectable: options.selectableInput.map(cloneRangeInput) }
      : typeof options.selectable === 'function'
        ? { selectable: options.selectable }
        : {}),
    tabNavigation: options.tabNavigation,
    tabBoundary: options.tabBoundary,
    autoExpand: options.autoExpand,
  };
}

function activeRange(options?: NavigationOptions): Range | undefined {
  const range = options?.range;
  return range;
}

/** The fixed rectangle currently affecting the active sheet, if any. */
export function navigationBoundsFor(store: SpreadsheetStore): Range | undefined {
  const options = recordFor(store)?.options;
  const range = activeRange(options);
  return range ? cloneRange(range) : undefined;
}

/** The selectable list's rectangles on `sheet`, or null when no entry
 * applies to that sheet (a list only constrains the sheets it reaches). */
function selectableOnSheet(
  options: NavigationOptions,
  sheet: number,
  wb: WorkbookHandle,
): Range[] | null {
  if (!Array.isArray(options.selectable)) return null;
  const ranges: Range[] = [];
  for (const pattern of options.selectable as readonly RangePattern[]) {
    const range = patternOnSheet(pattern, sheet, wb);
    if (range) ranges.push(cloneRange(range));
  }
  return ranges.length > 0 ? ranges : null;
}

function selectableRangesFor(store: SpreadsheetStore, options: NavigationOptions): Range[] {
  const record = recordFor(store);
  if (!record) return [];
  return selectableOnSheet(options, store.getState().data.sheetIndex, record.getWb()) ?? [];
}

/** Bounding rectangle used by full-row/full-column/select-all operations. */
export function navigationSelectionBoundsFor(store: SpreadsheetStore): Range | undefined {
  const options = recordFor(store)?.options;
  const fixed = activeRange(options);
  if (fixed) return cloneRange(fixed);
  const ranges = options ? selectableRangesFor(store, options) : [];
  if (ranges.length === 0) return undefined;
  return {
    sheet: ranges[0]?.sheet ?? store.getState().data.sheetIndex,
    r0: Math.min(...ranges.map((r) => r.r0)),
    c0: Math.min(...ranges.map((r) => r.c0)),
    r1: Math.max(...ranges.map((r) => r.r1)),
    c1: Math.max(...ranges.map((r) => r.c1)),
  };
}

function selectorAllows(addr: Addr, options: NavigationOptions, wb: WorkbookHandle): boolean {
  if (!options.selectable) return true;
  if (Array.isArray(options.selectable)) {
    // A selectable list only constrains the sheet(s) it mentions. This lets a
    // host configure one embedded sheet without making other sheets unusable.
    const ranges = selectableOnSheet(options, addr.sheet, wb);
    return !ranges || ranges.some((range) => rangeContainsAddr(range, addr));
  }
  const predicate = options.selectable;
  return typeof predicate === 'function' ? predicate(addr) : true;
}

function firstAllowedAddress(store: SpreadsheetStore, preferred: Addr): Addr | null {
  const clamped = clampNavigationAddr(store, preferred);
  if (clamped) return clamped;
  const options = recordFor(store)?.options;
  if (!options) return null;
  const bound = activeRange(options) ?? navigationSelectionBoundsFor(store);
  const sheet =
    bound?.sheet ??
    (preferred.sheet === store.getState().data.sheetIndex ? preferred.sheet : undefined);
  if (sheet === undefined) return null;
  const search: Range = bound
    ? { ...bound, sheet }
    : { sheet, r0: 0, c0: 0, r1: MAX_ROW, c1: MAX_COL };
  let scanned = 0;
  for (let row = search.r0; row <= search.r1 && scanned < MAX_TAB_SCAN; row += 1) {
    for (let col = search.c0; col <= search.c1 && scanned < MAX_TAB_SCAN; col += 1) {
      scanned += 1;
      const candidate = { sheet, row, col };
      if (isNavigationAddrAllowed(store, candidate)) return candidate;
    }
  }
  return null;
}

function syncNavigationSelection(store: SpreadsheetStore): void {
  if (!recordFor(store)?.options) return;
  const state = store.getState();
  const active = firstAllowedAddress(store, state.selection.active);
  if (!active) return;
  const anchor = firstAllowedAddress(store, state.selection.anchor) ?? active;
  const currentRange = clampNavigationRange(store, state.selection.range);
  const range =
    currentRange && rangeContainsAddr(currentRange, active)
      ? currentRange
      : { sheet: active.sheet, r0: active.row, c0: active.col, r1: active.row, c1: active.col };
  const extraRanges = (state.selection.extraRanges ?? [])
    .map((extra) => clampNavigationRange(store, extra))
    .filter((extra): extra is Range => extra !== null);
  const oldExtras = state.selection.extraRanges ?? [];
  const extrasChanged =
    oldExtras.length !== extraRanges.length ||
    oldExtras.some((extra, index) => !extraRanges[index] || !sameRange(extra, extraRanges[index]));
  if (
    sameAddr(state.selection.active, active) &&
    sameAddr(state.selection.anchor, anchor) &&
    sameRange(state.selection.range, range) &&
    !extrasChanged
  ) {
    return;
  }
  store.setState((current) => ({
    ...current,
    selection: { ...current.selection, active, anchor, range, extraRanges },
  }));
}

/** Whether an address is a legal navigation target. This is not a mutation
 * authorization check; callers still route writes through the controller. */
export function isNavigationAddrAllowed(store: SpreadsheetStore, addr: Addr): boolean {
  if (!addressInSheet(addr)) return false;
  const record = recordFor(store);
  const options = record?.options;
  if (!record || !options) return true;
  if (options.range) {
    if (addr.sheet !== options.range.sheet || !rangeContainsAddr(options.range, addr)) return false;
  }
  return selectorAllows(addr, options, record.getWb());
}

/** Clamp an address to a fixed range. Predicate/list restrictions reject an
 * address they cannot safely map to; they are intentionally never coerced to
 * an arbitrary last cell. */
export function clampNavigationAddr(store: SpreadsheetStore, addr: Addr): Addr | null {
  if (!addressInSheet(addr)) return null;
  const record = recordFor(store);
  const options = record?.options;
  if (!record || !options) return { sheet: addr.sheet, row: addr.row, col: addr.col };
  let next: Addr = { sheet: addr.sheet, row: addr.row, col: addr.col };
  if (options.range) {
    if (options.range.sheet !== addr.sheet) return null;
    next = {
      sheet: addr.sheet,
      row: Math.max(options.range.r0, Math.min(options.range.r1, addr.row)),
      col: Math.max(options.range.c0, Math.min(options.range.c1, addr.col)),
    };
  }
  return selectorAllows(next, options, record.getWb()) ? next : null;
}

/** Clamp a selection range to the active fixed rectangle. */
export function clampNavigationRange(store: SpreadsheetStore, input: Range): Range | null {
  const record = recordFor(store);
  const options = record?.options;
  if (!record || !options) return cloneRange(input);
  const bound = options.range;
  if (bound && bound.sheet !== input.sheet) return null;
  const range = bound
    ? {
        sheet: input.sheet,
        r0: Math.max(input.r0, bound.r0),
        c0: Math.max(input.c0, bound.c0),
        r1: Math.min(input.r1, bound.r1),
        c1: Math.min(input.c1, bound.c1),
      }
    : cloneRange(input);
  if (range.r0 > range.r1 || range.c0 > range.c1) return null;
  // A fixed selectable list is a range allow-list. A selection crossing two
  // disjoint entries remains ambiguous, so only accept a fully covered entry.
  if (Array.isArray(options.selectable)) {
    const ranges = selectableOnSheet(options, range.sheet, record.getWb());
    if (ranges && !ranges.some((r) => rangeContainsRange(r, range))) return null;
  }
  if (typeof options.selectable === 'function') {
    const area = (range.r1 - range.r0 + 1) * (range.c1 - range.c0 + 1);
    if (!Number.isSafeInteger(area) || area > MAX_TAB_SCAN) return null;
    for (let row = range.r0; row <= range.r1; row += 1) {
      for (let col = range.c0; col <= range.c1; col += 1) {
        if (!options.selectable({ sheet: range.sheet, row, col })) return null;
      }
    }
  }
  return range;
}

/** Keep the store's viewport starts inside the registered range. */
export function syncNavigationViewport(store: SpreadsheetStore): void {
  syncNavigationSelection(store);
  const bounds = navigationBoundsFor(store);
  store.setState((state) => {
    const current = state.viewport.navigationRange;
    const sameBounds =
      current?.sheet === bounds?.sheet &&
      current?.r0 === bounds?.r0 &&
      current?.c0 === bounds?.c0 &&
      current?.r1 === bounds?.r1 &&
      current?.c1 === bounds?.c1;
    const minRow = Math.max(frozenRowEnd(state.layout), bounds?.r0 ?? 0);
    const minCol = Math.max(frozenColEnd(state.layout), bounds?.c0 ?? 0);
    const maxRow = bounds
      ? Math.max(minRow, bounds.r1 + 1 - state.viewport.rowCount)
      : Math.max(minRow, MAX_ROW + 1 - state.viewport.rowCount);
    const maxCol = bounds
      ? Math.max(minCol, bounds.c1 + 1 - state.viewport.colCount)
      : Math.max(minCol, MAX_COL + 1 - state.viewport.colCount);
    const rowStart = Math.min(maxRow, Math.max(minRow, state.viewport.rowStart));
    const colStart = Math.min(maxCol, Math.max(minCol, state.viewport.colStart));
    if (
      sameBounds &&
      rowStart === state.viewport.rowStart &&
      colStart === state.viewport.colStart
    ) {
      return state;
    }
    return {
      ...state,
      viewport: {
        ...state.viewport,
        ...(bounds ? { navigationRange: cloneRange(bounds) } : { navigationRange: undefined }),
        rowStart,
        colStart,
      },
    };
  });
}

function nextAddressInRange(range: Range, addr: Addr, reverse: boolean): Addr | null {
  let row = addr.row;
  let col = addr.col;
  if (reverse) {
    col -= 1;
    if (col < range.c0) {
      col = range.c1;
      row -= 1;
    }
    if (row < range.r0) return null;
  } else {
    col += 1;
    if (col > range.c1) {
      col = range.c0;
      row += 1;
    }
    if (row > range.r1) return null;
  }
  return { sheet: range.sheet, row, col };
}

function firstAddress(range: Range, reverse: boolean): Addr {
  return reverse
    ? { sheet: range.sheet, row: range.r1, col: range.c1 }
    : { sheet: range.sheet, row: range.r0, col: range.c0 };
}

/**
 * Find the next Tab stop in row-major order. Editable traversal delegates
 * every candidate to the shared interaction controller; this helper never
 * grants edit permission by itself.
 */
export function nextTabStop(store: SpreadsheetStore, addr: Addr, reverse: boolean): Addr | null {
  const record = recordFor(store);
  const options = record?.options;
  if (!options) return null;
  const bound = activeRange(options) ?? navigationSelectionBoundsFor(store);
  if (!bound) return null;
  let candidate: Addr | null = rangeContainsAddr(bound, addr) ? addr : firstAddress(bound, reverse);
  if (rangeContainsAddr(bound, addr)) {
    candidate = nextAddressInRange(bound, addr, reverse) ?? null;
  }
  for (let i = 0; candidate && i < MAX_TAB_SCAN; i += 1) {
    if (
      isCellVisible(store.getState().layout, candidate.row, candidate.col) &&
      isNavigationAddrAllowed(store, candidate)
    ) {
      if (options.tabNavigation !== 'editable') return candidate;
      const controller = interactionControllerFor(store);
      const permission = controller?.canExecute({
        operation: 'valueEdit',
        origin: 'keyboard',
        effects: [{ kind: 'cells', cells: [candidate] }],
      });
      if (permission?.allowed === true) return candidate;
    }
    candidate = nextAddressInRange(bound, candidate, reverse);
  }
  return null;
}

/** A split's top/left pane scrolled back to show `addr` above or left of it. */
function splitPanesToward(
  layout: LayoutSlice,
  addr: Addr,
  rowStart: number,
  colStart: number,
): LayoutSlice {
  if (layout.splitPanes !== true) return layout;
  const originRow = layout.freezeOriginRow ?? 0;
  const originCol = layout.freezeOriginCol ?? 0;
  const freezeOriginRow =
    layout.freezeRows > 0 && addr.row < originRow
      ? clampSplitPaneOrigin(layout.freezeRows, Math.max(rowStart, frozenRowEnd(layout)), addr.row)
      : originRow;
  const freezeOriginCol =
    layout.freezeCols > 0 && addr.col < originCol
      ? clampSplitPaneOrigin(layout.freezeCols, Math.max(colStart, frozenColEnd(layout)), addr.col)
      : originCol;
  if (freezeOriginRow === originRow && freezeOriginCol === originCol) return layout;
  return { ...layout, freezeOriginRow, freezeOriginCol };
}

/** Scroll the viewport the least amount that brings `addr` into view. */
export function scrollActiveIntoView(store: SpreadsheetStore, addr: Addr): void {
  store.setState((s) => {
    if (addr.sheet !== s.data.sheetIndex) return s;
    const rowStartMin = frozenRowEnd(s.layout);
    const colStartMin = frozenColEnd(s.layout);
    const maxRowStart = Math.max(rowStartMin, MAX_ROW + 1 - s.viewport.rowCount);
    const maxColStart = Math.max(colStartMin, MAX_COL + 1 - s.viewport.colCount);
    let rowStart = s.viewport.rowStart;
    let colStart = s.viewport.colStart;
    const visibleRowStart = Math.max(rowStart, frozenRowEnd(s.layout));
    const visibleColStart = Math.max(colStart, frozenColEnd(s.layout));
    const visibleRowEnd = rowStart + s.viewport.rowCount - 1;
    const visibleColEnd = colStart + s.viewport.colCount - 1;
    if (addr.row >= frozenRowEnd(s.layout)) {
      if (addr.row < visibleRowStart) rowStart = addr.row;
      else if (addr.row > visibleRowEnd) rowStart = addr.row - s.viewport.rowCount + 1;
    }
    if (addr.col >= frozenColEnd(s.layout)) {
      if (addr.col < visibleColStart) colStart = addr.col;
      else if (addr.col > visibleColEnd) colStart = addr.col - s.viewport.colCount + 1;
    }
    rowStart = Math.min(maxRowStart, Math.max(rowStartMin, rowStart));
    colStart = Math.min(maxColStart, Math.max(colStartMin, colStart));
    const layout = splitPanesToward(s.layout, addr, rowStart, colStart);
    if (
      layout === s.layout &&
      rowStart === s.viewport.rowStart &&
      colStart === s.viewport.colStart
    ) {
      return s;
    }
    return { ...s, layout, viewport: { ...s.viewport, rowStart, colStart } };
  });
}
