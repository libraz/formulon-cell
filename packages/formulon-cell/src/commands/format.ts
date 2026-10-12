import { addrKey, MAX_COL, MAX_ROW, parseAddrKey } from '../engine/address.js';
import {
  type AxisFormatBand,
  type AxisFormatPatchResolver,
  prepareAxisFormatOverrideWithWorkbook,
} from '../engine/axis-format-sync.js';
import { cellFormatFromXf, effectiveCellFormatWithPending } from '../engine/cell-format-sync.js';
import { formatHistoryBridgeFor } from '../engine/format-history-bridge.js';
import type { Addr, Range } from '../engine/types.js';
import { isWholeSheetSelection } from '../engine/whole-sheet-font-sync.js';
import { currencySymbolForLocale } from '../format/locale.js';
import { formatWithPending, sameAddr } from '../store/pending-format.js';
import {
  isWholeColumnRange,
  isWholeRowRange,
  rangeContainsAddr,
  rangesIntersect,
  selectionCoversRange,
  subtractRange,
} from '../store/selection-geometry.js';
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
import { planAxisSelection } from './axis-selection-plan.js';
import { interactionControllerFor } from './interaction-controller.js';
import type { InteractionOperation, InteractionOrigin } from './interaction-policy.js';
import { gateProtection, isCellWritable, isSheetProtected } from './protection.js';

export { formatNumber } from '../format/number-format.js';

type ToggleKey = 'bold' | 'italic' | 'underline' | 'strike';

const MAX_MATERIALIZED_FORMAT_CELLS = 100_000;

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
  range.r1 <= MAX_ROW &&
  range.c1 <= MAX_COL;

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

const selectionPlan = (
  state: State,
): {
  ranges: Range[];
  cells: Addr[];
  exceptionCells: Addr[];
  axisBands: AxisFormatBand[];
} | null => {
  const ranges = selectionRanges(state);
  if (!ranges || !mergeCoverageValid(state, ranges)) return null;
  const normalized = disjointRanges(ranges);
  const geometry = planAxisSelection(ranges, MAX_MATERIALIZED_FORMAT_CELLS);
  if (!geometry) return null;
  const cells = materializeSelection(geometry.ordinaryRanges);
  const exceptionCells = materializeSelection(geometry.exceptionRanges);
  if (
    !cells ||
    !exceptionCells ||
    cells.length + exceptionCells.length > MAX_MATERIALIZED_FORMAT_CELLS
  )
    return null;
  const sparse = new Map<string, Addr>();
  for (const addr of cells) sparse.set(addrKey(addr), addr);
  const exceptions = new Set(exceptionCells.map(addrKey));
  // Axis selections modify the compact native style and only materialise
  // already-authored cells in the store. Blank/future cells remain implicit.
  if (geometry.axisBands.length > 0) {
    for (const key of state.format.formats.keys()) {
      const addr = parseAddrKey(key);
      if (!addr || addr.sheet !== normalized[0]?.sheet) continue;
      if (
        !exceptions.has(key) &&
        geometry.axisBands.some(
          (band) =>
            (band.axis === 'column' && addr.col >= band.first && addr.col <= band.last) ||
            (band.axis === 'row' && addr.row >= band.first && addr.row <= band.last),
        )
      )
        sparse.set(key, addr);
    }
  }
  if (sparse.size + exceptions.size > MAX_MATERIALIZED_FORMAT_CELLS) return null;
  return {
    ranges: normalized,
    cells: [...sparse.values()],
    exceptionCells,
    axisBands: geometry.axisBands.map((band) => ({ ...band })),
  };
};

export interface SelectionFormatPlan {
  readonly ranges: readonly Range[];
  readonly cells: readonly Addr[];
  readonly exceptionCells: readonly Addr[];
  readonly axisBands: readonly AxisFormatBand[];
}

/** Plan the complete primary-plus-extra selection for dialog-style actions. */
export function planSelectionFormat(state: State): SelectionFormatPlan | null {
  const plan = selectionPlan(state);
  if (!plan) return null;
  return {
    ranges: plan.ranges.map((range) => ({ ...range })),
    cells: plan.cells.map((addr) => ({ ...addr })),
    exceptionCells: plan.exceptionCells.map((addr) => ({ ...addr })),
    axisBands: plan.axisBands.map((band) => ({ ...band })),
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

const axisSafeFormatKeys = new Set<keyof CellFormat>([
  'numFmt',
  'bold',
  'italic',
  'underline',
  'strike',
  'align',
  'vAlign',
  'wrap',
  'justifyLastLine',
  'shrinkToFit',
  'indent',
  'rotation',
  'textDirection',
  'fontVertAlign',
  'borders',
  'color',
  'fill',
  'fillPattern',
  'fillPatternColor',
  'fontFamily',
  'fontSize',
  'locked',
  'formulaHidden',
  // Named styles carry visual fields in the payload. The native row/column
  // XF cannot retain the style id itself, but sparse authored cells still do;
  // allowing the route keeps Cell Styles on axis selections from silently
  // falling back to MAX_ROW/MAX_COL materialization.
  'cellStyle',
]);

const axisPatchSupported = (patch: Partial<CellFormat>): boolean =>
  Object.keys(patch).every((key) => axisSafeFormatKeys.has(key as keyof CellFormat));

/** Apply one style patch to row/column bands plus the bounded sparse cells in
 *  the same selection. The axis portion remains a compact native override;
 *  ordinary cells retain the existing CellFormat representation. */
const applyAxisFormatPatch = (
  state: State,
  store: SpreadsheetStore,
  plan: {
    ranges: readonly Range[];
    cells: readonly Addr[];
    exceptionCells: readonly Addr[];
    axisBands: readonly AxisFormatBand[];
  },
  patch: Partial<CellFormat>,
  options: SelectionFormatOptions = {},
  resolvePatch?: AxisFormatPatchResolver,
): boolean => {
  if (plan.axisBands.length === 0 || !axisPatchSupported(patch)) return false;
  const sheet = plan.ranges[0]?.sheet;
  if (sheet === undefined || isSheetProtected(state, sheet)) return false;
  if (!authorizeSelection(store, plan.ranges, plan.cells, options)) return false;
  const controller = interactionControllerFor(store);
  if (controller?.policy?.editable !== undefined) return false;
  const bridge = formatHistoryBridgeFor(store);
  const workbook = bridge?.workbook;
  if (
    workbook?.capabilities.cellFormatting !== true ||
    workbook?.capabilities.sheetFontStyles !== true
  )
    return false;
  const budgetKeys = new Set([...plan.cells, ...plan.exceptionCells].map(addrKey));
  const preparation = prepareAxisFormatOverrideWithWorkbook(
    state,
    store,
    plan.ranges,
    plan.axisBands,
    patch,
    workbook,
    plan.exceptionCells,
    resolvePatch,
    budgetKeys.size,
    budgetKeys,
  );
  if (!preparation) return false;
  const writable = plan.cells.filter((addr) => isCellWritable(state, addr));
  if (writable.length > MAX_MATERIALIZED_FORMAT_CELLS) return false;
  store.setState((current) => {
    const formats = new Map(current.format.formats);
    for (const addr of writable) {
      const key = addrKey(addr);
      const resolvedPatch = resolvePatch
        ? (() => {
            const effective = workbook.getEffectiveStyle(addr.sheet, addr.row, addr.col);
            const base =
              effective?.xfIndex ?? workbook.getCellXfIndex(addr.sheet, addr.row, addr.col);
            return base === null || base === undefined
              ? patch
              : (resolvePatch(workbook, base, addr) ?? patch);
          })()
        : patch;
      const next = mergeCellFormat(formats.get(key), resolvedPatch);
      if (next) formats.set(key, next);
      else formats.delete(key);
    }
    const axisOverrides = new Map(current.format.sheetAxisOverrides ?? []);
    axisOverrides.set(sheet, preparation.override);
    // A compact axis override becomes the sole authority after a user starts
    // editing a selected row/column. Keeping the old worksheet-wide entry
    // would make the two native sync paths race on every publication.
    const sheetFonts = new Map(current.format.sheetFontOverrides ?? []);
    sheetFonts.delete(sheet);
    return {
      ...current,
      format: {
        ...current.format,
        formats,
        sheetAxisOverrides: axisOverrides,
        sheetFontOverrides: sheetFonts,
      },
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
  if (plan.axisBands.length > 0) {
    return applyAxisFormatPatch(state, store, plan, patch ?? {}, options);
  }
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
    const addr = parseAddrKey(key);
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
  const wholeRanges = selectionRanges(state);
  const axisPlan = selectionPlan(state);
  const axisSheet = axisPlan?.ranges[0]?.sheet;
  const axisWorkbook = formatHistoryBridgeFor(store)?.workbook;
  const axisNativeAvailable =
    axisWorkbook?.capabilities.cellFormatting === true &&
    axisWorkbook.capabilities.sheetFontStyles === true;
  // Clearing a row/column must also clear its compact native XF. Reusing the
  // regular axis patch keeps the operation sparse and lets the native flush
  // journal the same before/after state as toolbar/dialog formatting. A
  // worksheet-wide selection continues through the legacy whole-sheet path
  // unless it already has an axis override, where that override is the sole
  // native authority.
  const axisClear =
    axisPlan &&
    axisPlan.axisBands.length > 0 &&
    axisSheet !== undefined &&
    axisNativeAvailable &&
    (!wholeRanges ||
      !isWholeSheetSelection(wholeRanges) ||
      state.format.sheetAxisOverrides?.has(axisSheet));
  if (axisClear) {
    const clearPatch = Object.fromEntries(
      visualFormatKeys.map((key) => [key, undefined]),
    ) as Partial<CellFormat>;
    if (!applyAxisFormatPatch(state, store, axisPlan, clearPatch, options)) return false;
    const selected = new Set(axisPlan.cells.map(addrKey));
    store.setState((current) => {
      const formats = new Map(current.format.formats);
      for (const key of selected) {
        const addr = parseAddrKey(key);
        if (!addr || !isCellWritable(current, addr)) continue;
        if (!visualOnly) formats.delete(key);
        else {
          const previous = formats.get(key);
          if (!previous) continue;
          const next = stripVisualFormat(previous);
          if (next) formats.set(key, next);
          else formats.delete(key);
        }
      }
      return {
        ...current,
        format: { ...current.format, formats },
        ui: { ...current.ui, pendingFormat: null },
      };
    });
    return true;
  }
  if (wholeRanges && isWholeSheetSelection(wholeRanges)) {
    const sheet = wholeRanges[0]?.sheet;
    if (sheet === undefined || isSheetProtected(state, sheet)) return false;
    const controller = interactionControllerFor(store);
    if (controller) {
      if (!controller.hasWholeSheetFontStyles()) return false;
      const ambient = selectionFormatOrigin.get(store);
      const decision = controller.canExecute({
        operation: 'format',
        origin: options.origin ?? ambient?.origin ?? 'instanceApi',
        commandId: options.commandId ?? ambient?.commandId,
        effects: [{ kind: 'workbook' }],
      });
      if (!decision.allowed || controller.policy?.editable !== undefined) return false;
    }
    store.setState((current) => {
      const formats = new Map(current.format.formats);
      for (const [key, previous] of formats) {
        const addr = parseAddrKey(key);
        if (addr?.sheet !== sheet) continue;
        if (!visualOnly) formats.delete(key);
        else {
          const next = stripVisualFormat(previous);
          if (next) formats.set(key, next);
          else formats.delete(key);
        }
      }
      const overrides = new Map(current.format.sheetFontOverrides ?? []);
      const previousOverride = overrides.get(sheet);
      overrides.set(sheet, {
        patch: {},
        clearAxes: true,
        ...(previousOverride?.baseline ? { baseline: previousOverride.baseline } : {}),
      });
      return {
        ...current,
        format: { ...current.format, formats, sheetFontOverrides: overrides },
        ui: { ...current.ui, pendingFormat: null },
      };
    });
    return true;
  }
  const plan = sparseSelectionPlan(state, store, options);
  if (!plan) return false;
  const selected = new Set(plan.cells.map(addrKey));
  store.setState((current) => {
    const formats = new Map(current.format.formats);
    for (const [key, format] of current.format.formats) {
      if (!selected.has(key) || !isCellWritable(current, parseAddrKey(key) as Addr)) continue;
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
  if (isWholeColumnRange(range) || isWholeRowRange(range)) {
    const axisBands: AxisFormatBand[] = [
      isWholeColumnRange(range)
        ? { axis: 'column', first: range.c0, last: range.c1 }
        : { axis: 'row', first: range.r0, last: range.r1 },
    ];
    const cells: Addr[] = [];
    for (const key of state.format.formats.keys()) {
      const addr = parseAddrKey(key);
      if (!addr || addr.sheet !== range.sheet) continue;
      if (
        (axisBands[0]?.axis === 'column' && addr.col >= range.c0 && addr.col <= range.c1) ||
        (axisBands[0]?.axis === 'row' && addr.row >= range.r0 && addr.row <= range.r1)
      )
        cells.push(addr);
    }
    return applyAxisFormatPatch(
      state,
      store,
      { ranges: [range], cells, exceptionCells: [], axisBands },
      patch ?? {},
    );
  }
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

const effectiveSelectionFormat = (
  state: State,
  store: SpreadsheetStore,
  addr: Addr,
): Partial<CellFormat> | undefined => {
  const workbook = formatHistoryBridgeFor(store)?.workbook;
  if (workbook) return effectiveCellFormatWithPending(workbook, state, addr) ?? undefined;
  return formatWithPending(state, addr);
};

const selectionAllHaveEffective = (
  state: State,
  store: SpreadsheetStore,
  plan: { cells: readonly Addr[]; axisBands: readonly AxisFormatBand[] },
  predicate: (format: Partial<CellFormat> | undefined) => boolean,
): boolean => {
  if (
    plan.cells.length > 0 &&
    !plan.cells.every((addr) => predicate(effectiveSelectionFormat(state, store, addr)))
  ) {
    return false;
  }
  if (plan.axisBands.length === 0)
    return plan.cells.every((addr) => predicate(effectiveSelectionFormat(state, store, addr)));
  const active = effectiveSelectionFormat(state, store, state.selection.active);
  return predicate(active);
};

const selectionAnyHaveEffective = (
  state: State,
  store: SpreadsheetStore,
  plan: { cells: readonly Addr[]; axisBands: readonly AxisFormatBand[] },
  predicate: (format: Partial<CellFormat> | undefined) => boolean,
): boolean => {
  if (plan.cells.some((addr) => predicate(effectiveSelectionFormat(state, store, addr))))
    return true;
  return plan.axisBands.length > 0
    ? predicate(effectiveSelectionFormat(state, store, state.selection.active))
    : false;
};

function toggleFlag(state: State, store: SpreadsheetStore, key: ToggleKey): void {
  const plan = selectionPlan(state);
  if (!plan) return;
  const allOn = selectionAllHaveEffective(state, store, plan, (f) => f?.[key] === true);
  const patch = { [key]: !allOn } as Partial<CellFormat>;
  if (plan.axisBands.length > 0) {
    applyAxisFormatPatch(state, store, plan, patch);
    return;
  }
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
  const allOn = selectionAllHaveEffective(state, store, plan, (f) => f?.wrap === true);
  if (plan.axisBands.length > 0) {
    applyAxisFormatPatch(state, store, plan, { wrap: !allOn });
    return;
  }
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
  if (plan.axisBands.length > 0) {
    const resolveIndent: AxisFormatPatchResolver = (wb, baseXfIndex, addr) => {
      const current = addr
        ? effectiveCellFormatWithPending(wb, state, addr)
        : (() => {
            const xf = wb.getCellXf(baseXfIndex);
            return xf ? cellFormatFromXf(wb, xf, wb.workbookDefaultFont) : null;
          })();
      const indent = current?.indent ?? 0;
      return { indent: Math.max(0, Math.min(15, indent + delta)) };
    };
    applyAxisFormatPatch(state, store, plan, {}, {}, resolveIndent);
    return;
  }
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

const defaultCurrencySymbol = currencySymbolForLocale;

export function cycleCurrency(state: State, store: SpreadsheetStore, locale = 'en-US'): void {
  const plan = selectionPlan(state);
  if (!plan) return;
  const hasCurrency = selectionAnyHaveEffective(
    state,
    store,
    plan,
    (f) => f?.numFmt?.kind === 'currency',
  );
  const next: NumFmt = hasCurrency
    ? { kind: 'general' }
    : { kind: 'currency', decimals: 2, symbol: defaultCurrencySymbol(locale) };
  if (plan.axisBands.length > 0) {
    applyAxisFormatPatch(state, store, plan, { numFmt: next });
    return;
  }
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
  const hasPercent = selectionAnyHaveEffective(
    state,
    store,
    plan,
    (f) => f?.numFmt?.kind === 'percent',
  );
  const next: NumFmt = hasPercent ? { kind: 'general' } : { kind: 'percent', decimals: 0 };
  if (plan.axisBands.length > 0) {
    applyAxisFormatPatch(state, store, plan, { numFmt: next });
    return;
  }
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
  if (plan.axisBands.length > 0) {
    const resolveDecimals: AxisFormatPatchResolver = (wb, baseXfIndex, addr) => {
      const current = addr
        ? effectiveCellFormatWithPending(wb, state, addr)
        : (() => {
            const xf = wb.getCellXf(baseXfIndex);
            return xf ? cellFormatFromXf(wb, xf, wb.workbookDefaultFont) : null;
          })();
      const cur = current?.numFmt;
      if (!cur || cur.kind === 'general') {
        return delta === 1 ? { numFmt: { kind: 'fixed', decimals: 2 } } : {};
      }
      if (cur.kind === 'fixed') {
        return { numFmt: { kind: 'fixed', decimals: clampDecimals(cur.decimals + delta) } };
      }
      if (cur.kind === 'currency') {
        return {
          numFmt: {
            kind: 'currency',
            decimals: clampDecimals(cur.decimals + delta),
            ...(cur.symbol !== undefined ? { symbol: cur.symbol } : {}),
          },
        };
      }
      if (cur.kind === 'percent') {
        return { numFmt: { kind: 'percent', decimals: clampDecimals(cur.decimals + delta) } };
      }
      return {};
    };
    applyAxisFormatPatch(state, store, plan, {}, {}, resolveDecimals);
    return;
  }
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

  if (plan.axisBands.length > 0) {
    // Presets such as Outline/Inside describe rectangle topology and cannot
    // be represented by a compact row/column style. Constant style fields
    // from the dialog are safe and follow the same sparse/native path as the
    // toolbar commands.
    if (action.border || !axisPatchSupported(action.patch)) return false;
    return applyAxisFormatPatch(state, store, plan, action.patch, options);
  }

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
  // Border presets describe rectangle topology. A compact row/column XF can
  // carry a constant border, but cannot encode an edge/inside geometry over a
  // full axis without expanding every cell.
  if (plan.axisBands.length > 0) return;
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
  if (plan.axisBands.length > 0) {
    const active = effectiveSelectionFormat(state, store, state.selection.active);
    const borders = active?.borders;
    const allFour = !!(borders?.top && borders.right && borders.bottom && borders.left);
    applyAxisFormatPatch(state, store, plan, {
      borders: allFour
        ? { top: false, right: false, bottom: false, left: false }
        : { top: true, right: true, bottom: true, left: true },
    });
    return;
  }
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
): boolean {
  const next: Partial<CellFormat> = {};
  if (patch.fontFamily !== undefined) {
    next.fontFamily = patch.fontFamily ?? undefined;
  }
  if (patch.fontSize !== undefined) {
    next.fontSize = patch.fontSize ?? undefined;
  }
  if (Object.keys(next).length === 0) return false;
  const ranges = selectionRanges(state);
  if (ranges && isWholeSheetSelection(ranges)) {
    const sheet = ranges[0]?.sheet;
    // Once a workbook has a compact row/column override, keep that override
    // as the single native authority. Applying a worksheet-wide font through
    // the legacy map would be overwritten by the axis flush on the next
    // publication (and would make the subsequent undo chain ambiguous).
    if (sheet !== undefined && state.format.sheetAxisOverrides?.has(sheet)) {
      const plan = selectionPlan(state);
      if (!plan) return false;
      const axisPlan = {
        ...plan,
        cells: [...state.format.formats.keys()]
          .map(parseAddrKey)
          .filter((addr): addr is Addr => addr?.sheet === sheet),
        exceptionCells: [],
        // A full-width column band covers every unstyled row. Existing row
        // styles are handled as sparse precedence intersections by the axis
        // planner; adding a MAX_ROW row band would require million-row native
        // writes for an otherwise compact worksheet-wide font update.
        axisBands: [{ axis: 'column' as const, first: 0, last: MAX_COL }],
      };
      return applyAxisFormatPatch(state, store, axisPlan, next);
    }
    return applyWholeSheetFontPatch(state, store, ranges, next);
  }
  const plan = selectionPlan(state);
  if (plan?.axisBands.length) return applyAxisFormatPatch(state, store, plan, next);
  return applySelectionFormatPatch(state, store, next);
}

const hasOwn = (value: object, key: PropertyKey): boolean => Object.hasOwn(value, key);

/**
 * The ordinary format planner intentionally refuses to materialise a whole
 * worksheet. Font family and size are the two dimensions for which the
 * engine has compact row/column style storage, so keep this branch narrow and
 * atomic. Existing store/data entries receive the same patch; blank grid
 * addresses remain implicit in `sheetFontOverrides`.
 */
function applyWholeSheetFontPatch(
  state: State,
  store: SpreadsheetStore,
  ranges: readonly Range[],
  patch: Partial<CellFormat>,
): boolean {
  const sheet = ranges[0]?.sheet;
  if (sheet === undefined || ranges.some((range) => range.sheet !== sheet)) return false;
  // A protected worksheet cannot safely be represented by one axis style:
  // locked and unlocked cells would need a per-cell authorization pass over
  // the entire grid. Keep the operation atomic and reject it instead.
  if (isSheetProtected(state, sheet)) return false;

  const controller = interactionControllerFor(store);
  if (controller) {
    if (!controller.hasWholeSheetFontStyles()) return false;
    const ambient = selectionFormatOrigin.get(store);
    const decision = controller.canExecute({
      operation: 'format',
      origin: ambient?.origin ?? 'instanceApi',
      commandId: ambient?.commandId,
      effects: [{ kind: 'workbook' }],
    });
    if (!decision.allowed) return false;
    // A host editable predicate/range cannot be proven over a 17-billion-cell
    // effect without defeating the sparse guarantee. Reject conservatively.
    if (controller.policy?.editable !== undefined) return false;
  }

  const hasFamily = hasOwn(patch, 'fontFamily');
  const hasSize = hasOwn(patch, 'fontSize');
  if (!hasFamily && !hasSize) return false;
  if (
    (hasFamily && patch.fontFamily !== undefined && patch.fontFamily.trim().length === 0) ||
    (hasSize &&
      patch.fontSize !== undefined &&
      (!Number.isFinite(patch.fontSize) || patch.fontSize <= 0))
  ) {
    return false;
  }

  store.setState((current) => {
    const formats = new Map(current.format.formats);
    const selectedKeys = new Set<string>();
    for (const key of current.data.cells.keys()) {
      const addr = parseAddrKey(key);
      if (addr?.sheet === sheet) selectedKeys.add(key);
    }
    for (const key of current.format.formats.keys()) {
      const addr = parseAddrKey(key);
      if (addr?.sheet === sheet) selectedKeys.add(key);
    }
    for (const key of selectedKeys) {
      const previous = formats.get(key);
      const next: CellFormat = { ...(previous ?? {}) };
      if (hasFamily) {
        if (patch.fontFamily === undefined) delete next.fontFamily;
        else next.fontFamily = patch.fontFamily;
      }
      if (hasSize) {
        if (patch.fontSize === undefined) delete next.fontSize;
        else next.fontSize = patch.fontSize;
      }
      if (Object.keys(next).length === 0) formats.delete(key);
      else formats.set(key, next);
    }

    const overrides = new Map(current.format.sheetFontOverrides ?? []);
    const previous = overrides.get(sheet);
    const nextPatch = { ...(previous?.patch ?? {}) };
    if (hasFamily) {
      if (patch.fontFamily === undefined) delete nextPatch.fontFamily;
      else nextPatch.fontFamily = patch.fontFamily;
    }
    if (hasSize) {
      if (patch.fontSize === undefined) delete nextPatch.fontSize;
      else nextPatch.fontSize = patch.fontSize;
    }
    if (Object.keys(nextPatch).length === 0) {
      if (previous?.clearAxes) {
        overrides.set(sheet, {
          patch: {},
          clearAxes: true,
          ...(previous.baseline ? { baseline: previous.baseline } : {}),
        });
      } else {
        overrides.delete(sheet);
      }
    } else {
      overrides.set(sheet, {
        patch: nextPatch,
        ...(previous?.clearAxes ? { clearAxes: true } : {}),
        ...(previous?.baseline ? { baseline: previous.baseline } : {}),
      });
    }

    return {
      ...current,
      format: {
        ...current.format,
        formats,
        sheetFontOverrides: overrides,
      },
      ui: { ...current.ui, pendingFormat: null },
    };
  });
  return true;
}
