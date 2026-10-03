import { syncLayoutToEngine } from '../engine/layout-sync.js';
import type { Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type {
  CellFormat,
  ConditionalRule,
  CustomCellStyle,
  LayoutSlice,
  PageMargins,
  PageSetup,
  SessionChart,
  SessionIllustration,
  SlicerSpec,
  Sparkline,
  SpreadsheetStore,
  State,
} from '../store/store.js';
import type {
  CustomTableStyle,
  PivotTableStyleAssignment,
  TableOverlay,
} from './format-as-table.js';
import type { History } from './history.js';

/* ---------- Snapshot helpers ---------- */

/** Capture the entire format map. Sufficient for undo since the map is sparse
 *  (only explicitly formatted cells have entries). */
export interface FormatSnapshot {
  formats: Map<string, CellFormat>;
  customCellStyles: CustomCellStyle[];
}

const cloneCustomCellStyle = (style: CustomCellStyle): CustomCellStyle => ({
  ...style,
  format: { ...style.format },
});

export function captureFormatSnapshot(state: State): FormatSnapshot {
  return {
    formats: new Map(state.format.formats),
    customCellStyles: (state.format.customCellStyles ?? []).map(cloneCustomCellStyle),
  };
}

export function applyFormatSnapshot(store: SpreadsheetStore, snap: FormatSnapshot): void {
  store.setState((s) => ({
    ...s,
    format: {
      formats: new Map(snap.formats),
      customCellStyles: snap.customCellStyles.map(cloneCustomCellStyle),
    },
  }));
}

const formatSnapshotKey = (snap: FormatSnapshot): string =>
  JSON.stringify({
    formats: [...snap.formats.entries()].sort(([a], [b]) => a.localeCompare(b)),
    customCellStyles: snap.customCellStyles,
  });

const sameFormatSnapshot = (a: FormatSnapshot, b: FormatSnapshot): boolean =>
  a.formats.size === b.formats.size &&
  a.customCellStyles.length === b.customCellStyles.length &&
  formatSnapshotKey(a) === formatSnapshotKey(b);

const mapSnapshotKey = <T>(snap: Map<string, T>): string =>
  JSON.stringify([...snap.entries()].sort(([a], [b]) => a.localeCompare(b)));

const sameMapSnapshot = <T>(a: Map<string, T>, b: Map<string, T>): boolean =>
  a.size === b.size && mapSnapshotKey(a) === mapSnapshotKey(b);

const sameJsonSnapshot = <T>(a: readonly T[], b: readonly T[]): boolean =>
  a.length === b.length && JSON.stringify(a) === JSON.stringify(b);

/** How one store slice is snapshotted, restored and compared for history. */
interface SliceCodec<T> {
  capture: (state: State) => T;
  apply: (store: SpreadsheetStore, snap: T) => void;
  same: (a: T, b: T) => boolean;
  repeat?: () => void;
}

/** Run `mutate`, capturing the slice before and after, and push one entry
 *  unless the slice is unchanged. Runs `mutate` bare when `history` is null or
 *  replaying. */
function recordSliceChange<T>(
  history: History | null,
  store: SpreadsheetStore,
  mutate: () => void,
  codec: SliceCodec<T>,
): void {
  if (!history || history.isReplaying()) {
    mutate();
    return;
  }
  const before = codec.capture(store.getState());
  mutate();
  const after = codec.capture(store.getState());
  if (codec.same(before, after)) return;
  history.push({
    undo: () => codec.apply(store, before),
    redo: () => codec.apply(store, after),
    repeat: codec.repeat,
  });
}

const sameTablesSnapshot = (a: TablesSnapshot, b: TablesSnapshot): boolean =>
  JSON.stringify(a) === JSON.stringify(b);

const cloneTableOverlay = (table: TableOverlay): TableOverlay => ({
  ...table,
  range: { ...table.range },
});

export interface TablesSnapshot {
  tables: TableOverlay[];
  customTableStyles: CustomTableStyle[];
  customPivotTableStyles: CustomTableStyle[];
  pivotTableStyles: PivotTableStyleAssignment[];
}

const cloneCustomTableStyle = (style: CustomTableStyle): CustomTableStyle => ({ ...style });
const clonePivotTableStyle = (style: PivotTableStyleAssignment): PivotTableStyleAssignment => ({
  ...style,
});

export function captureTableOverlaysSnapshot(state: State): TablesSnapshot {
  return {
    tables: state.tables.tables.map(cloneTableOverlay),
    customTableStyles: (state.tables.customTableStyles ?? []).map(cloneCustomTableStyle),
    customPivotTableStyles: (state.tables.customPivotTableStyles ?? []).map(cloneCustomTableStyle),
    pivotTableStyles: (state.tables.pivotTableStyles ?? []).map(clonePivotTableStyle),
  };
}

export function applyTableOverlaysSnapshot(
  store: SpreadsheetStore,
  snap: TablesSnapshot | readonly TableOverlay[],
): void {
  const normalized: TablesSnapshot = Array.isArray(snap)
    ? {
        tables: (snap as readonly TableOverlay[]).map(cloneTableOverlay),
        customTableStyles: [],
        customPivotTableStyles: [],
        pivotTableStyles: [],
      }
    : (snap as TablesSnapshot);
  store.setState((s) => ({
    ...s,
    tables: {
      tables: normalized.tables.map(cloneTableOverlay),
      customTableStyles: normalized.customTableStyles.map(cloneCustomTableStyle),
      customPivotTableStyles: normalized.customPivotTableStyles.map(cloneCustomTableStyle),
      pivotTableStyles: normalized.pivotTableStyles.map(clonePivotTableStyle),
    },
  }));
}

export function recordTablesChange(
  history: History | null,
  store: SpreadsheetStore,
  mutate: () => void,
): void {
  recordSliceChange(history, store, mutate, {
    capture: captureTableOverlaysSnapshot,
    apply: applyTableOverlaysSnapshot,
    same: sameTablesSnapshot,
  });
}

const cloneSessionChart = (chart: SessionChart): SessionChart => ({
  ...chart,
  source: { ...chart.source },
});

export function captureChartsSnapshot(state: State): SessionChart[] {
  return state.charts.charts.map(cloneSessionChart);
}

export function applyChartsSnapshot(store: SpreadsheetStore, snap: readonly SessionChart[]): void {
  store.setState((s) => ({
    ...s,
    charts: { charts: snap.map(cloneSessionChart) },
  }));
}

export function recordChartsChange(
  history: History | null,
  store: SpreadsheetStore,
  mutate: () => void,
): void {
  recordSliceChange(history, store, mutate, {
    capture: captureChartsSnapshot,
    apply: applyChartsSnapshot,
    same: sameJsonSnapshot,
  });
}

const cloneSessionIllustration = (item: SessionIllustration): SessionIllustration => ({ ...item });

export function captureIllustrationsSnapshot(state: State): SessionIllustration[] {
  return state.illustrations.illustrations.map(cloneSessionIllustration);
}

export function applyIllustrationsSnapshot(
  store: SpreadsheetStore,
  snap: readonly SessionIllustration[],
): void {
  store.setState((s) => ({
    ...s,
    illustrations: { illustrations: snap.map(cloneSessionIllustration) },
  }));
}

export function recordIllustrationsChange(
  history: History | null,
  store: SpreadsheetStore,
  mutate: () => void,
): void {
  recordSliceChange(history, store, mutate, {
    capture: captureIllustrationsSnapshot,
    apply: applyIllustrationsSnapshot,
    same: sameJsonSnapshot,
  });
}

export interface LayoutSnapshot {
  colWidths: Map<number, number>;
  rowHeights: Map<number, number>;
  freezeRows: number;
  freezeCols: number;
  hiddenRows: Set<number>;
  hiddenCols: Set<number>;
  outlineRows: Map<number, number>;
  outlineCols: Map<number, number>;
  outlineRowGutter: number;
  outlineColGutter: number;
  hiddenSheets: Set<number>;
  veryHiddenSheets: Set<number>;
  sheetTabColors: Map<number, string>;
}

export function captureLayoutSnapshot(state: State): LayoutSnapshot {
  return {
    colWidths: new Map(state.layout.colWidths),
    rowHeights: new Map(state.layout.rowHeights),
    freezeRows: state.layout.freezeRows,
    freezeCols: state.layout.freezeCols,
    hiddenRows: new Set(state.layout.hiddenRows),
    hiddenCols: new Set(state.layout.hiddenCols),
    outlineRows: new Map(state.layout.outlineRows),
    outlineCols: new Map(state.layout.outlineCols),
    outlineRowGutter: state.layout.outlineRowGutter,
    outlineColGutter: state.layout.outlineColGutter,
    hiddenSheets: new Set(state.layout.hiddenSheets),
    veryHiddenSheets: new Set(state.layout.veryHiddenSheets),
    sheetTabColors: new Map(state.layout.sheetTabColors),
  };
}

export function applyLayoutSnapshot(store: SpreadsheetStore, snap: LayoutSnapshot): void {
  store.setState((s) => ({
    ...s,
    layout: {
      ...s.layout,
      colWidths: new Map(snap.colWidths),
      rowHeights: new Map(snap.rowHeights),
      freezeRows: snap.freezeRows,
      freezeCols: snap.freezeCols,
      hiddenRows: new Set(snap.hiddenRows),
      hiddenCols: new Set(snap.hiddenCols),
      outlineRows: new Map(snap.outlineRows),
      outlineCols: new Map(snap.outlineCols),
      outlineRowGutter: snap.outlineRowGutter,
      outlineColGutter: snap.outlineColGutter,
      hiddenSheets: new Set(snap.hiddenSheets),
      veryHiddenSheets: new Set(snap.veryHiddenSheets),
      sheetTabColors: new Map(snap.sheetTabColors),
    } as LayoutSlice,
  }));
}

const sameNumberMap = (a: Map<number, number>, b: Map<number, number>): boolean =>
  a.size === b.size && [...a].every(([key, value]) => b.get(key) === value);

const sameNumberSet = (a: Set<number>, b: Set<number>): boolean =>
  a.size === b.size && [...a].every((value) => b.has(value));

const sameNumberStringMap = (a: Map<number, string>, b: Map<number, string>): boolean =>
  a.size === b.size && [...a].every(([key, value]) => b.get(key) === value);

const sameLayoutSnapshot = (a: LayoutSnapshot, b: LayoutSnapshot): boolean =>
  sameNumberMap(a.colWidths, b.colWidths) &&
  sameNumberMap(a.rowHeights, b.rowHeights) &&
  a.freezeRows === b.freezeRows &&
  a.freezeCols === b.freezeCols &&
  sameNumberSet(a.hiddenRows, b.hiddenRows) &&
  sameNumberSet(a.hiddenCols, b.hiddenCols) &&
  sameNumberMap(a.outlineRows, b.outlineRows) &&
  sameNumberMap(a.outlineCols, b.outlineCols) &&
  a.outlineRowGutter === b.outlineRowGutter &&
  a.outlineColGutter === b.outlineColGutter &&
  sameNumberSet(a.hiddenSheets, b.hiddenSheets) &&
  sameNumberSet(a.veryHiddenSheets, b.veryHiddenSheets) &&
  sameNumberStringMap(a.sheetTabColors, b.sheetTabColors);

/** Run `mutate`, capturing the format slice before and after, pushing one
 *  entry. No-op when `history` is null. */
export function recordFormatChange(
  history: History | null,
  store: SpreadsheetStore,
  mutate: () => void,
  opts: { repeat?: () => void } = {},
): void {
  recordSliceChange(history, store, mutate, {
    capture: captureFormatSnapshot,
    apply: applyFormatSnapshot,
    same: sameFormatSnapshot,
    repeat: opts.repeat,
  });
}

/** Record a format command whose mutation derives its target from the current
 *  store state. It is therefore safe to repeat with F4 after selection moves.
 *  Callers that captured a concrete range (paste/fill) must continue using
 *  `recordFormatChange` directly. */
export function recordRepeatableFormatChange(
  history: History | null,
  store: SpreadsheetStore,
  mutate: () => void,
): void {
  const repeat = (): void => recordRepeatableFormatChange(history, store, mutate);
  recordFormatChange(history, store, mutate, { repeat });
  history?.setRepeat(repeat);
}

/** Record a format command that was handed a concrete range, together with the
 *  repeat that should run in its place when F4 fires. Use this where the
 *  mutation cannot simply be re-run — the caller supplies a `repeat` that
 *  re-targets the current selection itself. */
export function recordFormatChangeWithRepeat(
  history: History | null,
  store: SpreadsheetStore,
  mutate: () => void,
  repeat: () => void,
): void {
  recordFormatChange(history, store, mutate, { repeat });
  history?.setRepeat(repeat);
}

export function recordLayoutChange(
  history: History | null,
  store: SpreadsheetStore,
  mutate: () => void,
): void {
  recordSliceChange(history, store, mutate, {
    capture: captureLayoutSnapshot,
    apply: applyLayoutSnapshot,
    same: sameLayoutSnapshot,
  });
}

const cloneConditionalRule = (rule: ConditionalRule): ConditionalRule => {
  const out = { ...rule, range: { ...rule.range } } as ConditionalRule;
  if ('apply' in out) out.apply = { ...out.apply };
  if ('stops' in out) {
    out.stops = [...out.stops] as [string, string] | [string, string, string];
  }
  return out;
};

export function captureConditionalRulesSnapshot(state: State): ConditionalRule[] {
  return state.conditional.rules.map(cloneConditionalRule);
}

export function applyConditionalRulesSnapshot(
  store: SpreadsheetStore,
  snap: readonly ConditionalRule[],
): void {
  store.setState((s) => ({
    ...s,
    conditional: { rules: snap.map(cloneConditionalRule) },
  }));
}

export function recordConditionalRulesChange(
  history: History | null,
  store: SpreadsheetStore,
  mutate: () => void,
): void {
  recordSliceChange(history, store, mutate, {
    capture: captureConditionalRulesSnapshot,
    apply: applyConditionalRulesSnapshot,
    same: sameJsonSnapshot,
  });
}

/** Engine-aware layout change. Same semantics as `recordLayoutChange` but
 *  the captured before/after pair is also pushed to the workbook engine for
 *  the active sheet, both at apply time and on every undo/redo replay.
 *  Skipped (including the engine sync) when `wb` is null. The per-method
 *  calls inside `syncLayoutToEngine` short-circuit on each capability flag,
 *  so engines that only support a subset still work. */
export function recordLayoutChangeWithEngine(
  history: History | null,
  store: SpreadsheetStore,
  wb: WorkbookHandle | null,
  mutate: () => void,
): void {
  if (!wb) {
    recordLayoutChange(history, store, mutate);
    return;
  }
  const sheet = store.getState().data.sheetIndex;
  const before = captureLayoutSnapshot(store.getState());
  mutate();
  const after = captureLayoutSnapshot(store.getState());
  if (sameLayoutSnapshot(before, after)) return;
  syncLayoutToEngine(wb, store.getState().layout, sheet, before, after);
  if (!history || history.isReplaying()) return;
  history.push({
    undo: () => {
      applyLayoutSnapshot(store, before);
      syncLayoutToEngine(wb, store.getState().layout, sheet, after, before);
    },
    redo: () => {
      applyLayoutSnapshot(store, after);
      syncLayoutToEngine(wb, store.getState().layout, sheet, before, after);
    },
  });
}

export interface MergesSnapshot {
  byAnchor: Map<string, Range>;
  byCell: Map<string, string>;
}

export function captureMergesSnapshot(state: State): MergesSnapshot {
  return {
    byAnchor: new Map(state.merges.byAnchor),
    byCell: new Map(state.merges.byCell),
  };
}

export function applyMergesSnapshot(store: SpreadsheetStore, snap: MergesSnapshot): void {
  store.setState((s) => ({
    ...s,
    merges: { byAnchor: new Map(snap.byAnchor), byCell: new Map(snap.byCell) },
  }));
}

const rangeKey = (range: Range): string =>
  `${range.sheet}:${range.r0}:${range.c0}:${range.r1}:${range.c1}`;

const sameStringMap = (a: Map<string, string>, b: Map<string, string>): boolean =>
  a.size === b.size && [...a].every(([key, value]) => b.get(key) === value);

const sameRangeMap = (a: Map<string, Range>, b: Map<string, Range>): boolean =>
  a.size === b.size &&
  [...a].every(([key, range]) => {
    const other = b.get(key);
    return other !== undefined && rangeKey(range) === rangeKey(other);
  });

const sameMergesSnapshot = (a: MergesSnapshot, b: MergesSnapshot): boolean =>
  sameRangeMap(a.byAnchor, b.byAnchor) && sameStringMap(a.byCell, b.byCell);

export function recordMergesChange(
  history: History | null,
  store: SpreadsheetStore,
  mutate: () => void,
): void {
  recordSliceChange(history, store, mutate, {
    capture: captureMergesSnapshot,
    apply: applyMergesSnapshot,
    same: sameMergesSnapshot,
  });
}

/** Engine-aware merges change. Mirrors the post-mutate state into the workbook
 *  via `clearMerges` + per-anchor `addMerge`. Both apply and undo/redo go
 *  through the same path, so the engine snapshot stays in lockstep with the
 *  store. No-op against the engine when `wb` is null or `capabilities.merges`
 *  is off. */
export function recordMergesChangeWithEngine(
  history: History | null,
  store: SpreadsheetStore,
  wb: WorkbookHandle | null,
  sheet: number,
  mutate: () => void,
): void {
  const sync = (snap: MergesSnapshot): void => {
    if (!wb?.capabilities.merges) return;
    wb.engineClearMerges(sheet);
    for (const r of snap.byAnchor.values()) {
      if (r.sheet !== sheet) continue;
      wb.engineAddMerge(sheet, r);
    }
  };
  const before = captureMergesSnapshot(store.getState());
  mutate();
  const after = captureMergesSnapshot(store.getState());
  if (sameMergesSnapshot(before, after)) return;
  sync(after);
  if (!history || history.isReplaying()) return;
  history.push({
    undo: () => {
      applyMergesSnapshot(store, before);
      sync(before);
    },
    redo: () => {
      applyMergesSnapshot(store, after);
      sync(after);
    },
  });
}

export function captureSparklineSnapshot(state: State): Map<string, Sparkline> {
  return new Map(state.sparkline.sparklines);
}

export function applySparklineSnapshot(
  store: SpreadsheetStore,
  snap: Map<string, Sparkline>,
): void {
  store.setState((s) => ({ ...s, sparkline: { sparklines: new Map(snap) } }));
}

export function recordSparklineChange(
  history: History | null,
  store: SpreadsheetStore,
  mutate: () => void,
): void {
  recordSliceChange(history, store, mutate, {
    capture: captureSparklineSnapshot,
    apply: applySparklineSnapshot,
    same: sameMapSnapshot,
  });
}

export function capturePageSetupSnapshot(state: State): Map<number, PageSetup> {
  const out = new Map<number, PageSetup>();
  for (const [k, v] of state.pageSetup.setupBySheet) {
    out.set(k, {
      ...v,
      margins: { ...v.margins },
      printableBounds: v.printableBounds ? { ...v.printableBounds } : undefined,
      manualPageBreakRows: v.manualPageBreakRows ? [...v.manualPageBreakRows] : undefined,
      manualPageBreakCols: v.manualPageBreakCols ? [...v.manualPageBreakCols] : undefined,
    });
  }
  return out;
}

export function applyPageSetupSnapshot(
  store: SpreadsheetStore,
  snap: Map<number, PageSetup>,
): void {
  const next = new Map<number, PageSetup>();
  for (const [k, v] of snap) {
    next.set(k, {
      ...v,
      margins: { ...v.margins },
      printableBounds: v.printableBounds ? { ...v.printableBounds } : undefined,
      manualPageBreakRows: v.manualPageBreakRows ? [...v.manualPageBreakRows] : undefined,
      manualPageBreakCols: v.manualPageBreakCols ? [...v.manualPageBreakCols] : undefined,
    });
  }
  store.setState((s) => ({ ...s, pageSetup: { setupBySheet: next } }));
}

const sameOptionalArray = (a: readonly number[] | undefined, b: readonly number[] | undefined) => {
  if (!a && !b) return true;
  if (!a || !b || a.length !== b.length) return false;
  return a.every((value, index) => value === b[index]);
};

const sameOptionalMargins = (a: PageMargins | undefined, b: PageMargins | undefined): boolean => {
  if (!a && !b) return true;
  if (!a || !b) return false;
  return a.top === b.top && a.right === b.right && a.bottom === b.bottom && a.left === b.left;
};

const samePageSetup = (a: PageSetup, b: PageSetup): boolean =>
  a.orientation === b.orientation &&
  a.paperSize === b.paperSize &&
  a.paperSizeCode === b.paperSizeCode &&
  a.margins.top === b.margins.top &&
  a.margins.right === b.margins.right &&
  a.margins.bottom === b.margins.bottom &&
  a.margins.left === b.margins.left &&
  sameOptionalMargins(a.printableBounds, b.printableBounds) &&
  a.headerMargin === b.headerMargin &&
  a.footerMargin === b.footerMargin &&
  a.centerHorizontally === b.centerHorizontally &&
  a.centerVertically === b.centerVertically &&
  a.headerLeft === b.headerLeft &&
  a.headerCenter === b.headerCenter &&
  a.headerRight === b.headerRight &&
  a.footerLeft === b.footerLeft &&
  a.footerCenter === b.footerCenter &&
  a.footerRight === b.footerRight &&
  a.differentOddEvenPages === b.differentOddEvenPages &&
  a.differentFirstPage === b.differentFirstPage &&
  a.scaleHeaderFooterWithDocument === b.scaleHeaderFooterWithDocument &&
  a.alignHeaderFooterWithMargins === b.alignHeaderFooterWithMargins &&
  a.printArea === b.printArea &&
  a.printTitleRows === b.printTitleRows &&
  a.printTitleCols === b.printTitleCols &&
  a.fitWidth === b.fitWidth &&
  a.fitHeight === b.fitHeight &&
  sameOptionalArray(a.manualPageBreakRows, b.manualPageBreakRows) &&
  sameOptionalArray(a.manualPageBreakCols, b.manualPageBreakCols) &&
  a.scale === b.scale &&
  a.printQuality === b.printQuality &&
  a.firstPageNumber === b.firstPageNumber &&
  a.showGridlines === b.showGridlines &&
  a.showHeadings === b.showHeadings &&
  a.blackAndWhite === b.blackAndWhite &&
  a.draftQuality === b.draftQuality &&
  a.comments === b.comments &&
  a.cellErrorsAs === b.cellErrorsAs &&
  a.pageOrder === b.pageOrder;

const samePageSetupSnapshot = (a: Map<number, PageSetup>, b: Map<number, PageSetup>): boolean => {
  if (a.size !== b.size) return false;
  for (const [sheet, setup] of a) {
    const other = b.get(sheet);
    if (!other || !samePageSetup(setup, other)) return false;
  }
  return true;
};

/** Run `mutate`, capturing the page-setup slice before and after, pushing one
 *  entry. No-op when `history` is null. */
export function recordPageSetupChange(
  history: History | null,
  store: SpreadsheetStore,
  mutate: () => void,
): void {
  recordSliceChange(history, store, mutate, {
    capture: capturePageSetupSnapshot,
    apply: applyPageSetupSnapshot,
    same: samePageSetupSnapshot,
  });
}

/** Capture a deep-cloned slicer list for undo replay. Each spec is freshly
 *  cloned so future mutators that mutate the `selected` array can't
 *  retroactively pollute a prior snapshot. */
export function captureSlicersSnapshot(state: State): SlicerSpec[] {
  return state.slicers.slicers.map((sp) => ({ ...sp, selected: [...sp.selected] }));
}

export function applySlicersSnapshot(store: SpreadsheetStore, snap: readonly SlicerSpec[]): void {
  store.setState((s) => ({
    ...s,
    slicers: { slicers: snap.map((sp) => ({ ...sp, selected: [...sp.selected] })) },
  }));
}

/** Run `mutate` and push one history entry capturing the slicer-slice
 *  before/after. Use for any add/remove/update/setSelected call that should
 *  be undoable. */
export function recordSlicersChange(
  history: History | null,
  store: SpreadsheetStore,
  mutate: () => void,
): void {
  recordSliceChange(history, store, mutate, {
    capture: captureSlicersSnapshot,
    apply: applySlicersSnapshot,
    same: sameJsonSnapshot,
  });
}
