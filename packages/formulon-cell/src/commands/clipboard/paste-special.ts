import { addrKey } from '../../engine/address.js';
import type { Addr, Range } from '../../engine/types.js';
import type { WorkbookHandle } from '../../engine/workbook-handle.js';
import { type CellFormat, mutators, type SpreadsheetStore, type State } from '../../store/store.js';
import { recordCommentChange } from '../comment.js';
import { adjustFormulaForCutPasteMove } from '../formula-refs.js';
import type { History } from '../history.js';
import {
  recordFormatChange,
  recordLayoutChangeWithEngine,
  recordMergesChangeWithEngine,
} from '../history.js';
import { isCellWritable } from '../protection.js';
import { shiftFormulaRefs } from '../refs.js';
import type { ClipboardCell, ClipboardSnapshot } from './snapshot.js';

export type PasteWhat =
  | 'all'
  | 'values'
  | 'formulas'
  | 'formats'
  | 'formulas-and-numfmt'
  | 'values-and-numfmt';

export type PasteOperation = 'none' | 'add' | 'subtract' | 'multiply' | 'divide';

export interface PasteSpecialOptions {
  what: PasteWhat;
  operation: PasteOperation;
  skipBlanks: boolean;
  transpose: boolean;
}

export interface PasteSpecialResult {
  writtenRange: Range;
  /** Arithmetic operations that produced NaN/Infinity and were written as
   *  static spreadsheet error values. */
  skippedNonFiniteOperations: number;
}

const numericValue = (cell: ClipboardCell | undefined): number | null => {
  if (!cell) return null;
  return cell.value.kind === 'number' ? cell.value.value : null;
};

const writeClipboardValue = (wb: WorkbookHandle, addr: Addr, src: ClipboardCell): void => {
  switch (src.value.kind) {
    case 'number':
      wb.setNumber(addr, src.value.value);
      break;
    case 'text':
      wb.setText(addr, src.value.value);
      break;
    case 'bool':
      wb.setBool(addr, src.value.value);
      break;
    case 'blank':
      wb.setBlank(addr);
      break;
    case 'error':
      wb.setError(addr, src.value.code);
      break;
  }
};

const existingNumeric = (state: State, sheet: number, row: number, col: number): number => {
  const cell = state.data.cells.get(addrKey({ sheet, row, col }));
  if (!cell) return 0;
  if (cell.value.kind === 'number') return cell.value.value;
  return 0;
};

const combine = (op: PasteOperation, dest: number, src: number): number => {
  switch (op) {
    case 'add':
      return dest + src;
    case 'subtract':
      return dest - src;
    case 'multiply':
      return dest * src;
    case 'divide':
      return src === 0 ? Number.NaN : dest / src;
    default:
      return src;
  }
};

const errorCodeForNonFiniteOperation = (op: PasteOperation): number => (op === 'divide' ? 1 : 5); // #DIV/0! for divide, #NUM! for arithmetic overflow.

const wantsValues = (what: PasteWhat): boolean =>
  what === 'all' || what === 'values' || what === 'values-and-numfmt';
const wantsFormulas = (what: PasteWhat): boolean =>
  what === 'all' || what === 'formulas' || what === 'formulas-and-numfmt';
const wantsFormats = (what: PasteWhat): boolean => what === 'all' || what === 'formats';
const wantsNumFmt = (what: PasteWhat): boolean =>
  what === 'all' ||
  what === 'formats' ||
  what === 'values-and-numfmt' ||
  what === 'formulas-and-numfmt';

const MAX_ROW = 1_048_575;
const MAX_COL = 16_383;
const MAX_PASTE_CELLS = 1_000_000;

type BandAxis = 'row' | 'column';

export interface MaterializedPasteCell {
  row: number;
  col: number;
  sourceRow: number;
  sourceCol: number;
  sourceRowIndex: number;
  sourceColIndex: number;
}

const rangesIntersect = (a: Range, b: Range): boolean =>
  a.sheet === b.sheet && !(a.r1 < b.r0 || a.r0 > b.r1 || a.c1 < b.c0 || a.c0 > b.c1);

const rangeContains = (outer: Range, inner: Range): boolean =>
  outer.sheet === inner.sheet &&
  inner.r0 >= outer.r0 &&
  inner.c0 >= outer.c0 &&
  inner.r1 <= outer.r1 &&
  inner.c1 <= outer.c1;

const sameRange = (left: Range, right: Range): boolean =>
  left.sheet === right.sheet &&
  left.r0 === right.r0 &&
  left.c0 === right.c0 &&
  left.r1 === right.r1 &&
  left.c1 === right.c1;

const logicalRangeFor = (snap: ClipboardSnapshot): Range => snap.logicalRange ?? snap.range;

const bandAxisFor = (range: Range): BandAxis | null => {
  if (range.r0 === 0 && range.r1 >= MAX_ROW) return 'column';
  if (range.c0 === 0 && range.c1 >= MAX_COL) return 'row';
  return null;
};

const normalizedSelection = (state: State): Range => {
  const selected = state.selection.range;
  return {
    sheet: selected.sheet,
    r0: Math.min(selected.r0, selected.r1),
    c0: Math.min(selected.c0, selected.c1),
    r1: Math.max(selected.r0, selected.r1),
    c1: Math.max(selected.c0, selected.c1),
  };
};

/** Expand a whole-row/-column source to the logical destination band. Excel
 * accepts a single anchor cell only on the first row/column of the sheet;
 * another row/column is outside the band and is rejected. */
const destinationBandFor = (
  state: State,
  snap: ClipboardSnapshot,
  transpose: boolean,
): { range: Range; axis: BandAxis } | null => {
  if (transpose) return null;
  const logical = logicalRangeFor(snap);
  const axis = bandAxisFor(logical);
  if (!axis) return null;
  const selected = normalizedSelection(state);
  if (selected.sheet !== state.selection.active.sheet) return null;
  if (axis === 'column') {
    const firstRowOnly = selected.r0 === 0 && selected.r1 === 0;
    const fullRows = selected.r0 === 0 && selected.r1 >= MAX_ROW;
    if (!firstRowOnly && !fullRows) return null;
    const selectedWidth = selected.c1 - selected.c0 + 1;
    const logicalWidth = logical.c1 - logical.c0 + 1;
    const width =
      snap.mode === 'copy' && selectedWidth >= logicalWidth && selectedWidth % logicalWidth === 0
        ? selectedWidth
        : logicalWidth;
    if (selected.c0 + width - 1 > MAX_COL) return null;
    return {
      axis,
      range: {
        sheet: selected.sheet,
        r0: 0,
        c0: selected.c0,
        r1: MAX_ROW,
        c1: selected.c0 + width - 1,
      },
    };
  }
  const firstColOnly = selected.c0 === 0 && selected.c1 === 0;
  const fullCols = selected.c0 === 0 && selected.c1 >= MAX_COL;
  if (!firstColOnly && !fullCols) return null;
  const selectedHeight = selected.r1 - selected.r0 + 1;
  const logicalHeight = logical.r1 - logical.r0 + 1;
  const height =
    snap.mode === 'copy' && selectedHeight >= logicalHeight && selectedHeight % logicalHeight === 0
      ? selectedHeight
      : logicalHeight;
  if (selected.r0 + height - 1 > MAX_ROW) return null;
  return {
    axis,
    range: {
      sheet: selected.sheet,
      r0: selected.r0,
      c0: 0,
      r1: selected.r0 + height - 1,
      c1: MAX_COL,
    },
  };
};

const bandMaterializedCellCount = (
  snap: ClipboardSnapshot,
  destination: Range,
  axis: BandAxis,
): number =>
  snap.rows *
  snap.cols *
  Math.ceil(
    (axis === 'column'
      ? destination.c1 - destination.c0 + 1
      : destination.r1 - destination.r0 + 1) /
      (axis === 'column'
        ? logicalRangeFor(snap).c1 - logicalRangeFor(snap).c0 + 1
        : logicalRangeFor(snap).r1 - logicalRangeFor(snap).r0 + 1),
  );

/** Enumerate only the bounded source payload projected onto a whole-band
 * destination. The full logical axis is intentionally never materialized. */
export function materializedPasteCells(
  snap: ClipboardSnapshot,
  destination: Range,
  transpose = false,
): MaterializedPasteCell[] | null {
  if (transpose) return null;
  const logical = logicalRangeFor(snap);
  const axis = bandAxisFor(logical);
  if (!axis || bandAxisFor(destination) !== axis) return null;
  if (bandMaterializedCellCount(snap, destination, axis) > MAX_PASTE_CELLS) return null;
  const out: MaterializedPasteCell[] = [];
  const seen = new Set<string>();
  const add = (row: number, col: number, sourceRowIndex: number, sourceColIndex: number): void => {
    if (
      row < destination.r0 ||
      row > destination.r1 ||
      col < destination.c0 ||
      col > destination.c1
    ) {
      return;
    }
    const key = `${row}:${col}`;
    if (seen.has(key)) return;
    seen.add(key);
    out.push({
      row,
      col,
      sourceRow: snap.range.r0 + sourceRowIndex,
      sourceCol: snap.range.c0 + sourceColIndex,
      sourceRowIndex,
      sourceColIndex,
    });
  };
  if (axis === 'column') {
    const logicalWidth = logical.c1 - logical.c0 + 1;
    for (let baseCol = destination.c0; baseCol <= destination.c1; baseCol += logicalWidth) {
      for (let sr = 0; sr < snap.rows; sr += 1) {
        for (let sc = 0; sc < snap.cols; sc += 1) {
          add(
            destination.r0 + snap.range.r0 + sr - logical.r0,
            baseCol + snap.range.c0 + sc - logical.c0,
            sr,
            sc,
          );
        }
      }
    }
  } else {
    const logicalHeight = logical.r1 - logical.r0 + 1;
    for (let baseRow = destination.r0; baseRow <= destination.r1; baseRow += logicalHeight) {
      for (let sr = 0; sr < snap.rows; sr += 1) {
        for (let sc = 0; sc < snap.cols; sc += 1) {
          add(
            baseRow + snap.range.r0 + sr - logical.r0,
            destination.c0 + snap.range.c0 + sc - logical.c0,
            sr,
            sc,
          );
        }
      }
    }
  }
  return out;
}

const bandProtectionAllows = (state: State, range: Range): boolean => {
  if (!state.protection.protectedSheets.has(range.sheet)) return true;
  // A whole-band operation is atomic. For a protected sheet an explicit
  // allowed-edit interval must cover the complete band; sparse unlocked cells
  // cannot prove that the omitted cells are writable without enumerating the
  // million-cell axis.
  return state.protection.allowedEditRanges.some((entry) => rangeContains(entry.range, range));
};

const removeIntersectingMerges = (
  byAnchor: Map<string, Range>,
  byCell: Map<string, string>,
  range: Range,
): void => {
  for (const [anchorKey, merge] of byAnchor) {
    if (!rangesIntersect(merge, range)) continue;
    byAnchor.delete(anchorKey);
    for (let row = merge.r0; row <= merge.r1; row += 1) {
      for (let col = merge.c0; col <= merge.c1; col += 1) {
        byCell.delete(addrKey({ sheet: merge.sheet, row, col }));
      }
    }
  }
};

const addMerge = (
  byAnchor: Map<string, Range>,
  byCell: Map<string, string>,
  range: Range,
): void => {
  const anchorKey = addrKey({ sheet: range.sheet, row: range.r0, col: range.c0 });
  byAnchor.set(anchorKey, range);
  for (let row = range.r0; row <= range.r1; row += 1) {
    for (let col = range.c0; col <= range.c1; col += 1) {
      if (row === range.r0 && col === range.c0) continue;
      byCell.set(addrKey({ sheet: range.sheet, row, col }), anchorKey);
    }
  }
};

const translateMerges = (
  snap: ClipboardSnapshot,
  destination: Range,
  transpose: boolean,
): Range[] | null => {
  const out: Range[] = [];
  const logical = logicalRangeFor(snap);
  const bandAxis = !transpose ? bandAxisFor(logical) : null;
  if (bandAxis && bandAxisFor(destination) === bandAxis) {
    const addTranslated = (
      baseRow: number,
      baseCol: number,
      merge: NonNullable<ClipboardSnapshot['merges']>[number],
    ): boolean => {
      if (
        !Number.isInteger(merge.r0) ||
        !Number.isInteger(merge.c0) ||
        !Number.isInteger(merge.r1) ||
        !Number.isInteger(merge.c1) ||
        merge.r0 < 0 ||
        merge.c0 < 0 ||
        merge.r1 < merge.r0 ||
        merge.c1 < merge.c0 ||
        merge.r1 >= snap.rows ||
        merge.c1 >= snap.cols ||
        (merge.sheet !== undefined && merge.sheet !== snap.range.sheet)
      ) {
        return false;
      }
      const sourceRow0 = snap.range.r0 + merge.r0;
      const sourceCol0 = snap.range.c0 + merge.c0;
      const sourceRow1 = snap.range.r0 + merge.r1;
      const sourceCol1 = snap.range.c0 + merge.c1;
      const translated: Range = {
        sheet: destination.sheet,
        r0: baseRow + sourceRow0 - logical.r0,
        c0: baseCol + sourceCol0 - logical.c0,
        r1: baseRow + sourceRow1 - logical.r0,
        c1: baseCol + sourceCol1 - logical.c0,
      };
      if (rangeContains(destination, translated)) out.push(translated);
      return true;
    };
    if (bandAxis === 'column') {
      const logicalWidth = logical.c1 - logical.c0 + 1;
      for (let baseCol = destination.c0; baseCol <= destination.c1; baseCol += logicalWidth) {
        for (const merge of snap.merges ?? []) {
          if (!addTranslated(destination.r0, baseCol, merge)) return null;
        }
      }
    } else {
      const logicalHeight = logical.r1 - logical.r0 + 1;
      for (let baseRow = destination.r0; baseRow <= destination.r1; baseRow += logicalHeight) {
        for (const merge of snap.merges ?? []) {
          if (!addTranslated(baseRow, destination.c0, merge)) return null;
        }
      }
    }
    return out;
  }
  const tileRows = transpose ? snap.cols : snap.rows;
  const tileCols = transpose ? snap.rows : snap.cols;
  for (let tileRow = destination.r0; tileRow <= destination.r1; tileRow += tileRows) {
    for (let tileCol = destination.c0; tileCol <= destination.c1; tileCol += tileCols) {
      for (const merge of snap.merges ?? []) {
        if (
          !Number.isInteger(merge.r0) ||
          !Number.isInteger(merge.c0) ||
          !Number.isInteger(merge.r1) ||
          !Number.isInteger(merge.c1) ||
          merge.r0 < 0 ||
          merge.c0 < 0 ||
          merge.r1 < merge.r0 ||
          merge.c1 < merge.c0 ||
          merge.r1 >= snap.rows ||
          merge.c1 >= snap.cols ||
          (merge.sheet !== undefined && merge.sheet !== snap.range.sheet)
        ) {
          return null;
        }
        out.push(
          transpose
            ? {
                sheet: destination.sheet,
                r0: tileRow + merge.c0,
                c0: tileCol + merge.r0,
                r1: tileRow + merge.c1,
                c1: tileCol + merge.r1,
              }
            : {
                sheet: destination.sheet,
                r0: tileRow + merge.r0,
                c0: tileCol + merge.c0,
                r1: tileRow + merge.r1,
                c1: tileCol + merge.c1,
              },
        );
      }
    }
  }
  return out;
};

const destinationRangeFor = (origin: Addr, rows: number, cols: number): Range | null => {
  if (rows <= 0 || cols <= 0) return null;
  const r1 = origin.row + rows - 1;
  const c1 = origin.col + cols - 1;
  if (origin.row < 0 || origin.col < 0 || r1 > MAX_ROW || c1 > MAX_COL) return null;
  return { sheet: origin.sheet, r0: origin.row, c0: origin.col, r1, c1 };
};

/**
 * Resolve the destination footprint for an internal clipboard paste.
 *
 * A copied matrix repeats only when the current primary selection is a
 * normalized, exact multiple of the (possibly transposed) source tile. This
 * includes the full logical footprint selected by a merged-cell click; the
 * paste path preserves that merge for a scalar source.
 */
export function resolvePasteDestination(
  state: State,
  snap: ClipboardSnapshot,
  transpose = false,
): Range | null {
  const tileRows = transpose ? snap.cols : snap.rows;
  const tileCols = transpose ? snap.rows : snap.cols;
  if (!Number.isInteger(tileRows) || !Number.isInteger(tileCols)) return null;
  if (tileRows <= 0 || tileCols <= 0) return null;

  const band = destinationBandFor(state, snap, transpose);
  if (bandAxisFor(logicalRangeFor(snap)) && !band) return null;
  if (band) {
    if (bandMaterializedCellCount(snap, band.range, band.axis) > MAX_PASTE_CELLS) return null;
    return band.range;
  }

  const active = state.selection.active;
  let origin = { row: active.row, col: active.col };
  let rows = tileRows;
  let cols = tileCols;
  const selected = state.selection.range;
  const normalized = {
    sheet: selected.sheet,
    r0: Math.min(selected.r0, selected.r1),
    c0: Math.min(selected.c0, selected.c1),
    r1: Math.max(selected.r0, selected.r1),
    c1: Math.max(selected.c0, selected.c1),
  };
  const selectedRows = normalized.r1 - normalized.r0 + 1;
  const selectedCols = normalized.c1 - normalized.c0 + 1;
  const canRepeat =
    snap.mode === 'copy' &&
    normalized.sheet === active.sheet &&
    normalized.r0 >= 0 &&
    normalized.c0 >= 0 &&
    selectedRows >= tileRows &&
    selectedCols >= tileCols &&
    selectedRows % tileRows === 0 &&
    selectedCols % tileCols === 0;
  if (canRepeat) {
    origin = { row: normalized.r0, col: normalized.c0 };
    rows = selectedRows;
    cols = selectedCols;
  }
  const destination = destinationRangeFor({ sheet: active.sheet, ...origin }, rows, cols);
  if (!destination || rows * cols > MAX_PASTE_CELLS) return null;
  return destination;
}

/**
 * Validate every cell and merge touched by a paste before the first mutation.
 * Excel rejects a matrix that would split a destination merge; allowing the
 * write and repairing the merge afterward loses data and makes undo partial.
 */
const preflight = (
  state: State,
  snap: ClipboardSnapshot,
  destination: Range,
  what: PasteWhat,
  transpose: boolean,
): { destination: Range; translatedMerges: Range[]; preservedMerges: Range[] } | null => {
  const band = bandAxisFor(logicalRangeFor(snap));
  const bandDestination = band && bandAxisFor(destination) === band;
  if (bandDestination) {
    // Do not enumerate a million-cell whole-band range. The operation is
    // atomic on a protected sheet and therefore requires one allowed-edit
    // interval covering the complete band.
    if (!bandProtectionAllows(state, destination)) return null;
  } else {
    for (let row = destination.r0; row <= destination.r1; row += 1) {
      for (let col = destination.c0; col <= destination.c1; col += 1) {
        if (!isCellWritable(state, { sheet: destination.sheet, row, col })) return null;
      }
    }
  }

  const source = snap.range;
  if (source.r0 < 0 || source.c0 < 0 || source.r1 > MAX_ROW || source.c1 > MAX_COL) {
    return null;
  }
  if (snap.mode === 'cut') {
    const sourceLogical = logicalRangeFor(snap);
    if (bandAxisFor(sourceLogical)) {
      if (!bandProtectionAllows(state, sourceLogical)) return null;
    } else {
      for (let row = source.r0; row <= source.r1; row += 1) {
        for (let col = source.c0; col <= source.c1; col += 1) {
          if (!isCellWritable(state, { sheet: source.sheet, row, col })) return null;
        }
      }
    }
  }

  // Any partially covered merge is ambiguous for a matrix paste. A merge that
  // is fully contained by the destination can be replaced by the source
  // topology (or removed for an unmerged All/Formats paste).
  for (const merge of state.merges.byAnchor.values()) {
    if (merge.sheet !== destination.sheet || !rangesIntersect(merge, destination)) continue;
    if (!rangeContains(destination, merge)) return null;
  }
  if (snap.mode === 'cut') {
    for (const merge of state.merges.byAnchor.values()) {
      if (merge.sheet !== source.sheet || !rangesIntersect(merge, source)) continue;
      if (!rangeContains(source, merge)) return null;
    }
  }

  const translatedMerges = wantsFormats(what) ? translateMerges(snap, destination, transpose) : [];
  if (translatedMerges === null) return null;
  for (const merge of translatedMerges) {
    if (!rangeContains(destination, merge)) return null;
  }
  // A scalar paste into the exact logical footprint of a merged cell updates
  // its anchor and keeps the merge. Non-scalar tiles intentionally continue
  // through the normal topology replacement path (Excel unmerges those
  // covered cells unless the source carries a matching merge).
  const preservedMerges =
    snap.rows === 1 && snap.cols === 1
      ? [...state.merges.byAnchor.values()].filter((merge) => sameRange(merge, destination))
      : [];
  return { destination, translatedMerges, preservedMerges };
};

const commentText = (format: CellFormat | undefined): string | null =>
  typeof format?.comment === 'string' && format.comment.length > 0 ? format.comment : null;

const commentAuthor = (format: CellFormat | undefined): string | null =>
  commentText(format) && typeof format?.commentAuthor === 'string' ? format.commentAuthor : null;

/** Bring engine-only notes into the store before a format/comment history
 * snapshot is captured. The engine is authoritative for persisted comments,
 * while the store remains the source used by the history helpers. */
const hydrateEngineComments = (
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  ranges: readonly Range[],
): void => {
  if (!wb.capabilities.commentsEnumerable || ranges.length === 0) return;
  const sheets = new Set(ranges.map((range) => range.sheet));
  const comments: Array<{ sheet: number; row: number; col: number; author: string; text: string }> =
    [];
  for (const sheet of sheets) {
    for (const comment of wb.getComments(sheet)) {
      if (!comment.text) continue;
      if (
        ranges.some(
          (range) =>
            range.sheet === sheet &&
            comment.row >= range.r0 &&
            comment.row <= range.r1 &&
            comment.col >= range.c0 &&
            comment.col <= range.c1,
        )
      ) {
        comments.push({ sheet, ...comment });
      }
    }
  }
  if (comments.length === 0) return;
  store.setState((s) => {
    const formats = new Map(s.format.formats);
    for (const comment of comments) {
      const key = addrKey({ sheet: comment.sheet, row: comment.row, col: comment.col });
      formats.set(key, {
        ...formats.get(key),
        comment: comment.text,
        commentAuthor: comment.author || undefined,
      });
    }
    return { ...s, format: { ...s.format, formats } };
  });
};

const formatWithoutComment = (format: CellFormat | undefined): CellFormat | null => {
  if (!format) return null;
  const next = { ...format };
  delete next.comment;
  delete next.commentAuthor;
  return next;
};

const formatCommentsOnly = (format: CellFormat | undefined): CellFormat | null => {
  const text = commentText(format);
  if (!text) return null;
  const author = commentAuthor(format);
  return author ? { comment: text, commentAuthor: author } : { comment: text };
};

const clearSourceFormats = (store: SpreadsheetStore, source: Range): void => {
  store.setState((s) => {
    const formats = new Map(s.format.formats);
    const area = (source.r1 - source.r0 + 1) * (source.c1 - source.c0 + 1);
    const clear = (key: string): void => {
      const comments = formatCommentsOnly(formats.get(key));
      if (comments) formats.set(key, comments);
      else formats.delete(key);
    };
    if (area > MAX_PASTE_CELLS) {
      for (const key of formats.keys()) {
        const [sheetRaw, rowRaw, colRaw] = key.split(':');
        const sheet = Number(sheetRaw);
        const row = Number(rowRaw);
        const col = Number(colRaw);
        if (
          sheet === source.sheet &&
          row >= source.r0 &&
          row <= source.r1 &&
          col >= source.c0 &&
          col <= source.c1
        ) {
          clear(key);
        }
      }
    } else {
      for (let row = source.r0; row <= source.r1; row += 1) {
        for (let col = source.c0; col <= source.c1; col += 1) {
          clear(addrKey({ sheet: source.sheet, row, col }));
        }
      }
    }
    return { ...s, format: { ...s.format, formats } };
  });
};

const clearSourceValues = (wb: WorkbookHandle, source: Range): void => {
  // `cellFormula`/`getValue` each scan a sheet in the engine. Materialize the
  // source sheet once so a large cut remains linear in the sheet's populated
  // cells rather than quadratic in the source rectangle.
  const sourceCells = Array.from(wb.physicalCells(source.sheet)).filter(
    (cell) =>
      cell.addr.row >= source.r0 &&
      cell.addr.row <= source.r1 &&
      cell.addr.col >= source.c0 &&
      cell.addr.col <= source.c1 &&
      (cell.formula !== null || cell.value.kind !== 'blank'),
  );
  wb.withBatchedRecalc(() => {
    for (const cell of sourceCells) wb.setBlank(cell.addr);
  });
};

const mutateMergeSlice = (
  store: SpreadsheetStore,
  sheet: number,
  removals: readonly Range[],
  additions: readonly Range[],
): void => {
  store.setState((s) => {
    const byAnchor = new Map(s.merges.byAnchor);
    const byCell = new Map(s.merges.byCell);
    for (const range of removals) {
      if (range.sheet === sheet) removeIntersectingMerges(byAnchor, byCell, range);
    }
    for (const range of additions) {
      if (range.sheet === sheet) {
        removeIntersectingMerges(byAnchor, byCell, range);
        addMerge(byAnchor, byCell, range);
      }
    }
    return { ...s, merges: { byAnchor, byCell } };
  });
};

const consumeCutMarquee = (store: SpreadsheetStore): void => {
  mutators.setCopyRange(store, null);
  mutators.setCopyRanges(store, null);
};

/**
 * Apply a clipboard snapshot to the destination starting at `state.selection.active`,
 * filtered by the spreadsheet-style "Paste Special" options. Returns the range that was
 * actually written. Caller is responsible for refreshing the cached cell map.
 */
export function pasteSpecial(
  state: State,
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  snap: ClipboardSnapshot,
  opt: PasteSpecialOptions,
  history: History | null = null,
): PasteSpecialResult | null {
  // A cut owns the source payload and clears it only after a legal move has
  // passed preflight. Paste Special variants would omit part of that payload
  // (or change its geometry), so accepting them would destroy source data.
  // Excel keeps those cut variants unavailable; reject them atomically.
  if (
    snap.mode === 'cut' &&
    (opt.what !== 'all' || opt.operation !== 'none' || opt.skipBlanks || opt.transpose)
  ) {
    return null;
  }

  const beforeHydration = store.getState();
  const resolvedDestination = resolvePasteDestination(beforeHydration, snap, opt.transpose);
  if (!resolvedDestination) return null;
  const sourceLogical = logicalRangeFor(snap);
  hydrateEngineComments(store, wb, [
    resolvedDestination,
    ...(snap.mode === 'cut' ? [sourceLogical] : []),
  ]);
  const current = store.getState();
  const checked = preflight(current, snap, resolvedDestination, opt.what, opt.transpose);
  if (!checked) return null;

  // Excel treats cutting and pasting back onto the exact source rectangle as a
  // no-op while consuming the cut marquee. In particular, do not clear and
  // rewrite the source, which would manufacture an unnecessary undo entry.
  if (snap.mode === 'cut' && !opt.transpose && sameRange(sourceLogical, checked.destination)) {
    consumeCutMarquee(store);
    return { writtenRange: checked.destination, skippedNonFiniteOperations: 0 };
  }

  // Pre-compute format patches we'll merge into the format slice in one pass.
  const formatWrites: { key: string; format: CellFormat | null }[] = [];
  const commentWrites: { addr: Addr; text: string | null; author: string | null }[] = [];
  let skippedNonFiniteOperations = 0;

  const applyTopology = wantsFormats(opt.what);
  const destination = checked.destination;
  const sheet = destination.sheet;
  const destRows = destination.r1 - destination.r0 + 1;
  const destCols = destination.c1 - destination.c0 + 1;
  const source = snap.range;
  const sourceCut = snap.mode === 'cut';
  const sheetNames = sourceCut ? sheetNamesFor(wb) : [];
  const bandAxis = bandAxisFor(sourceLogical);
  const bandDestination = bandAxis !== null && bandAxisFor(destination) === bandAxis;
  const bandCells = bandDestination
    ? materializedPasteCells(snap, destination, opt.transpose)
    : null;
  if (bandDestination && !bandCells) return null;
  const bandTargetKeys = bandCells
    ? new Set(bandCells.map((cell) => addrKey({ sheet, row: cell.row, col: cell.col })))
    : null;

  // Whole-band copies overwrite the sparse tail of the destination even
  // though the logical range is too large to enumerate. Blank-source and
  // operation variants preserve that tail, matching Paste Special semantics.
  if (bandCells && bandTargetKeys) {
    const clearValues =
      !opt.skipBlanks &&
      opt.operation === 'none' &&
      (wantsValues(opt.what) || wantsFormulas(opt.what));
    if (clearValues) {
      const stale = Array.from(wb.physicalCells(sheet)).filter(
        (cell) =>
          cell.addr.row >= destination.r0 &&
          cell.addr.row <= destination.r1 &&
          cell.addr.col >= destination.c0 &&
          cell.addr.col <= destination.c1 &&
          !bandTargetKeys.has(addrKey(cell.addr)) &&
          (cell.formula !== null || cell.value.kind !== 'blank') &&
          isCellWritable(current, cell.addr),
      );
      wb.withBatchedRecalc(() => {
        for (const cell of stale) wb.setBlank(cell.addr);
      });
    }
    if (!opt.skipBlanks && wantsFormats(opt.what)) {
      for (const [key, format] of current.format.formats) {
        const [formatSheetRaw, rowRaw, colRaw] = key.split(':');
        const formatSheet = Number(formatSheetRaw);
        const row = Number(rowRaw);
        const col = Number(colRaw);
        if (
          formatSheet === sheet &&
          row >= destination.r0 &&
          row <= destination.r1 &&
          col >= destination.c0 &&
          col <= destination.c1 &&
          !bandTargetKeys.has(key) &&
          Object.keys(format).length > 0
        ) {
          formatWrites.push({ key, format: null });
          if (opt.what === 'all' && commentText(format) !== null) {
            commentWrites.push({
              addr: { sheet, row, col },
              text: null,
              author: null,
            });
          }
        }
      }
    }
    if (!opt.skipBlanks && opt.what === 'all') {
      for (const comment of wb.getComments(sheet)) {
        const addr = { sheet, row: comment.row, col: comment.col };
        if (
          comment.row >= destination.r0 &&
          comment.row <= destination.r1 &&
          comment.col >= destination.c0 &&
          comment.col <= destination.c1
        ) {
          commentWrites.push({ addr, text: null, author: null });
        }
      }
    }
  }

  // Remove source/destination merges before writing cells. This is especially
  // important for overlapping cuts: the captured snapshot remains available
  // while the live source is cleared, so the destination can be rebuilt from
  // the original payload.
  const removalsBySheet = new Map<number, Range[]>();
  const addRemoval = (range: Range): void => {
    const list = removalsBySheet.get(range.sheet) ?? [];
    list.push(range);
    removalsBySheet.set(range.sheet, list);
  };
  if (sourceCut) addRemoval(sourceLogical);
  if (applyTopology && checked.preservedMerges.length === 0) addRemoval(destination);
  for (const [mergeSheet, removals] of removalsBySheet) {
    recordMergesChangeWithEngine(history, store, wb, mergeSheet, () => {
      mutateMergeSlice(store, mergeSheet, removals, []);
    });
  }

  if (sourceCut) {
    clearSourceValues(wb, sourceLogical);
    recordFormatChange(history, store, () => clearSourceFormats(store, sourceLogical));
  }

  wb.withBatchedRecalc(() => {
    const applyCell = (row: number, col: number, sr: number, sc: number): void => {
      const preservedBody = checked.preservedMerges.some(
        (merge) =>
          row >= merge.r0 &&
          row <= merge.r1 &&
          col >= merge.c0 &&
          col <= merge.c1 &&
          (row !== merge.r0 || col !== merge.c0),
      );
      if (preservedBody) return;
      const src = snap.cells[sr]?.[sc];
      if (!src) return;
      const isBlankSrc = src.value.kind === 'blank' && !src.formula && !src.format;
      if (opt.skipBlanks && isBlankSrc) return;

      const addr: Addr = { sheet, row, col };
      // Sheet protection — silently skip locked destinations (spreadsheet parity).
      if (!isCellWritable(state, addr)) return;

      if (opt.what === 'all') {
        const text = commentText(src.format);
        const existing = commentText(state.format.formats.get(addrKey(addr)));
        if (text !== null || existing !== null) {
          commentWrites.push({ addr, text, author: commentAuthor(src.format) });
        }
      }

      // Layer 1: values / formulas.
      // A Paste Special arithmetic operation always combines by VALUE. When one
      // is active it takes precedence over formula-pasting, using the source's
      // computed number even if the source cell is a formula — otherwise an
      // "Add" over a formula source silently pastes the formula and drops the
      // operation. Formats-only pastes carry no value, so they never
      // operate.
      const operating =
        opt.operation !== 'none' && (wantsValues(opt.what) || wantsFormulas(opt.what));
      const shouldPasteFormula = Boolean(src.formula && wantsFormulas(opt.what) && !operating);
      if (operating) {
        const srcNum = numericValue(src);
        if (srcNum !== null) {
          const dest = existingNumeric(state, sheet, row, col);
          const result = combine(opt.operation, dest, srcNum);
          if (Number.isFinite(result)) {
            wb.setNumber(addr, result);
          } else {
            wb.setError(addr, errorCodeForNonFiniteOperation(opt.operation));
            skippedNonFiniteOperations += 1;
          }
        }
        // Non-numeric source cells leave the destination unchanged (parity).
      } else if (shouldPasteFormula && src.formula) {
        // Cut moves cells and keeps references bound to the cells' original
        // sheets. The contextual transform also follows references that are
        // inside the moved block to their new destination coordinates. Copy
        // re-anchors relative refs by the paste offset.
        if (snap.mode === 'cut') {
          wb.setFormula(
            addr,
            adjustFormulaForCutPasteMove(
              src.formula,
              {
                r0: sourceLogical.r0,
                c0: sourceLogical.c0,
                r1: sourceLogical.r1,
                c1: sourceLogical.c1,
              },
              { r0: destination.r0, c0: destination.c0 },
              {
                sourceSheet: source.sheet,
                destinationSheet: destination.sheet,
                formulaSheet: source.sheet,
                outputSheet: destination.sheet,
                sheetNames,
              },
            ),
          );
        } else {
          const sourceRow = snap.range.r0 + sr;
          const sourceCol = snap.range.c0 + sc;
          wb.setFormula(addr, shiftFormulaRefs(src.formula, row - sourceRow, col - sourceCol));
        }
      } else if (wantsValues(opt.what) || wantsFormulas(opt.what)) {
        writeClipboardValue(wb, addr, src);
      }

      // Layer 2: formats
      const fmt = src.format;
      if (wantsFormats(opt.what)) {
        // A full "Formats" paste copies the source's *absence* of formatting
        // too: an unformatted source cell clears the destination format rather
        // than leaving a stale one behind (spreadsheet parity).
        formatWrites.push({
          key: addrKey(addr),
          format: formatWithoutComment(fmt),
        });
      } else if (wantsNumFmt(opt.what) && fmt?.numFmt) {
        // Number format only — cherry-pick.
        formatWrites.push({ key: addrKey(addr), format: { numFmt: fmt.numFmt } });
      }
    };
    if (bandCells) {
      for (const cell of bandCells) {
        applyCell(cell.row, cell.col, cell.sourceRowIndex, cell.sourceColIndex);
      }
    } else {
      for (let dr = 0; dr < destRows; dr += 1) {
        for (let dc = 0; dc < destCols; dc += 1) {
          applyCell(
            destination.r0 + dr,
            destination.c0 + dc,
            opt.transpose ? dc % snap.rows : dr % snap.rows,
            opt.transpose ? dr % snap.cols : dc % snap.cols,
          );
        }
      }
    }
  });

  if (formatWrites.length > 0) {
    recordFormatChange(history, store, () => {
      store.setState((s) => {
        const formats = new Map(s.format.formats);
        for (const { key, format } of formatWrites) {
          if (format === null) {
            const comments = formatCommentsOnly(formats.get(key));
            if (comments) formats.set(key, comments);
            else formats.delete(key);
          } else if (wantsFormats(opt.what)) {
            const comments = formatCommentsOnly(formats.get(key));
            formats.set(key, comments ? { ...format, ...comments } : format);
          } else {
            // For cherry-picks (number format only), merge into the existing cell format.
            const existing = formats.get(key) ?? {};
            formats.set(key, { ...existing, ...format });
          }
        }
        return { ...s, format: { ...s.format, formats } };
      });
    });
  }

  // Recreate the source merge topology only for All/Formats. Values and
  // Formulas intentionally leave destination topology untouched.
  if (applyTopology && checked.translatedMerges.length > 0) {
    const additionsBySheet = new Map<number, Range[]>();
    additionsBySheet.set(destination.sheet, checked.translatedMerges);
    for (const [mergeSheet, additions] of additionsBySheet) {
      recordMergesChangeWithEngine(history, store, wb, mergeSheet, () => {
        mutateMergeSlice(store, mergeSheet, [], additions);
      });
    }
  }

  // Comments/notes are part of Paste All, but Excel's Formats paste leaves
  // the destination note untouched. Keep comment history separate from the
  // ordinary format snapshot so the engine comment table and store restore in
  // lockstep on undo/redo.
  if (opt.what === 'all') {
    const sourceCommentsToClear: Addr[] = [];
    const tracked = new Map<string, Addr>();
    if (sourceCut) {
      if (bandAxisFor(sourceLogical)) {
        for (const [key, format] of current.format.formats) {
          const [commentSheetRaw, rowRaw, colRaw] = key.split(':');
          const commentSheet = Number(commentSheetRaw);
          const row = Number(rowRaw);
          const col = Number(colRaw);
          if (
            commentSheet !== sourceLogical.sheet ||
            row < sourceLogical.r0 ||
            row > sourceLogical.r1 ||
            col < sourceLogical.c0 ||
            col > sourceLogical.c1 ||
            commentText(format) === null
          ) {
            continue;
          }
          const addr = { sheet: commentSheet, row, col };
          sourceCommentsToClear.push(addr);
          tracked.set(addrKey(addr), addr);
        }
        for (const comment of wb.getComments(sourceLogical.sheet)) {
          if (
            comment.row < sourceLogical.r0 ||
            comment.row > sourceLogical.r1 ||
            comment.col < sourceLogical.c0 ||
            comment.col > sourceLogical.c1
          ) {
            continue;
          }
          const addr = {
            sheet: sourceLogical.sheet,
            row: comment.row,
            col: comment.col,
          };
          sourceCommentsToClear.push(addr);
          tracked.set(addrKey(addr), addr);
        }
      } else {
        for (let row = 0; row < snap.rows; row += 1) {
          for (let col = 0; col < snap.cols; col += 1) {
            const src = snap.cells[row]?.[col];
            if (!src || commentText(src.format) === null) continue;
            const addr = { sheet: source.sheet, row: source.r0 + row, col: source.c0 + col };
            sourceCommentsToClear.push(addr);
            tracked.set(addrKey(addr), addr);
          }
        }
      }
    }
    for (const write of commentWrites) tracked.set(addrKey(write.addr), write.addr);
    if (tracked.size > 0) {
      recordCommentChange(history, store, wb, [...tracked.values()], () => {
        for (const addr of sourceCommentsToClear) {
          mutators.setCellFormat(store, addr, {
            comment: undefined,
            commentAuthor: undefined,
          });
          wb.setCommentEntry(addr.sheet, addr.row, addr.col, '', '');
        }
        for (const write of commentWrites) {
          mutators.setCellFormat(store, write.addr, {
            comment: write.text ?? undefined,
            commentAuthor: write.text ? (write.author ?? undefined) : undefined,
          });
          wb.setCommentEntry(
            write.addr.sheet,
            write.addr.row,
            write.addr.col,
            write.author ?? '',
            write.text ?? '',
          );
        }
      });
    }
  }
  if ((opt.what === 'all' || opt.what === 'formats') && bandDestination) {
    const rowHeights = snap.rowHeights;
    const colWidths = snap.colWidths;
    const logicalRows = sourceLogical.r1 - sourceLogical.r0 + 1;
    const logicalCols = sourceLogical.c1 - sourceLogical.c0 + 1;
    if (bandAxis === 'row' && rowHeights) {
      recordLayoutChangeWithEngine(history, store, wb, () => {
        store.setState((s) => {
          const next = new Map(s.layout.rowHeights);
          for (const row of [...next.keys()]) {
            if (row >= destination.r0 && row <= destination.r1) next.delete(row);
          }
          for (const [offset, height] of rowHeights) {
            if (!Number.isInteger(offset) || offset < 0 || !Number.isFinite(height)) continue;
            for (let row = destination.r0 + offset; row <= destination.r1; row += logicalRows) {
              if (height !== s.layout.defaultRowHeight) next.set(row, height);
            }
          }
          return { ...s, layout: { ...s.layout, rowHeights: next } };
        });
      });
    }
    if (bandAxis === 'column' && colWidths) {
      recordLayoutChangeWithEngine(history, store, wb, () => {
        store.setState((s) => {
          const next = new Map(s.layout.colWidths);
          for (const col of [...next.keys()]) {
            if (col >= destination.c0 && col <= destination.c1) next.delete(col);
          }
          for (const [offset, width] of colWidths) {
            if (!Number.isInteger(offset) || offset < 0 || !Number.isFinite(width)) continue;
            for (let col = destination.c0 + offset; col <= destination.c1; col += logicalCols) {
              if (width !== s.layout.defaultColWidth) next.set(col, width);
            }
          }
          return { ...s, layout: { ...s.layout, colWidths: next } };
        });
      });
    }
  }
  if (skippedNonFiniteOperations > 0) {
    console.warn(
      `formulon-cell: paste special wrote ${skippedNonFiniteOperations} non-finite arithmetic result(s) as static error value(s)`,
    );
  }

  // Move active selection to the written range.
  const writtenRange: Range = { ...destination };
  mutators.setActive(store, { sheet, row: writtenRange.r0, col: writtenRange.c0 });
  if (bandDestination) {
    mutators.setRange(store, writtenRange);
  } else if (writtenRange.r0 !== writtenRange.r1 || writtenRange.c0 !== writtenRange.c1) {
    mutators.extendRangeTo(store, { sheet, row: writtenRange.r1, col: writtenRange.c1 });
  }
  if (snap.mode === 'cut') {
    updateExternalRefsForCutPaste(wb, sourceLogical, writtenRange);
    consumeCutMarquee(store);
  }
  return { writtenRange, skippedNonFiniteOperations };
}

function updateExternalRefsForCutPaste(
  wb: WorkbookHandle,
  source: Range,
  writtenRange: Range,
): void {
  const sheetNames = sheetNamesFor(wb);
  wb.withBatchedRecalc(() => {
    for (let formulaSheet = 0; formulaSheet < wb.sheetCount; formulaSheet += 1) {
      const context = {
        sourceSheet: source.sheet,
        destinationSheet: writtenRange.sheet,
        formulaSheet,
        outputSheet: formulaSheet,
        sheetNames,
      };
      for (const entry of wb.cells(formulaSheet)) {
        if (!entry.formula) continue;
        if (
          entry.addr.sheet === writtenRange.sheet &&
          entry.addr.row >= writtenRange.r0 &&
          entry.addr.row <= writtenRange.r1 &&
          entry.addr.col >= writtenRange.c0 &&
          entry.addr.col <= writtenRange.c1
        ) {
          continue;
        }
        const next = adjustFormulaForCutPasteMove(entry.formula, source, writtenRange, context);
        if (next !== entry.formula) wb.setFormula(entry.addr, next);
      }
    }
  });
}

function sheetNamesFor(wb: WorkbookHandle): string[] {
  const names: string[] = [];
  for (let sheet = 0; sheet < wb.sheetCount; sheet += 1) names.push(wb.sheetName(sheet));
  return names;
}
