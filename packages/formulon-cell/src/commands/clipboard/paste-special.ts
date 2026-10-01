import { addrKey } from '../../engine/address.js';
import type { Addr, Range } from '../../engine/types.js';
import type { WorkbookHandle } from '../../engine/workbook-handle.js';
import { type CellFormat, mutators, type SpreadsheetStore, type State } from '../../store/store.js';
import { recordCommentChange } from '../comment.js';
import { adjustFormulaForCutPasteMove } from '../formula-refs.js';
import type { History } from '../history.js';
import { recordFormatChange, recordMergesChangeWithEngine } from '../history.js';
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

const rangesIntersect = (a: Range, b: Range): boolean =>
  a.sheet === b.sheet && !(a.r1 < b.r0 || a.r0 > b.r1 || a.c1 < b.c0 || a.c0 > b.c1);

const rangeContains = (outer: Range, inner: Range): boolean =>
  outer.sheet === inner.sheet &&
  inner.r0 >= outer.r0 &&
  inner.c0 >= outer.c0 &&
  inner.r1 <= outer.r1 &&
  inner.c1 <= outer.c1;

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
  origin: Addr,
  transpose: boolean,
): Range[] | null => {
  const out: Range[] = [];
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
            sheet: origin.sheet,
            r0: origin.row + merge.c0,
            c0: origin.col + merge.r0,
            r1: origin.row + merge.c1,
            c1: origin.col + merge.r1,
          }
        : {
            sheet: origin.sheet,
            r0: origin.row + merge.r0,
            c0: origin.col + merge.c0,
            r1: origin.row + merge.r1,
            c1: origin.col + merge.c1,
          },
    );
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
 * Validate every cell and merge touched by a paste before the first mutation.
 * Excel rejects a matrix that would split a destination merge; allowing the
 * write and repairing the merge afterward loses data and makes undo partial.
 */
const preflight = (
  state: State,
  snap: ClipboardSnapshot,
  origin: Addr,
  destRows: number,
  destCols: number,
  what: PasteWhat,
  transpose: boolean,
): { destination: Range; translatedMerges: Range[] } | null => {
  const destination = destinationRangeFor(origin, destRows, destCols);
  if (!destination) return null;

  for (let row = destination.r0; row <= destination.r1; row += 1) {
    for (let col = destination.c0; col <= destination.c1; col += 1) {
      if (!isCellWritable(state, { sheet: destination.sheet, row, col })) return null;
    }
  }

  const source = snap.range;
  if (source.r0 < 0 || source.c0 < 0 || source.r1 > MAX_ROW || source.c1 > MAX_COL) {
    return null;
  }
  if (snap.mode === 'cut') {
    for (let row = source.r0; row <= source.r1; row += 1) {
      for (let col = source.c0; col <= source.c1; col += 1) {
        if (!isCellWritable(state, { sheet: source.sheet, row, col })) return null;
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

  const translatedMerges = wantsFormats(what) ? translateMerges(snap, origin, transpose) : [];
  if (translatedMerges === null) return null;
  for (const merge of translatedMerges) {
    if (!rangeContains(destination, merge)) return null;
  }
  return { destination, translatedMerges };
};

const commentText = (format: CellFormat | undefined): string | null =>
  typeof format?.comment === 'string' && format.comment.length > 0 ? format.comment : null;

const commentAuthor = (format: CellFormat | undefined): string | null =>
  commentText(format) && typeof format?.commentAuthor === 'string' ? format.commentAuthor : null;

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
    for (let row = source.r0; row <= source.r1; row += 1) {
      for (let col = source.c0; col <= source.c1; col += 1) {
        const key = addrKey({ sheet: source.sheet, row, col });
        const comments = formatCommentsOnly(formats.get(key));
        if (comments) formats.set(key, comments);
        else formats.delete(key);
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
  const origin: Addr = state.selection.active;
  const sheet = origin.sheet;

  const destRows = opt.transpose ? snap.cols : snap.rows;
  const destCols = opt.transpose ? snap.rows : snap.cols;

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

  const current = store.getState();
  const checked = preflight(current, snap, origin, destRows, destCols, opt.what, opt.transpose);
  if (!checked) return null;

  // Excel treats cutting and pasting back onto the exact source rectangle as a
  // no-op while consuming the cut marquee. In particular, do not clear and
  // rewrite the source, which would manufacture an unnecessary undo entry.
  if (
    snap.mode === 'cut' &&
    !opt.transpose &&
    snap.range.sheet === checked.destination.sheet &&
    snap.range.r0 === checked.destination.r0 &&
    snap.range.c0 === checked.destination.c0 &&
    snap.range.r1 === checked.destination.r1 &&
    snap.range.c1 === checked.destination.c1
  ) {
    consumeCutMarquee(store);
    return { writtenRange: checked.destination, skippedNonFiniteOperations: 0 };
  }

  // Pre-compute format patches we'll merge into the format slice in one pass.
  const formatWrites: { key: string; format: CellFormat | null }[] = [];
  const commentWrites: { addr: Addr; text: string | null; author: string | null }[] = [];
  let skippedNonFiniteOperations = 0;

  const applyTopology = wantsFormats(opt.what);
  const destination = checked.destination;
  const source = snap.range;
  const sourceCut = snap.mode === 'cut';
  const sheetNames = sourceCut ? sheetNamesFor(wb) : [];

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
  if (sourceCut) addRemoval(source);
  if (applyTopology) addRemoval(destination);
  for (const [mergeSheet, removals] of removalsBySheet) {
    recordMergesChangeWithEngine(history, store, wb, mergeSheet, () => {
      mutateMergeSlice(store, mergeSheet, removals, []);
    });
  }

  if (sourceCut) {
    clearSourceValues(wb, source);
    recordFormatChange(history, store, () => clearSourceFormats(store, source));
  }

  wb.withBatchedRecalc(() => {
    for (let dr = 0; dr < destRows; dr += 1) {
      for (let dc = 0; dc < destCols; dc += 1) {
        const sr = opt.transpose ? dc : dr;
        const sc = opt.transpose ? dr : dc;
        const src = snap.cells[sr]?.[sc];
        if (!src) continue;
        const isBlankSrc = src.value.kind === 'blank' && !src.formula && !src.format;
        if (opt.skipBlanks && isBlankSrc) continue;

        const row = origin.row + dr;
        const col = origin.col + dc;
        const addr: Addr = { sheet, row, col };
        // Sheet protection — silently skip locked destinations (spreadsheet parity).
        if (!isCellWritable(state, addr)) continue;

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
                  r0: source.r0,
                  c0: source.c0,
                  r1: source.r1,
                  c1: source.c1,
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
  if (skippedNonFiniteOperations > 0) {
    console.warn(
      `formulon-cell: paste special wrote ${skippedNonFiniteOperations} non-finite arithmetic result(s) as static error value(s)`,
    );
  }

  // Move active selection to the written range.
  const writtenRange: Range = {
    sheet,
    r0: origin.row,
    c0: origin.col,
    r1: origin.row + destRows - 1,
    c1: origin.col + destCols - 1,
  };
  mutators.setActive(store, { sheet, row: writtenRange.r0, col: writtenRange.c0 });
  if (writtenRange.r0 !== writtenRange.r1 || writtenRange.c0 !== writtenRange.c1) {
    mutators.extendRangeTo(store, { sheet, row: writtenRange.r1, col: writtenRange.c1 });
  }
  if (snap.mode === 'cut') {
    updateExternalRefsForCutPaste(wb, snap.range, writtenRange);
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
