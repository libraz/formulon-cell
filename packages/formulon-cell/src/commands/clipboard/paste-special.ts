import { addrKey, colLetter, MAX_COL, MAX_ROW, parseAddrKey } from '../../engine/address.js';
import type { Addr, Range } from '../../engine/types.js';
import { writeCell } from '../../engine/value.js';
import type { CellPatch, WorkbookHandle } from '../../engine/workbook-handle.js';
import { styleElementFormatSource } from '../../render/grid/table-format.js';
import { addMergeToMaps, removeIntersectingMerges } from '../../store/merge-maps.js';
import {
  isWholeColumnRange,
  isWholeRowRange,
  rangeContainsRange,
  rangesIntersect,
  sameRange,
} from '../../store/selection-geometry.js';
import {
  type CellBorders,
  type CellFormat,
  type ConditionalRule,
  mutators,
  type SpreadsheetStore,
  type State,
} from '../../store/store.js';
import { recordCommentChange } from '../comment.js';
import { DEFAULT_COMMENT_AUTHOR } from '../comment-author.js';
import { blockDataTableEdit, blockTableFunction } from '../data-table-guard.js';
import { adjustFormulaForCutPasteMove } from '../formula-refs.js';
import type { History } from '../history.js';
import { blockPivotEdit } from '../pivot-guard.js';
import { pastePivotTableCopy, pivotTableAtRange } from '../pivot-table.js';
import { blockedByProtection, isCellWritable, isRangeInAllowedEditRange } from '../protection.js';
import { shiftFormulaRefs } from '../refs.js';
import {
  recordConditionalRulesChange,
  recordFormatChange,
  recordLayoutChangeWithEngine,
  recordMergesChangeWithEngine,
} from '../slice-history.js';
import { quoteSheetName } from '../sparkline.js';
import { applyPatchesAtomic } from './apply-patches-atomic.js';
import {
  bandAxisFor,
  logicalRangeFor,
  MAX_PASTE_CELLS,
  materializedPasteCells,
  resolvePasteDestination,
} from './paste-destination.js';
import type { ClipboardCell, ClipboardSnapshot } from './snapshot.js';
import { extendTablesForPaste } from './table-extension.js';
import { transposeFormulaRefs } from './transpose-refs.js';

export type { MaterializedPasteCell } from './paste-destination.js';
export { materializedPasteCells, resolvePasteDestination } from './paste-destination.js';

export type PasteWhat =
  | 'all'
  | 'all-using-source-theme'
  | 'all-except-borders'
  | 'all-merging-conditional-formats'
  | 'values'
  | 'formulas'
  | 'formats'
  | 'comments'
  | 'validation'
  | 'column-widths'
  | 'formulas-and-numfmt'
  | 'values-and-numfmt';

/** What a paste writes: a Paste Special kind, or the references Paste Link
 *  writes in place of the copied cells. */
type PasteKind = PasteWhat | 'link';

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

/** Cell-format fields a Formats paste leaves to the destination: data
 *  validation and the hyperlink (Excel pastes neither as a format). */
const FORMATS_PASTE_KEEPS = [
  'validation',
  'hyperlink',
  'hyperlinkDisplay',
  'hyperlinkTooltip',
  'hyperlinkLocation',
  'hyperlinkSpan',
] as const satisfies readonly (keyof CellFormat)[];

/** The kinds that paste every layer. A paste within one workbook shares its
 *  theme, so the source-theme variant pastes exactly what All does. */
const pastesAll = (what: PasteKind): boolean =>
  what === 'all' ||
  what === 'all-using-source-theme' ||
  what === 'all-except-borders' ||
  what === 'all-merging-conditional-formats';
const wantsValues = (what: PasteKind): boolean =>
  pastesAll(what) || what === 'values' || what === 'values-and-numfmt';
const wantsFormulas = (what: PasteKind): boolean =>
  pastesAll(what) || what === 'formulas' || what === 'formulas-and-numfmt';
const wantsFormats = (what: PasteKind): boolean => pastesAll(what) || what === 'formats';
const wantsNumFmt = (what: PasteKind): boolean =>
  wantsFormats(what) || what === 'values-and-numfmt' || what === 'formulas-and-numfmt';
const carriesComments = (what: PasteKind): boolean => pastesAll(what) || what === 'comments';

const bandProtectionAllows = (state: State, range: Range): boolean => {
  if (!state.protection.protectedSheets.has(range.sheet)) return true;
  // A whole-band operation is atomic. For a protected sheet an explicit
  // allowed-edit interval must cover the complete band; sparse unlocked cells
  // cannot prove that the omitted cells are writable without enumerating the
  // million-cell axis.
  return isRangeInAllowedEditRange(state, range);
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
      if (rangeContainsRange(destination, translated)) out.push(translated);
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

/**
 * Validate every cell and merge touched by a paste before the first mutation.
 * Excel rejects a matrix that would split a destination merge; allowing the
 * write and repairing the merge afterward loses data and makes undo partial.
 */
const preflight = (
  state: State,
  snap: ClipboardSnapshot,
  destination: Range,
  what: PasteKind,
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
    if (!rangeContainsRange(destination, merge)) return null;
  }
  if (snap.mode === 'cut') {
    for (const merge of state.merges.byAnchor.values()) {
      if (merge.sheet !== source.sheet || !rangesIntersect(merge, source)) continue;
      if (!rangeContainsRange(source, merge)) return null;
    }
  }

  const translatedMerges = wantsFormats(what) ? translateMerges(snap, destination, transpose) : [];
  if (translatedMerges === null) return null;
  for (const merge of translatedMerges) {
    if (!rangeContainsRange(destination, merge)) return null;
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
  const comments: Array<{
    sheet: number;
    row: number;
    col: number;
    author: string;
    text: string;
    visible?: boolean;
  }> = [];
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
        commentVisible: comment.visible || undefined,
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
  delete next.commentVisible;
  return next;
};

const formatCommentsOnly = (format: CellFormat | undefined): CellFormat | null => {
  const text = commentText(format);
  if (!text) return null;
  const author = commentAuthor(format);
  return {
    comment: text,
    ...(author ? { commentAuthor: author } : {}),
    ...(format?.commentVisible === true ? { commentVisible: true } : {}),
  };
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
        const addr = parseAddrKey(key);
        if (!addr) continue;
        const { sheet, row, col } = addr;
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
        addMergeToMaps(byAnchor, byCell, range);
      }
    }
    return { ...s, merges: { byAnchor, byCell } };
  });
};

/** A single copied cell pasted onto a merged cell's anchor fills the merge,
 *  which keeps its shape, as when the whole merged cell is selected. */
const scalarIntoMergedAnchor = (
  state: State,
  snap: ClipboardSnapshot,
  destination: Range | null,
): Range | null => {
  if (!destination || snap.rows !== 1 || snap.cols !== 1) return destination;
  if (destination.r0 !== destination.r1 || destination.c0 !== destination.c1) return destination;
  const merge = state.merges.byAnchor.get(
    addrKey({ sheet: destination.sheet, row: destination.r0, col: destination.c0 }),
  );
  return merge ? { ...merge } : destination;
};

/** The reference Paste Link writes for a source cell: relative A1, qualified
 *  by the source sheet's name when the link lands on another sheet. */
const linkFormula = (
  sourceSheet: number,
  destinationSheet: number,
  wb: WorkbookHandle,
  row: number,
  col: number,
): string => {
  const qualifier =
    sourceSheet === destinationSheet ? '' : `${quoteSheetName(wb.sheetName(sourceSheet))}!`;
  return `=${qualifier}${colLetter(col)}${row + 1}`;
};

/** Paste Special Column Widths: each destination column takes the width of
 *  the source column it lines up with, the copied block repeating across a
 *  wider destination. Cells are left as they are. */
const pasteColumnWidths = (
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  history: History | null,
  snap: ClipboardSnapshot,
  destination: Range,
): void => {
  const layout = store.getState().layout;
  const source = logicalRangeFor(snap);
  const sourceCols = source.c1 - source.c0 + 1;
  const widthAt = (offset: number): number =>
    snap.colWidths?.get(offset) ??
    layout.colWidths.get(source.c0 + offset) ??
    layout.defaultColWidth;
  const widths = Array.from(
    { length: Math.min(sourceCols, destination.c1 - destination.c0 + 1) },
    (_, i) => widthAt(i),
  );
  recordLayoutChangeWithEngine(history, store, wb, () => {
    store.setState((s) => {
      const next = new Map(s.layout.colWidths);
      for (let col = destination.c0; col <= destination.c1; col += 1) {
        const width = widths[(col - destination.c0) % sourceCols] ?? s.layout.defaultColWidth;
        if (width === s.layout.defaultColWidth) next.delete(col);
        else next.set(col, width);
      }
      return { ...s, layout: { ...s.layout, colWidths: next } };
    });
  });
};

/** The copied conditional-format rules placed on every pasted tile. */
const translatedConditionalRules = (
  snap: ClipboardSnapshot,
  destination: Range,
): ConditionalRule[] => {
  const rules = snap.conditional ?? [];
  const added: ConditionalRule[] = [];
  for (let r0 = destination.r0; r0 <= destination.r1; r0 += snap.rows) {
    for (let c0 = destination.c0; c0 <= destination.c1; c0 += snap.cols) {
      const dRow = r0 - snap.range.r0;
      const dCol = c0 - snap.range.c0;
      for (const rule of rules) {
        const copy = structuredClone(rule);
        copy.range = {
          sheet: destination.sheet,
          r0: rule.range.r0 + dRow,
          c0: rule.range.c0 + dCol,
          r1: rule.range.r1 + dRow,
          c1: rule.range.c1 + dCol,
        };
        if (copy.kind === 'formula') copy.formula = shiftFormulaRefs(copy.formula, dRow, dCol);
        added.push(copy);
      }
    }
  }
  return added;
};

/** All, Merging Conditional Formats: the copied rules land on every pasted
 *  tile ahead of the destination's own rules, which stay (Excel lists a
 *  merged cell's rules source first). */
const pasteMergedConditionalRules = (
  store: SpreadsheetStore,
  history: History | null,
  snap: ClipboardSnapshot,
  destination: Range,
): void => {
  const added = translatedConditionalRules(snap, destination);
  if (added.length === 0) return;
  recordConditionalRulesChange(history, store, () => {
    store.setState((s) => ({
      ...s,
      conditional: { ...s.conditional, rules: [...added, ...s.conditional.rules] },
    }));
  });
};

/** `rule` with `cut` taken out of its range: up to four rules over what is
 *  left, a formula rule re-anchored to each piece's top-left. */
const subtractFromRule = (rule: ConditionalRule, cut: Range): ConditionalRule[] => {
  const r = rule.range;
  if (r.sheet !== cut.sheet || !rangesIntersect(r, cut)) return [rule];
  const rows0 = Math.max(r.r0, cut.r0);
  const rows1 = Math.min(r.r1, cut.r1);
  const pieces: Range[] = [
    { ...r, r1: cut.r0 - 1 },
    { ...r, r0: cut.r1 + 1 },
    { ...r, r0: rows0, r1: rows1, c1: cut.c0 - 1 },
    { ...r, r0: rows0, r1: rows1, c0: cut.c1 + 1 },
  ].filter((p) => p.r0 <= p.r1 && p.c0 <= p.c1 && p.r0 >= r.r0 && p.r1 <= r.r1);
  return pieces.map((piece) => {
    const copy = structuredClone(rule);
    copy.range = piece;
    if (copy.kind === 'formula') {
      copy.formula = shiftFormulaRefs(copy.formula, piece.r0 - r.r0, piece.c0 - r.c0);
    }
    return copy;
  });
};

/** All and Formats replace the conditional formats of the pasted cells: the
 *  rules lose every `cleared` range (the destination, and a cut's source) and
 *  the copied rules take the destination (an uncovered source leaves none).
 *  Cells outside keep theirs. */
const replaceConditionalRules = (
  store: SpreadsheetStore,
  history: History | null,
  added: readonly ConditionalRule[],
  cleared: readonly Range[],
): void => {
  const current = store.getState().conditional.rules;
  const touched = current.some((rule) =>
    cleared.some((range) => rangesIntersect(rule.range, range)),
  );
  if (added.length === 0 && !touched) return;
  recordConditionalRulesChange(history, store, () => {
    store.setState((s) => ({
      ...s,
      conditional: {
        ...s.conditional,
        rules: [
          ...cleared.reduce<ConditionalRule[]>(
            (rules, range) => rules.flatMap((rule) => subtractFromRule(rule, range)),
            s.conditional.rules,
          ),
          ...added,
        ],
      },
    }));
  });
};

/** Source cells (row-major index into `snap.cells`) that a dynamic-array
 *  formula copied with them spills into. Excel pastes the anchor and lets it
 *  spill again: I1:I3 =SEQUENCE(3) pasted at K1 leaves K2:K3 to the spill. */
const spilledSourceCells = (wb: WorkbookHandle, snap: ClipboardSnapshot): Set<number> => {
  const out = new Set<number>();
  if (!wb.capabilities.spillInfo || snap.mode !== 'copy') return out;
  const { sheet, r0, c0, r1, c1 } = snap.range;
  for (let sr = 0; sr < snap.rows; sr += 1) {
    for (let sc = 0; sc < snap.cols; sc += 1) {
      const cell = snap.cells[sr]?.[sc];
      if (!cell || cell.formula || cell.value.kind === 'blank') continue;
      const spill = wb.spillInfo(sheet, r0 + sr, c0 + sc);
      if (!spill) continue;
      const anchorInside =
        spill.anchorRow >= r0 &&
        spill.anchorRow <= r1 &&
        spill.anchorCol >= c0 &&
        spill.anchorCol <= c1 &&
        (spill.anchorRow !== r0 + sr || spill.anchorCol !== c0 + sc);
      if (anchorInside) out.add(sr * snap.cols + sc);
    }
  }
  return out;
};

/** Make the written range the selection, its top-left cell active. */
const selectWritten = (store: SpreadsheetStore, written: Range, band: boolean): void => {
  const { sheet } = written;
  mutators.setActive(store, { sheet, row: written.r0, col: written.c0 });
  if (band) {
    mutators.setRange(store, written);
  } else if (written.r0 !== written.r1 || written.c0 !== written.c1) {
    mutators.extendRangeTo(store, { sheet, row: written.r1, col: written.c1 });
  }
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
  return applyPaste(state, store, wb, snap, opt, history);
}

/**
 * Paste Link: write a reference to each copied cell (`=A1`, relative, sheet
 * qualified when the source is another sheet) over the destination that a
 * paste of `snap` would take. Formats, notes and merges are left as they are.
 * A cut has nothing to link to and is refused.
 */
export function pasteLink(
  state: State,
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  snap: ClipboardSnapshot,
  history: History | null = null,
): PasteSpecialResult | null {
  if (snap.mode === 'cut') return null;
  return applyPaste(
    state,
    store,
    wb,
    snap,
    { what: 'link', operation: 'none', skipBlanks: false, transpose: false },
    history,
  );
}

type PasteOptions = Omit<PasteSpecialOptions, 'what'> & { what: PasteKind };

function applyPaste(
  state: State,
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  snap: ClipboardSnapshot,
  opt: PasteOptions,
  history: History | null,
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
  const resolvedDestination = scalarIntoMergedAnchor(
    beforeHydration,
    snap,
    resolvePasteDestination(beforeHydration, snap, opt.transpose),
  );
  if (!resolvedDestination) return null;
  // Column widths write no cell: a protected sheet takes them when it allows
  // formatting columns, whatever the destination cells' locks.
  if (opt.what === 'column-widths') {
    if (
      blockedByProtection(store, resolvedDestination.sheet, 'paste column widths', 'formatColumns')
    ) {
      return null;
    }
    pasteColumnWidths(store, wb, history, snap, resolvedDestination);
    selectWritten(store, resolvedDestination, false);
    return { writtenRange: { ...resolvedDestination }, skippedNonFiniteOperations: 0 };
  }
  const sourceLogical = logicalRangeFor(snap);
  hydrateEngineComments(store, wb, [
    resolvedDestination,
    ...(snap.mode === 'cut' ? [sourceLogical] : []),
  ]);
  const current = store.getState();
  const checked = preflight(current, snap, resolvedDestination, opt.what, opt.transpose);
  if (!checked) return null;
  if (blockPivotEdit(wb, checked.destination)) return null;
  // A cut would also clear the source, and a data table cannot move.
  const guarded =
    snap.mode === 'cut' ? [checked.destination, sourceLogical] : [checked.destination];
  if (blockDataTableEdit(wb, ...guarded)) return null;
  if (
    wantsFormulas(opt.what) &&
    blockTableFunction(
      wb,
      snap.cells.flatMap((row) => row.map((cell) => cell.formula)),
    )
  ) {
    return null;
  }

  // A whole PivotTable pastes as a second report rather than as its cells.
  if (snap.mode === 'copy' && opt.what === 'all' && opt.operation === 'none' && !opt.transpose) {
    const pivotIndex = pivotTableAtRange(wb, snap.range);
    if (pivotIndex !== null) {
      const written = pastePivotTableCopy(
        { store, workbook: wb, history },
        { sheet: snap.range.sheet, pivotIndex },
        {
          sheet: checked.destination.sheet,
          row: checked.destination.r0,
          col: checked.destination.c0,
        },
      );
      if (!written) return null;
      selectWritten(store, written, false);
      return { writtenRange: written, skippedNonFiniteOperations: 0 };
    }
  }

  // Excel treats cutting and pasting back onto the exact source rectangle as a
  // no-op while consuming the cut marquee. In particular, do not clear and
  // rewrite the source, which would manufacture an unnecessary undo entry.
  if (snap.mode === 'cut' && !opt.transpose && sameRange(sourceLogical, checked.destination)) {
    consumeCutMarquee(store);
    return { writtenRange: checked.destination, skippedNonFiniteOperations: 0 };
  }

  const spilled = opt.transpose ? new Set<number>() : spilledSourceCells(wb, snap);
  // A table or PivotTable style's boundary pastes as formats onto both cells it
  // separates: a header's bottom rule also tops the row below it, a total's top
  // rule also closes the row above it.
  const sourceStyleAt =
    opt.what === 'formats' && snap.range.sheet === current.data.sheetIndex
      ? styleElementFormatSource(current, wb)
      : null;
  const withMirroredStyleBorders = (format: CellFormat, sr: number, sc: number): CellFormat => {
    if (!sourceStyleAt) return format;
    const row = snap.range.r0 + sr;
    const col = snap.range.c0 + sc;
    const inside = (r: number, c: number): boolean =>
      r >= snap.range.r0 && r <= snap.range.r1 && c >= snap.range.c0 && c <= snap.range.c1;
    const facing = (r: number, c: number, side: keyof CellBorders) =>
      inside(r, c) ? sourceStyleAt(r, c)?.borders?.[side] : undefined;
    const sides: CellBorders = {};
    const top = format.borders?.top ? undefined : facing(row - 1, col, 'bottom');
    const bottom = format.borders?.bottom ? undefined : facing(row + 1, col, 'top');
    const left = format.borders?.left ? undefined : facing(row, col - 1, 'right');
    const right = format.borders?.right ? undefined : facing(row, col + 1, 'left');
    if (top) sides.top = top;
    if (bottom) sides.bottom = bottom;
    if (left) sides.left = left;
    if (right) sides.right = right;
    return Object.keys(sides).length > 0
      ? { ...format, borders: { ...format.borders, ...sides } }
      : format;
  };

  // Pre-compute format patches we'll merge into the format slice in one pass.
  const formatWrites: { key: string; format: CellFormat | null }[] = [];
  const commentWrites: {
    addr: Addr;
    text: string | null;
    author: string | null;
    visible?: boolean;
  }[] = [];
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
  const valuePatches: CellPatch[] | null =
    snap.mode === 'copy' &&
    pastesAll(opt.what) &&
    opt.operation === 'none' &&
    !opt.transpose &&
    history !== null
      ? []
      : null;
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
      if (valuePatches) {
        for (const cell of stale) {
          valuePatches.push({
            addr: cell.addr,
            value: { kind: 'blank' },
            formula: null,
          });
        }
      } else {
        wb.withBatchedRecalc(() => {
          for (const cell of stale) wb.setBlank(cell.addr);
        });
      }
    }
    if (!opt.skipBlanks && wantsFormats(opt.what)) {
      for (const [key, format] of current.format.formats) {
        const addr = parseAddrKey(key);
        if (!addr) continue;
        const { sheet: formatSheet, row, col } = addr;
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
          if (carriesComments(opt.what) && commentText(format) !== null) {
            commentWrites.push({
              addr: { sheet, row, col },
              text: null,
              author: null,
            });
          }
        }
      }
    }
    if (!opt.skipBlanks && carriesComments(opt.what)) {
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

  // A copied formula re-anchors to its destination; under Transpose the
  // references into the copied block follow their cells to the transposed tile.
  const copiedFormula = (formula: string, row: number, col: number, sr: number, sc: number) => {
    const from = { row: snap.range.r0 + sr, col: snap.range.c0 + sc };
    return opt.transpose
      ? transposeFormulaRefs(
          formula,
          snap.range,
          from,
          { row, col },
          { row: row - sc, col: col - sr },
        )
      : shiftFormulaRefs(formula, row - from.row, col - from.col);
  };

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
      const copied = snap.cells[sr]?.[sc];
      if (!copied) return;
      // A spilled cell pastes blank so its anchor's formula spills again.
      const src = spilled.has(sr * snap.cols + sc)
        ? { ...copied, value: { kind: 'blank' as const }, formula: null }
        : copied;
      // A cell that only carries a format is still blank to Skip Blanks.
      const isBlankSrc = src.value.kind === 'blank' && !src.formula;
      if (opt.skipBlanks && isBlankSrc) return;

      const addr: Addr = { sheet, row, col };
      // Sheet protection — silently skip locked destinations (spreadsheet parity).
      if (!isCellWritable(state, addr)) return;

      if (opt.what === 'link') {
        wb.setFormula(
          addr,
          linkFormula(snap.range.sheet, sheet, wb, snap.range.r0 + sr, snap.range.c0 + sc),
        );
        return;
      }

      if (carriesComments(opt.what)) {
        const text = commentText(src.format);
        const existing = commentText(state.format.formats.get(addrKey(addr)));
        // Skip Blanks keeps a destination note a noteless source would delete.
        const clears = text === null && existing !== null && !opt.skipBlanks;
        if (text !== null || clears) {
          commentWrites.push({
            addr,
            text,
            author: commentAuthor(src.format),
            visible: text !== null && src.format?.commentVisible === true,
          });
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
      if (valuePatches) {
        const formula =
          shouldPasteFormula && src.formula ? copiedFormula(src.formula, row, col, sr, sc) : null;
        valuePatches.push({ addr, value: src.value, formula });
      } else if (operating) {
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
          wb.setFormula(addr, copiedFormula(src.formula, row, col, sr, sc));
        }
      } else if (wantsValues(opt.what) || wantsFormulas(opt.what)) {
        writeCell(wb, addr, src.value, null);
      }

      // Layer 2: formats
      const fmt = src.format;
      if (opt.what === 'all-except-borders') {
        // Everything but the borders: the destination keeps its own.
        const borders = state.format.formats.get(addrKey(addr))?.borders;
        const next = formatWithoutComment(fmt);
        if (next) delete next.borders;
        formatWrites.push({
          key: addrKey(addr),
          format: borders ? { ...next, borders } : next,
        });
      } else if (opt.what === 'formats') {
        // Formats copies the source's *absence* of formatting too, but not its
        // validation or hyperlink: the destination keeps its own.
        const existing = state.format.formats.get(addrKey(addr));
        const next = withMirroredStyleBorders(formatWithoutComment(fmt) ?? {}, sr, sc);
        for (const field of FORMATS_PASTE_KEEPS) {
          delete next[field];
          if (existing?.[field] !== undefined) Object.assign(next, { [field]: existing[field] });
        }
        formatWrites.push({
          key: addrKey(addr),
          format: Object.keys(next).length > 0 ? next : null,
        });
      } else if (wantsFormats(opt.what)) {
        // A full paste copies the source's *absence* of formatting too: an
        // unformatted source cell clears the destination format rather than
        // leaving a stale one behind (spreadsheet parity).
        formatWrites.push({
          key: addrKey(addr),
          format: formatWithoutComment(fmt),
        });
      } else if (opt.what === 'validation') {
        // The rule replaces the destination's; a source without one clears it.
        formatWrites.push({ key: addrKey(addr), format: { validation: fmt?.validation } });
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

  if (valuePatches) applyPatchesAtomic(wb, valuePatches, history);

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
            // Cherry-picks (number format, validation) merge into the existing
            // cell format; an undefined field removes it.
            const next: CellFormat = { ...formats.get(key), ...format };
            for (const field of Object.keys(format) as (keyof CellFormat)[]) {
              if (format[field] === undefined) delete next[field];
            }
            if (Object.keys(next).length > 0) formats.set(key, next);
            else formats.delete(key);
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

  // Merging is the guess for a kind Excel only offers through the dialog;
  // every other formats-carrying paste replaces the rules on the pasted cells.
  if (opt.what === 'all-merging-conditional-formats') {
    if (!opt.transpose && !bandDestination)
      pasteMergedConditionalRules(store, history, snap, destination);
  } else if (wantsFormats(opt.what)) {
    // A cut moves the rules: the source keeps none of them.
    replaceConditionalRules(
      store,
      history,
      opt.transpose || bandDestination ? [] : translatedConditionalRules(snap, destination),
      sourceCut ? [destination, sourceLogical] : [destination],
    );
  }

  // Comments/notes are part of Paste All, but Excel's Formats paste leaves
  // the destination note untouched. Keep comment history separate from the
  // ordinary format snapshot so the engine comment table and store restore in
  // lockstep on undo/redo.
  if (carriesComments(opt.what)) {
    const sourceCommentsToClear: Addr[] = [];
    const tracked = new Map<string, Addr>();
    if (sourceCut) {
      if (bandAxisFor(sourceLogical)) {
        for (const [key, format] of current.format.formats) {
          const addr = parseAddrKey(key);
          if (
            !addr ||
            addr.sheet !== sourceLogical.sheet ||
            addr.row < sourceLogical.r0 ||
            addr.row > sourceLogical.r1 ||
            addr.col < sourceLogical.c0 ||
            addr.col > sourceLogical.c1 ||
            commentText(format) === null
          ) {
            continue;
          }
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
            commentVisible: undefined,
          });
          wb.setCommentEntry(addr.sheet, addr.row, addr.col, '', '');
        }
        for (const write of commentWrites) {
          mutators.setCellFormat(store, write.addr, {
            comment: write.text ?? undefined,
            commentAuthor: write.text ? (write.author ?? undefined) : undefined,
            commentVisible: write.visible || undefined,
          });
          wb.setCommentEntry(
            write.addr.sheet,
            write.addr.row,
            write.addr.col,
            write.author ?? DEFAULT_COMMENT_AUTHOR,
            write.text ?? '',
            write.visible === true,
          );
        }
      });
    }
  }
  if (wantsFormats(opt.what) && bandDestination) {
    const rowHeights = snap.rowHeights;
    const colWidths = snap.colWidths;
    const logicalRows = sourceLogical.r1 - sourceLogical.r0 + 1;
    const logicalCols = sourceLogical.c1 - sourceLogical.c0 + 1;
    const wholeSheet = isWholeRowRange(sourceLogical) && isWholeColumnRange(sourceLogical);
    if ((bandAxis === 'row' || wholeSheet) && rowHeights) {
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
    if ((bandAxis === 'column' || wholeSheet) && colWidths) {
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

  const writtenRange: Range = { ...destination };
  selectWritten(store, writtenRange, bandDestination);
  if (!bandDestination) extendTablesForPaste(store, history, writtenRange);
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
