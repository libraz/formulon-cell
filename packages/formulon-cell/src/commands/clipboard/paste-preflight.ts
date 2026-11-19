import { MAX_COL, MAX_ROW } from '../../engine/address.js';
import type { Range } from '../../engine/types.js';
import { rangeContainsRange, rangesIntersect, sameRange } from '../../store/selection-geometry.js';
import type { State } from '../../store/store.js';
import { isCellWritable, isRangeInAllowedEditRange } from '../protection.js';
import { bandAxisFor, logicalRangeFor } from './paste-destination.js';
import type { PasteKind } from './paste-kind.js';
import { wantsFormats } from './paste-kind.js';
import type { ClipboardSnapshot } from './snapshot.js';

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
export const preflight = (
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

  // A merge the destination cuts off (a cropped or wrapped tile) is not pasted.
  const translatedMerges = wantsFormats(what)
    ? (translateMerges(snap, destination, transpose)?.filter((merge) =>
        rangeContainsRange(destination, merge),
      ) ?? null)
    : [];
  if (translatedMerges === null) return null;
  // A scalar paste into the exact logical footprint of a merged cell updates
  // its anchor and keeps the merge; a Formats paste unmerges it. Non-scalar tiles intentionally continue
  // through the normal topology replacement path (Excel unmerges those
  // covered cells unless the source carries a matching merge).
  const preservedMerges =
    snap.rows === 1 && snap.cols === 1 && what !== 'formats'
      ? [...state.merges.byAnchor.values()].filter((merge) => sameRange(merge, destination))
      : [];
  return { destination, translatedMerges, preservedMerges };
};
