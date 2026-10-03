/**
 * Shared formula-reference rewriting.
 *
 * A single tokenizer underpins every reference transform the UI performs:
 * relative-offset shifting (fill / paste), row/column insert-delete
 * adjustment, and cell-band shifting. Centralizing the scanner fixes a class
 * of Excel-fidelity bugs that used to differ between the three former copies:
 *  - function names ending in digits (`LOG10(`, `ATAN2(`) are never mistaken
 *    for cell references;
 *  - sheet-qualified references (`Sheet2!A1`, `'My Sheet'!A1`, 3-D
 *    `Sheet1:Sheet3!A1`) keep their sheet name intact and, for structural
 *    edits on the current sheet, are left untouched (they point elsewhere);
 *  - column/row indices are always range-guarded (16383 / 1048575);
 *  - ranges are adjusted as ranges — deleting one endpoint of `A5:A20`
 *    clamps to the band boundary instead of injecting `#REF!` mid-range.
 */

import { colFromLetters, MAX_COL, MAX_ROW } from '../engine/address.js';
import {
  type Atom,
  type CellRefToken,
  renderAtom,
  renderAtomRaw,
  renderToken,
  renderTokenBody,
  renderWholeColAtom,
  renderWholeRowAtom,
  rewriteRefs,
  type WholeColAtom,
  type WholeColRefToken,
  type WholeRowAtom,
  type WholeRowRefToken,
} from './formula-ref-lexer.js';

/** The outcome of transforming one endpoint against a structural edit. */
type EndpointResult = { kind: 'keep'; col: number; row: number } | { kind: 'ref' }; // fully inside a deleted band → #REF!
interface AxisEndpoint {
  abs: boolean;
  index: number;
}
type AxisEndpointResult = { kind: 'keep'; index: number } | { kind: 'ref' };
/**
 * Shift every relative reference in `formula` by (dRow, dCol) — the transform
 * used when a formula is copied/filled/pasted to a new anchor. Refs pinned
 * with `$` keep that axis. Sheet qualifiers are preserved verbatim (a relative
 * `Sheet2!A1` still shifts its A1 part, matching Excel). Out-of-grid results
 * become `#REF!` when the transformed reference leaves the grid.
 */
export function shiftFormulaRefs(formula: string, dRow: number, dCol: number): string {
  if (!formula.startsWith('=') || (dRow === 0 && dCol === 0)) return formula;
  return rewriteRefs(formula, (tok) => {
    if (tok.sheetQual.includes('[') || tok.sheetQual.includes(']')) return renderToken(tok);
    if (tok.kind === 'whole-col') {
      const shiftCol = (at: WholeColAtom): string | null => {
        const col = colFromLetters(at.label);
        return renderWholeColAtom(at.abs, at.abs ? col : col + dCol);
      };
      const aTxt = shiftCol(tok.a);
      const bTxt = shiftCol(tok.b);
      if (aTxt === null || bTxt === null) return null;
      return `${tok.sheetQual}${aTxt}:${bTxt}`;
    }
    if (tok.kind === 'whole-row') {
      const shiftRow = (at: WholeRowAtom): string | null => {
        const row = Number.parseInt(at.rowStr, 10) - 1;
        return renderWholeRowAtom(at.abs, at.abs ? row : row + dRow);
      };
      const aTxt = shiftRow(tok.a);
      const bTxt = shiftRow(tok.b);
      if (aTxt === null || bTxt === null) return null;
      return `${tok.sheetQual}${aTxt}:${bTxt}`;
    }
    const shiftAtom = (at: Atom): string | null => {
      const col = colFromLetters(at.label);
      const row = Number.parseInt(at.rowStr, 10) - 1;
      const nc = at.absCol ? col : col + dCol;
      const nr = at.absRow ? row : row + dRow;
      return renderAtom(at.absCol, nc, at.absRow, nr);
    };
    const aTxt = shiftAtom(tok.a);
    if (aTxt === null) return null;
    if (!tok.b) return `${tok.sheetQual}${aTxt}`;
    const bTxt = shiftAtom(tok.b);
    if (bTxt === null) return null;
    return `${tok.sheetQual}${aTxt}:${bTxt}`;
  });
}

/**
 * Adjust references for a row/column insert or delete on the current sheet.
 * `axis` is the edited axis, `split` the 0-indexed insertion/deletion start,
 * `delta` the signed shift (>0 insert, <0 delete). Only references pointing at
 * the edited sheet (i.e. *unqualified*) are adjusted — sheet-qualified refs
 * point elsewhere and are left untouched, matching Excel. Ranges are clamped:
 * a range with one endpoint inside the deleted band collapses to the band
 * boundary; a range wholly inside the band becomes `#REF!`.
 */
export function adjustFormulaForRowColEdit(
  formula: string,
  axis: 'row' | 'col',
  split: number,
  delta: number,
): string {
  if (delta === 0) return formula;
  return rewriteRefs(formula, (tok) => {
    // Cross-sheet refs point at another sheet — untouched by an edit here.
    if (tok.sheetQual) return renderToken(tok);
    if (tok.kind === 'whole-col') {
      if (axis !== 'col') return renderToken(tok);
      const adjust = (at: AxisEndpoint): AxisEndpointResult => {
        const col = at.index;
        if (col < split) return { kind: 'keep', index: col };
        if (delta < 0 && col < split - delta) return { kind: 'ref' };
        return { kind: 'keep', index: col + delta };
      };
      return clampAxisRange(
        { abs: tok.a.abs, index: colFromLetters(tok.a.label) },
        { abs: tok.b.abs, index: colFromLetters(tok.b.label) },
        adjust,
        split,
        delta < 0,
        renderWholeColAtom,
      );
    }
    if (tok.kind === 'whole-row') {
      if (axis !== 'row') return renderToken(tok);
      const adjust = (at: AxisEndpoint): AxisEndpointResult => {
        const row = at.index;
        if (row < split) return { kind: 'keep', index: row };
        if (delta < 0 && row < split - delta) return { kind: 'ref' };
        return { kind: 'keep', index: row + delta };
      };
      return clampAxisRange(
        { abs: tok.a.abs, index: Number.parseInt(tok.a.rowStr, 10) - 1 },
        { abs: tok.b.abs, index: Number.parseInt(tok.b.rowStr, 10) - 1 },
        adjust,
        split,
        delta < 0,
        renderWholeRowAtom,
      );
    }
    const adjust = (at: Atom): EndpointResult => {
      const col = colFromLetters(at.label);
      const row = Number.parseInt(at.rowStr, 10) - 1;
      if (axis === 'row') {
        if (row < split) return { kind: 'keep', col, row };
        if (delta < 0 && row < split - delta) return { kind: 'ref' };
        return { kind: 'keep', col, row: row + delta };
      }
      if (col < split) return { kind: 'keep', col, row };
      if (delta < 0 && col < split - delta) return { kind: 'ref' };
      return { kind: 'keep', col: col + delta, row };
    };
    if (!tok.b) {
      const r = adjust(tok.a);
      if (r.kind === 'ref') return null;
      return renderAtom(tok.a.absCol, r.col, tok.a.absRow, r.row);
    }
    // Range: clamp endpoints rather than dropping the whole ref for a partial
    // deletion.
    return clampRange(tok, adjust, axis, split, delta < 0);
  });
}

/**
 * Adjust references for a partial-row/column cell-band shift (Insert/Delete
 * Cells, Shift Down/Right). Only references inside the affected band on the
 * shifted axis move. Cross-sheet refs are left untouched unless `context`
 * resolves them to the edited sheet.
 */
export interface CellBandShiftSheetContext {
  /** Sheet whose cells were inserted or deleted. */
  editedSheet: number;
  /** Sheet where the formula lived. Unqualified refs bind here. */
  formulaSheet: number;
  /** Workbook sheet names indexed by sheet number. */
  sheetNames: readonly string[];
}

export function adjustFormulaForCellBandShift(
  formula: string,
  affected: { r0: number; c0: number; r1: number; c1: number },
  axis: 'down' | 'right' | 'up' | 'left',
  delta: number,
  context?: CellBandShiftSheetContext,
): string {
  if (delta === 0) return formula;
  const vertical = axis === 'down' || axis === 'up';
  return rewriteRefs(formula, (tok) => {
    if (context) {
      const targetSheet = tok.sheetQual
        ? resolveQualifiedSheet(tok.sheetQual, context.sheetNames)
        : context.formulaSheet;
      if (targetSheet === null || targetSheet !== context.editedSheet) return renderToken(tok);
    } else if (tok.sheetQual) {
      return renderToken(tok);
    }
    if (tok.kind !== 'cell') return renderToken(tok);
    const shiftAtom = (at: Atom): EndpointResult => {
      const col = colFromLetters(at.label);
      const row = Number.parseInt(at.rowStr, 10) - 1;
      const inAffectedBand = vertical
        ? row >= affected.r0 && row <= affected.r1 && col >= affected.c0 && col <= affected.c1
        : col >= affected.c0 && col <= affected.c1 && row >= affected.r0 && row <= affected.r1;
      if (!inAffectedBand) return { kind: 'keep', col, row };

      const axisIndex = vertical ? row : col;
      const split = vertical ? affected.r0 : affected.c0;
      if (delta < 0 && axisIndex < split - delta) return { kind: 'ref' };
      return {
        kind: 'keep',
        col: vertical ? col : col + delta,
        row: vertical ? row + delta : row,
      };
    };
    const aResult = shiftAtom(tok.a);
    if (!tok.b) {
      if (aResult.kind === 'ref') return tok.sheetQual ? `${tok.sheetQual}#REF!` : null;
      const atom = renderAtom(tok.a.absCol, aResult.col, tok.a.absRow, aResult.row);
      return atom === null ? null : `${tok.sheetQual}${atom}`;
    }
    const range = clampRange(
      tok,
      shiftAtom,
      vertical ? 'row' : 'col',
      vertical ? affected.r0 : affected.c0,
      delta < 0,
    );
    return range === null
      ? tok.sheetQual
        ? `${tok.sheetQual}#REF!`
        : null
      : `${tok.sheetQual}${range}`;
  });
}

/**
 * Context for moving a contiguous row or column band to an insertion point.
 * `insertionAt` is expressed in the pre-move coordinate space: when the
 * target lies after the source band, removing the source first makes the
 * destination `insertionAt - count`.
 */
export interface AxisBandMoveContext {
  axis: 'row' | 'col';
  sheet: number;
  sourceStart: number;
  count: number;
  insertionAt: number;
  formulaSheet: number;
  sheetNames: readonly string[];
}

/**
 * Rewrite references after a stable row/column band permutation. Structural
 * movement ignores `$` pinning: the marker remains in the rendered formula,
 * while the referenced row/column follows the moved band. References are
 * resolved against the formula's owning sheet, so a qualified reference to
 * the edited sheet is rewritten even when its formula lives elsewhere.
 */
export function adjustFormulaForAxisBandMove(
  formula: string,
  context: AxisBandMoveContext,
): string {
  if (!formula.startsWith('=') || !isValidAxisBandMove(context)) return formula;
  return rewriteRefs(formula, (tok) => {
    const targetSheet = tok.sheetQual
      ? resolveQualifiedSheet(tok.sheetQual, context.sheetNames)
      : context.formulaSheet;
    if (targetSheet === null || targetSheet !== context.sheet) return renderToken(tok);

    if (tok.kind === 'whole-col') {
      if (context.axis !== 'col') return renderToken(tok);
      const range = transformAxisMoveRange(
        colFromLetters(tok.a.label),
        colFromLetters(tok.b.label),
        context,
      );
      if (range === null) return null;
      const a = renderWholeColAtom(tok.a.abs, range.a);
      const b = renderWholeColAtom(tok.b.abs, range.b);
      return a === null || b === null ? null : `${tok.sheetQual}${a}:${b}`;
    }
    if (tok.kind === 'whole-row') {
      if (context.axis !== 'row') return renderToken(tok);
      const range = transformAxisMoveRange(
        Number.parseInt(tok.a.rowStr, 10) - 1,
        Number.parseInt(tok.b.rowStr, 10) - 1,
        context,
      );
      if (range === null) return null;
      const a = renderWholeRowAtom(tok.a.abs, range.a);
      const b = renderWholeRowAtom(tok.b.abs, range.b);
      return a === null || b === null ? null : `${tok.sheetQual}${a}:${b}`;
    }

    const a = transformAxisMoveAtom(tok.a, context);
    if (a === null) return null;
    if (!tok.b) return `${tok.sheetQual}${a}`;
    const range = transformAxisMoveCellRange(tok, context);
    return range === null ? null : `${tok.sheetQual}${range.a}:${range.b}`;
  });
}

/**
 * Update formulas outside a cut/paste payload so references that pointed at
 * the moved cells follow them to the destination. Unlike copy/fill shifting,
 * absolute markers do not pin a moved-cell reference: `$A$1` should become
 * `$C$3` when the cell it points at is cut from A1 to C3.
 */
export interface CutPasteSheetContext {
  /** Sheet containing the cells that were cut. */
  sourceSheet: number;
  /** Sheet receiving the cut cells. */
  destinationSheet: number;
  /** Sheet where this formula lived before the move. */
  formulaSheet: number;
  /** Sheet where this rewritten formula will live. */
  outputSheet: number;
  /** Workbook sheet names indexed by sheet number. */
  sheetNames: readonly string[];
}

export function adjustFormulaForCutPasteMove(
  formula: string,
  source: { r0: number; c0: number; r1: number; c1: number },
  dest: { r0: number; c0: number },
  context?: CutPasteSheetContext,
): string {
  if (context) return adjustFormulaForCutPasteWithContext(formula, source, dest, context);
  const dRow = dest.r0 - source.r0;
  const dCol = dest.c0 - source.c0;
  if (dRow === 0 && dCol === 0) return formula;
  const moveAtom = (at: Atom): string | null => {
    const col = colFromLetters(at.label);
    const row = Number.parseInt(at.rowStr, 10) - 1;
    if (row < source.r0 || row > source.r1 || col < source.c0 || col > source.c1) {
      return renderAtomRaw(at);
    }
    return renderAtom(at.absCol, col + dCol, at.absRow, row + dRow);
  };
  return rewriteRefs(formula, (tok) => {
    if (tok.sheetQual) return renderToken(tok);
    if (tok.kind !== 'cell') {
      const cutAxis = fullAxisForCut(source);
      const transformed = transformWholeAxisCutToken(tok, cutAxis, source, dRow, dCol);
      return transformed ?? renderToken(tok);
    }
    const aTxt = moveAtom(tok.a);
    if (!tok.b) return aTxt;

    // Excel treats a range as one rectangular reference during a cut. A
    // partially overlapping range on the same sheet is left intact; moving
    // each endpoint independently would create a range that never existed.
    const disposition = classifyCutRange(tok.a, tok.b, source, false);
    if (disposition.kind !== 'move') return renderToken(tok);
    if (aTxt === null) return null;
    const bTxt = moveAtom(tok.b);
    if (bTxt === null) return null;
    return `${aTxt}:${bTxt}`;
  });
}

function adjustFormulaForCutPasteWithContext(
  formula: string,
  source: { r0: number; c0: number; r1: number; c1: number },
  dest: { r0: number; c0: number },
  context: CutPasteSheetContext,
): string {
  const dRow = dest.r0 - source.r0;
  const dCol = dest.c0 - source.c0;
  return rewriteRefs(formula, (tok) => {
    if (tok.kind !== 'cell') {
      const targetSheet = resolveReferenceSheet(tok.sheetQual, context);
      if (targetSheet === null) return renderToken(tok);
      const originalQualifier = tok.sheetQual || qualifierForOutputSheet(targetSheet, context);
      const cutAxis = fullAxisForCut(source);
      if (targetSheet !== context.sourceSheet || cutAxis === null) {
        return `${originalQualifier}${renderTokenBody(tok)}`;
      }
      const transformed = transformWholeAxisCutToken(tok, cutAxis, source, dRow, dCol);
      if (transformed === null) return `${originalQualifier}${renderTokenBody(tok)}`;
      const wasMoved = wholeAxisCutContainsToken(tok, cutAxis, source);
      const qualifier = wasMoved
        ? qualifierForOutputSheet(context.destinationSheet, context)
        : originalQualifier;
      return `${qualifier}${transformed}`;
    }
    const targetSheet = resolveReferenceSheet(tok.sheetQual, context);
    if (targetSheet === null) return renderToken(tok);

    const a = rewriteCutPasteEndpoint(
      tok.a,
      tok.sheetQual,
      targetSheet,
      source,
      dRow,
      dCol,
      context,
    );
    if (!tok.b) {
      if (a.atom === null) return null;
      return `${a.qualifier}${a.atom}`;
    }

    // A range is transformed as a rectangle. When both corners are inside
    // the cut rectangle, both corners move together. For partial overlap,
    // same-sheet cuts keep the range verbatim; cross-sheet cuts trim only a
    // complete edge strip. Corner and middle overlaps remain unchanged.
    const originalQualifier = tok.sheetQual || qualifierForOutputSheet(targetSheet, context);
    if (targetSheet !== context.sourceSheet) {
      return `${originalQualifier}${renderAtomRaw(tok.a)}:${renderAtomRaw(tok.b)}`;
    }
    const disposition = classifyCutRange(
      tok.a,
      tok.b,
      source,
      context.sourceSheet !== context.destinationSheet && targetSheet === context.sourceSheet,
    );
    if (disposition.kind === 'move') {
      if (a.atom === null) return null;
      const b = rewriteCutPasteEndpoint(
        tok.b,
        tok.sheetQual,
        targetSheet,
        source,
        dRow,
        dCol,
        context,
      );
      if (b.atom === null) return null;
      const qualifier = qualifierForOutputSheet(context.destinationSheet, context);
      return `${qualifier}${a.atom}:${b.atom}`;
    }
    if (disposition.kind === 'keep') {
      return `${originalQualifier}${renderAtomRaw(tok.a)}:${renderAtomRaw(tok.b)}`;
    }

    const trimmed = trimCutRange(tok.a, tok.b, disposition);
    if (trimmed === null) return null;
    return `${originalQualifier}${trimmed.a}:${trimmed.b}`;
  });
}

interface CellRangeBounds {
  r0: number;
  c0: number;
  r1: number;
  c1: number;
}

type CutRangeDisposition =
  | { kind: 'move' }
  | { kind: 'keep' }
  | {
      kind: 'trim';
      side: 'top' | 'bottom' | 'left' | 'right';
      intersection: CellRangeBounds;
    };

function atomBounds(a: Atom, b: Atom): CellRangeBounds {
  const aRow = Number.parseInt(a.rowStr, 10) - 1;
  const aCol = colFromLetters(a.label);
  const bRow = Number.parseInt(b.rowStr, 10) - 1;
  const bCol = colFromLetters(b.label);
  return {
    r0: Math.min(aRow, bRow),
    c0: Math.min(aCol, bCol),
    r1: Math.max(aRow, bRow),
    c1: Math.max(aCol, bCol),
  };
}

function rectangleBounds(rect: {
  r0: number;
  c0: number;
  r1: number;
  c1: number;
}): CellRangeBounds {
  return {
    r0: Math.min(rect.r0, rect.r1),
    c0: Math.min(rect.c0, rect.c1),
    r1: Math.max(rect.r0, rect.r1),
    c1: Math.max(rect.c0, rect.c1),
  };
}

/** Classify a cell range against a cut rectangle before rewriting either end. */
function classifyCutRange(
  a: Atom,
  b: Atom,
  source: { r0: number; c0: number; r1: number; c1: number },
  allowEdgeTrim: boolean,
): CutRangeDisposition {
  const range = atomBounds(a, b);
  const cut = rectangleBounds(source);
  const aInside =
    Number.parseInt(a.rowStr, 10) - 1 >= cut.r0 &&
    Number.parseInt(a.rowStr, 10) - 1 <= cut.r1 &&
    colFromLetters(a.label) >= cut.c0 &&
    colFromLetters(a.label) <= cut.c1;
  const bInside =
    Number.parseInt(b.rowStr, 10) - 1 >= cut.r0 &&
    Number.parseInt(b.rowStr, 10) - 1 <= cut.r1 &&
    colFromLetters(b.label) >= cut.c0 &&
    colFromLetters(b.label) <= cut.c1;
  if (aInside && bInside) return { kind: 'move' };
  if (!allowEdgeTrim) return { kind: 'keep' };

  const intersection: CellRangeBounds = {
    r0: Math.max(range.r0, cut.r0),
    c0: Math.max(range.c0, cut.c0),
    r1: Math.min(range.r1, cut.r1),
    c1: Math.min(range.c1, cut.c1),
  };
  if (intersection.r0 > intersection.r1 || intersection.c0 > intersection.c1) {
    return { kind: 'keep' };
  }

  // Only a full-width top/bottom strip or full-height left/right strip leaves
  // one contiguous rectangular remainder. A corner or middle intersection
  // must retain the original range.
  if (
    intersection.c0 === range.c0 &&
    intersection.c1 === range.c1 &&
    intersection.r0 === range.r0 &&
    intersection.r1 < range.r1
  ) {
    return { kind: 'trim', side: 'top', intersection };
  }
  if (
    intersection.c0 === range.c0 &&
    intersection.c1 === range.c1 &&
    intersection.r1 === range.r1 &&
    intersection.r0 > range.r0
  ) {
    return { kind: 'trim', side: 'bottom', intersection };
  }
  if (
    intersection.r0 === range.r0 &&
    intersection.r1 === range.r1 &&
    intersection.c0 === range.c0 &&
    intersection.c1 < range.c1
  ) {
    return { kind: 'trim', side: 'left', intersection };
  }
  if (
    intersection.r0 === range.r0 &&
    intersection.r1 === range.r1 &&
    intersection.c1 === range.c1 &&
    intersection.c0 > range.c0
  ) {
    return { kind: 'trim', side: 'right', intersection };
  }
  return { kind: 'keep' };
}

/** Render a range after removing one complete edge strip, retaining endpoint order. */
function trimCutRange(
  a: Atom,
  b: Atom,
  disposition: Extract<CutRangeDisposition, { kind: 'trim' }>,
): { a: string; b: string } | null {
  let aCol = colFromLetters(a.label);
  let aRow = Number.parseInt(a.rowStr, 10) - 1;
  let bCol = colFromLetters(b.label);
  let bRow = Number.parseInt(b.rowStr, 10) - 1;
  const { intersection } = disposition;
  if (disposition.side === 'top') {
    const row = intersection.r1 + 1;
    if (aRow <= bRow) aRow = row;
    else bRow = row;
  } else if (disposition.side === 'bottom') {
    const row = intersection.r0 - 1;
    if (aRow >= bRow) aRow = row;
    else bRow = row;
  } else if (disposition.side === 'left') {
    const col = intersection.c1 + 1;
    if (aCol <= bCol) aCol = col;
    else bCol = col;
  } else {
    const col = intersection.c0 - 1;
    if (aCol >= bCol) aCol = col;
    else bCol = col;
  }
  const aTxt = renderAtom(a.absCol, aCol, a.absRow, aRow);
  const bTxt = renderAtom(b.absCol, bCol, b.absRow, bRow);
  if (aTxt === null || bTxt === null) return null;
  return { a: aTxt, b: bTxt };
}

function rewriteCutPasteEndpoint(
  at: Atom,
  originalQualifier: string,
  targetSheet: number,
  source: { r0: number; c0: number; r1: number; c1: number },
  dRow: number,
  dCol: number,
  context: CutPasteSheetContext,
): { atom: string | null; qualifier: string; moved: boolean } {
  const col = colFromLetters(at.label);
  const row = Number.parseInt(at.rowStr, 10) - 1;
  const moved =
    targetSheet === context.sourceSheet &&
    row >= source.r0 &&
    row <= source.r1 &&
    col >= source.c0 &&
    col <= source.c1;
  if (moved) {
    const atom = renderAtom(at.absCol, col + dCol, at.absRow, row + dRow);
    if (atom === null) return { atom: null, qualifier: '', moved: true };
    return {
      atom,
      qualifier: qualifierForOutputSheet(context.destinationSheet, context),
      moved: true,
    };
  }

  const qualifier = originalQualifier
    ? originalQualifier
    : qualifierForOutputSheet(targetSheet, context);
  return { atom: renderAtomRaw(at), qualifier, moved: false };
}

function axisBandLimit(axis: 'row' | 'col'): number {
  return axis === 'row' ? MAX_ROW + 1 : MAX_COL + 1;
}

function isValidAxisBandMove(context: AxisBandMoveContext): boolean {
  if (
    !Number.isInteger(context.sourceStart) ||
    !Number.isInteger(context.count) ||
    !Number.isInteger(context.insertionAt)
  ) {
    return false;
  }
  if (context.count <= 0 || context.sourceStart < 0 || context.insertionAt < 0) return false;
  const limit = axisBandLimit(context.axis);
  const sourceEnd = context.sourceStart + context.count;
  if (sourceEnd > limit || context.insertionAt > limit) return false;
  // The core rejects an insertion point inside the source band. Moving to
  // either edge is a no-op and is deliberately handled as such.
  if (context.insertionAt > context.sourceStart && context.insertionAt < sourceEnd) return false;
  return context.insertionAt !== context.sourceStart && context.insertionAt !== sourceEnd;
}

function mapAxisIndexForMove(index: number, context: AxisBandMoveContext): number {
  const sourceEnd = context.sourceStart + context.count;
  const finalStart =
    context.insertionAt < context.sourceStart
      ? context.insertionAt
      : context.insertionAt - context.count;
  if (index >= context.sourceStart && index < sourceEnd) {
    return finalStart + (index - context.sourceStart);
  }
  if (context.insertionAt < context.sourceStart) {
    if (index >= context.insertionAt && index < context.sourceStart) return index + context.count;
    return index;
  }
  if (index >= sourceEnd && index < context.insertionAt) return index - context.count;
  return index;
}

interface AxisMoveRange {
  a: number;
  b: number;
}

/**
 * Apply the range boundary rules used by a structural band move. The move is
 * modelled as insertion at the destination followed by removal of the source;
 * source endpoints that move with the band are restored after that interval
 * calculation so a scalar source reference and a range beginning at it keep
 * their moved destination.
 */
function transformAxisMoveRange(
  first: number,
  last: number,
  context: AxisBandMoveContext,
): AxisMoveRange | null {
  const sourceEnd = context.sourceStart + context.count;

  const mappedFirst = mapAxisIndexForMove(first, context);
  const mappedLast = mapAxisIndexForMove(last, context);
  const mappedOrderIsValid = first <= last ? mappedFirst <= mappedLast : mappedFirst >= mappedLast;
  if (mappedOrderIsValid) {
    return {
      a: mappedFirst,
      b: mappedLast,
    };
  }

  const insertIndex = context.insertionAt;
  const insertedFirst = first >= insertIndex ? first + context.count : first;
  const insertedLast = last >= insertIndex ? last + context.count : last;
  const deleteStart = insertIndex < context.sourceStart ? sourceEnd : context.sourceStart;
  const deleteEnd = deleteStart + context.count;
  const firstDeleted = insertedFirst >= deleteStart && insertedFirst < deleteEnd;
  const lastDeleted = insertedLast >= deleteStart && insertedLast < deleteEnd;
  const survivor = (value: number): number => (value >= deleteEnd ? value - context.count : value);
  const boundaryFor = (other: number): number =>
    other < deleteStart ? deleteStart - 1 : deleteStart;

  const outFirst = firstDeleted ? boundaryFor(insertedLast) : survivor(insertedFirst);
  const outLast = lastDeleted ? boundaryFor(insertedFirst) : survivor(insertedLast);

  // A range can straddle the moved band while neither endpoint is inside it.
  // Keep the endpoint orientation from the source formula; callers render the
  // returned pair directly rather than normalising the range.
  return { a: outFirst, b: outLast };
}

function transformAxisMoveAtom(at: Atom, context: AxisBandMoveContext): string | null {
  const col = colFromLetters(at.label);
  const row = Number.parseInt(at.rowStr, 10) - 1;
  if (context.axis === 'row') {
    return renderAtom(at.absCol, col, at.absRow, mapAxisIndexForMove(row, context));
  }
  return renderAtom(at.absCol, mapAxisIndexForMove(col, context), at.absRow, row);
}

function transformAxisMoveCellRange(
  tok: CellRefToken,
  context: AxisBandMoveContext,
): { a: string; b: string } | null {
  const b = tok.b as Atom;
  const first =
    context.axis === 'row' ? Number.parseInt(tok.a.rowStr, 10) - 1 : colFromLetters(tok.a.label);
  const last = context.axis === 'row' ? Number.parseInt(b.rowStr, 10) - 1 : colFromLetters(b.label);
  const range = transformAxisMoveRange(first, last, context);
  if (range === null) return null;
  const aCol = colFromLetters(tok.a.label);
  const aRow = Number.parseInt(tok.a.rowStr, 10) - 1;
  const bCol = colFromLetters(b.label);
  const bRow = Number.parseInt(b.rowStr, 10) - 1;
  const a =
    context.axis === 'row'
      ? renderAtom(tok.a.absCol, aCol, tok.a.absRow, range.a)
      : renderAtom(tok.a.absCol, range.a, tok.a.absRow, aRow);
  const renderedB =
    context.axis === 'row'
      ? renderAtom(b.absCol, bCol, b.absRow, range.b)
      : renderAtom(b.absCol, range.b, b.absRow, bRow);
  return a === null || renderedB === null ? null : { a, b: renderedB };
}

type FullAxis = 'row' | 'col' | null;

function fullAxisForCut(source: { r0: number; c0: number; r1: number; c1: number }): FullAxis {
  const r0 = Math.min(source.r0, source.r1);
  const r1 = Math.max(source.r0, source.r1);
  const c0 = Math.min(source.c0, source.c1);
  const c1 = Math.max(source.c0, source.c1);
  if (c0 === 0 && c1 === MAX_COL) return 'row';
  if (r0 === 0 && r1 === MAX_ROW) return 'col';
  return null;
}

function wholeAxisCutContainsToken(
  tok: WholeColRefToken | WholeRowRefToken,
  axis: FullAxis,
  source: { r0: number; c0: number; r1: number; c1: number },
): boolean {
  const first =
    tok.kind === 'whole-row' ? Number.parseInt(tok.a.rowStr, 10) - 1 : colFromLetters(tok.a.label);
  const last =
    tok.kind === 'whole-row' ? Number.parseInt(tok.b.rowStr, 10) - 1 : colFromLetters(tok.b.label);
  const low = Math.min(first, last);
  const high = Math.max(first, last);
  if (axis === 'row' && tok.kind === 'whole-row') {
    return low >= Math.min(source.r0, source.r1) && high <= Math.max(source.r0, source.r1);
  }
  if (axis === 'col' && tok.kind === 'whole-col') {
    return low >= Math.min(source.c0, source.c1) && high <= Math.max(source.c0, source.c1);
  }
  return false;
}

function transformWholeAxisCutToken(
  tok: WholeColRefToken | WholeRowRefToken,
  axis: FullAxis,
  source: { r0: number; c0: number; r1: number; c1: number },
  dRow: number,
  dCol: number,
): string | null {
  if (axis === null || !wholeAxisCutContainsToken(tok, axis, source)) return renderTokenBody(tok);
  if (axis === 'row' && tok.kind === 'whole-row') {
    const a = renderWholeRowAtom(tok.a.abs, Number.parseInt(tok.a.rowStr, 10) - 1 + dRow);
    const b = renderWholeRowAtom(tok.b.abs, Number.parseInt(tok.b.rowStr, 10) - 1 + dRow);
    return a === null || b === null ? null : `${a}:${b}`;
  }
  if (axis === 'col' && tok.kind === 'whole-col') {
    const a = renderWholeColAtom(tok.a.abs, colFromLetters(tok.a.label) + dCol);
    const b = renderWholeColAtom(tok.b.abs, colFromLetters(tok.b.label) + dCol);
    return a === null || b === null ? null : `${a}:${b}`;
  }
  return renderTokenBody(tok);
}

function resolveReferenceSheet(sheetQual: string, context: CutPasteSheetContext): number | null {
  if (!sheetQual) return context.formulaSheet;
  return resolveQualifiedSheet(sheetQual, context.sheetNames);
}

function resolveQualifiedSheet(sheetQual: string, sheetNames: readonly string[]): number | null {
  const body = sheetQual.slice(0, -1);
  // 3-D references and external workbook qualifiers are deliberately kept
  // verbatim until their own workbook-aware transform exists.
  if (body.includes(':') || body.includes('[') || body.includes(']')) return null;
  const name =
    body.startsWith("'") && body.endsWith("'") ? body.slice(1, -1).replace(/''/g, "'") : body;
  const folded = name.toLocaleLowerCase();
  const index = sheetNames.findIndex((sheetName) => sheetName.toLocaleLowerCase() === folded);
  return index >= 0 ? index : null;
}

function qualifierForOutputSheet(sheet: number, context: CutPasteSheetContext): string {
  if (sheet === context.outputSheet) return '';
  const name = context.sheetNames[sheet];
  if (name === undefined) return '';
  const safe = /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name);
  return `${safe ? name : `'${name.replace(/'/g, "''")}'`}!`;
}

/** Clamp a range against an edit so a partial deletion keeps the surviving
 * span instead of turning the whole reference into `#REF!`. */
function clampRange(
  tok: CellRefToken,
  adjust: (at: Atom) => EndpointResult,
  axis: 'row' | 'col',
  split: number,
  deletion: boolean,
): string | null {
  const b = tok.b as Atom;
  const ra = adjust(tok.a);
  const rb = adjust(b);
  if (ra.kind === 'ref' && rb.kind === 'ref') return null; // whole range deleted
  // A non-deletion edit can only produce `ref` when the shifted endpoint
  // leaves the worksheet; there is no surviving boundary to clamp to.
  if (!deletion && (ra.kind === 'ref' || rb.kind === 'ref')) return null;
  // One endpoint deleted → clamp it to the boundary that survives.
  const resolve = (r: EndpointResult, at: Atom, other: Atom): { col: number; row: number } => {
    if (r.kind === 'keep') return { col: r.col, row: r.row };
    // An endpoint at the upper edge of a range clamps to the first surviving
    // line after the deleted band. An endpoint at the lower edge clamps to
    // the last surviving line before it.
    const col = colFromLetters(at.label);
    const row = Number.parseInt(at.rowStr, 10) - 1;
    const otherAxis =
      axis === 'row' ? Number.parseInt(other.rowStr, 10) - 1 : colFromLetters(other.label);
    const boundary = otherAxis < split ? split - 1 : split;
    if (axis === 'row') return { col, row: boundary };
    return { col: boundary, row };
  };
  const pa = resolve(ra, tok.a, b);
  const pb = resolve(rb, b, tok.a);
  const aTxt = renderAtom(tok.a.absCol, pa.col, tok.a.absRow, pa.row);
  const bTxt = renderAtom(b.absCol, pb.col, b.absRow, pb.row);
  if (aTxt === null || bTxt === null) return null;
  return `${aTxt}:${bTxt}`;
}

function clampAxisRange(
  a: AxisEndpoint,
  b: AxisEndpoint,
  adjust: (at: AxisEndpoint) => AxisEndpointResult,
  split: number,
  deletion: boolean,
  render: (abs: boolean, index: number) => string | null,
): string | null {
  const ra = adjust(a);
  const rb = adjust(b);
  if (ra.kind === 'ref' && rb.kind === 'ref') return null;
  if (!deletion && (ra.kind === 'ref' || rb.kind === 'ref')) return null;

  const resolve = (result: AxisEndpointResult, other: AxisEndpoint): number => {
    if (result.kind === 'keep') return result.index;
    return other.index < split ? split - 1 : split;
  };
  const aTxt = render(a.abs, resolve(ra, b));
  const bTxt = render(b.abs, resolve(rb, a));
  if (aTxt === null || bTxt === null) return null;
  return `${aTxt}:${bTxt}`;
}
