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

export const MAX_COL_INDEX = 16383;
export const MAX_ROW_INDEX = 1048575;

/** Convert an uppercase A1 column label to a 0-indexed column. */
export function colLabelToIndex(label: string): number {
  let n = 0;
  for (let i = 0; i < label.length; i += 1) n = n * 26 + (label.charCodeAt(i) - 64);
  return n - 1;
}

/** Convert a 0-indexed column to its A1 label. */
export function colIndexToLabel(col: number): string {
  let n = col;
  let out = '';
  do {
    out = String.fromCharCode(65 + (n % 26)) + out;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return out;
}

/** A single A1 endpoint (`$A$1` → absCol,label,absRow,row). */
interface Atom {
  absCol: boolean;
  label: string;
  absRow: boolean;
  rowStr: string;
}

/** A cell reference token surfaced in a formula. */
interface CellRefToken {
  kind: 'cell';
  /** Raw sheet-qualifier text including the trailing `!` (e.g. `Sheet2!`,
   *  `'My Sheet'!`, `Sheet1:Sheet3!`), or '' when the ref is unqualified. */
  sheetQual: string;
  a: Atom;
  b: Atom | null;
  /** Character offset just past the whole token. */
  end: number;
}

interface WholeColAtom {
  abs: boolean;
  label: string;
}

interface WholeRowAtom {
  abs: boolean;
  rowStr: string;
}

interface WholeColRefToken {
  kind: 'whole-col';
  sheetQual: string;
  a: WholeColAtom;
  b: WholeColAtom;
  end: number;
}

interface WholeRowRefToken {
  kind: 'whole-row';
  sheetQual: string;
  a: WholeRowAtom;
  b: WholeRowAtom;
  end: number;
}

type RefToken = CellRefToken | WholeColRefToken | WholeRowRefToken;

const isAsciiLetter = (c: string): boolean => (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z');
const isAsciiDigit = (c: string): boolean => c >= '0' && c <= '9';
const isUnicodeLetter = (c: string): boolean => /^\p{L}$/u.test(c);
const isUnicodeDigit = (c: string): boolean => /^\p{N}$/u.test(c);
const isUnicodeMark = (c: string): boolean => /^\p{M}$/u.test(c);
const isLetter = (c: string): boolean => isAsciiLetter(c);
const isDigit = (c: string): boolean => isAsciiDigit(c);
const isIdentChar = (c: string): boolean =>
  isAsciiLetter(c) ||
  isAsciiDigit(c) ||
  isUnicodeLetter(c) ||
  isUnicodeDigit(c) ||
  isUnicodeMark(c) ||
  c === '_';

/** Parse a single A1 atom (`$?[A-Za-z]+$?[0-9]+`) at `start`. Returns the atom
 *  and its end offset, or null when the text there is not an atom. */
function parseAtom(src: string, start: number): { atom: Atom; end: number } | null {
  let i = start;
  let absCol = false;
  if (src[i] === '$') {
    absCol = true;
    i += 1;
  }
  const lettersStart = i;
  while (i < src.length && isLetter(src[i] ?? '')) i += 1;
  if (i === lettersStart) return null;
  const label = src.slice(lettersStart, i).toUpperCase();
  let absRow = false;
  if (src[i] === '$') {
    absRow = true;
    i += 1;
  }
  const digitsStart = i;
  while (i < src.length && isDigit(src[i] ?? '')) i += 1;
  if (i === digitsStart) return null;
  return { atom: { absCol, label, absRow, rowStr: src.slice(digitsStart, i) }, end: i };
}

/** Parse a bare sheet-name word (`Sheet1`, `Data`, `_x`) — a name may contain
 *  digits and dots after the first char, but must start with a Unicode letter
 *  or `_`. */
function parseSheetWord(src: string, start: number): number | null {
  let i = start;
  const first = src[i] ?? '';
  if (!(isUnicodeLetter(first) || first === '_')) return null;
  i += 1;
  while (i < src.length) {
    const c = src[i] ?? '';
    if (isUnicodeLetter(c) || isUnicodeDigit(c) || isUnicodeMark(c) || c === '_' || c === '.')
      i += 1;
    else break;
  }
  return i;
}

/** Parse a whole-column endpoint (`$A` or `A`). */
function parseWholeColAtom(src: string, start: number): { atom: WholeColAtom; end: number } | null {
  let i = start;
  let abs = false;
  if (src[i] === '$') {
    abs = true;
    i += 1;
  }
  const lettersStart = i;
  while (i < src.length && isLetter(src[i] ?? '')) i += 1;
  if (i === lettersStart) return null;
  return { atom: { abs, label: src.slice(lettersStart, i).toUpperCase() }, end: i };
}

/** Parse a whole-row endpoint (`$1` or `1`). */
function parseWholeRowAtom(src: string, start: number): { atom: WholeRowAtom; end: number } | null {
  let i = start;
  let abs = false;
  if (src[i] === '$') {
    abs = true;
    i += 1;
  }
  const digitsStart = i;
  while (i < src.length && isDigit(src[i] ?? '')) i += 1;
  if (i === digitsStart) return null;
  return { atom: { abs, rowStr: src.slice(digitsStart, i) }, end: i };
}

/** Parse an optional sheet qualifier (`Name!`, `'Name'!`, `A:B!`) at `start`.
 *  Returns the raw qualifier text (with trailing `!`) and its end, or null. */
function parseSheetQualifier(src: string, start: number): { text: string; end: number } | null {
  const readOne = (from: number): number | null => {
    if (src[from] === "'") {
      let i = from + 1;
      while (i < src.length) {
        if (src[i] === "'") {
          if (src[i + 1] === "'") {
            i += 2;
            continue;
          }
          i += 1;
          return i;
        }
        i += 1;
      }
      return null; // unterminated quote
    }
    return parseSheetWord(src, from);
  };
  const firstEnd = readOne(start);
  if (firstEnd === null) return null;
  let end = firstEnd;
  // 3-D qualifier: Sheet1:Sheet3!
  if (src[end] === ':') {
    const secondEnd = readOne(end + 1);
    if (secondEnd !== null) end = secondEnd;
  }
  if (src[end] !== '!') return null;
  return { text: src.slice(start, end + 1), end: end + 1 };
}

/** Attempt to match a whole reference token at `start`. Rejects tokens that are
 *  actually function names (immediately followed by `(`) or that run into an
 *  identifier continuation. */
function matchRefToken(src: string, start: number): RefToken | null {
  let i = start;
  let sheetQual = '';
  const qual = parseSheetQualifier(src, i);
  if (qual) {
    sheetQual = qual.text;
    i = qual.end;
  }

  const wholeColA = parseWholeColAtom(src, i);
  if (wholeColA && src[wholeColA.end] === ':') {
    const wholeColB = parseWholeColAtom(src, wholeColA.end + 1);
    if (wholeColB) {
      const end = wholeColB.end;
      const next = src[end] ?? '';
      if (
        next !== '(' &&
        next !== '[' &&
        !isIdentChar(next) &&
        wholeColInGrid(wholeColA.atom) &&
        wholeColInGrid(wholeColB.atom)
      ) {
        return { kind: 'whole-col', sheetQual, a: wholeColA.atom, b: wholeColB.atom, end };
      }
    }
  }

  const wholeRowA = parseWholeRowAtom(src, i);
  if (wholeRowA && src[wholeRowA.end] === ':') {
    const wholeRowB = parseWholeRowAtom(src, wholeRowA.end + 1);
    if (wholeRowB) {
      const end = wholeRowB.end;
      const next = src[end] ?? '';
      if (
        next !== '(' &&
        next !== '[' &&
        !isIdentChar(next) &&
        wholeRowInGrid(wholeRowA.atom) &&
        wholeRowInGrid(wholeRowB.atom)
      ) {
        return { kind: 'whole-row', sheetQual, a: wholeRowA.atom, b: wholeRowB.atom, end };
      }
    }
  }

  const first = parseAtom(src, i);
  if (!first) return null;
  i = first.end;
  let b: Atom | null = null;
  if (src[i] === ':') {
    const second = parseAtom(src, i + 1);
    if (second) {
      b = second.atom;
      i = second.end;
    }
  }
  // Reject function calls (`SUM(`, `LOG10(`) and identifier run-ons.
  const next = src[i] ?? '';
  if (next === '(') return null;
  // A table name in a structured reference can look exactly like an A1
  // token (for example `T1[Amount]`). The bracketed part can also contain
  // cell-looking column names, so leave the complete structured reference to
  // the bracket-aware scanner below.
  if (next === '[') return null;
  if (isIdentChar(next)) return null;
  // Reject tokens outside the grid (`Year2024`, `SHEET`) — those are defined
  // names / words, not cell references, and must pass through untouched.
  if (!atomInGrid(first.atom)) return null;
  if (b && !atomInGrid(b)) return null;
  return { kind: 'cell', sheetQual, a: first.atom, b, end: i };
}

/** True when an atom addresses a cell inside the grid (a real ref, not a name
 *  like `Year2024` whose "column" exceeds the last column). */
function atomInGrid(at: Atom): boolean {
  const col = colLabelToIndex(at.label);
  const row = Number.parseInt(at.rowStr, 10) - 1;
  return col >= 0 && col <= MAX_COL_INDEX && row >= 0 && row <= MAX_ROW_INDEX;
}

function wholeColInGrid(at: WholeColAtom): boolean {
  const col = colLabelToIndex(at.label);
  return col >= 0 && col <= MAX_COL_INDEX;
}

function wholeRowInGrid(at: WholeRowAtom): boolean {
  const row = Number.parseInt(at.rowStr, 10) - 1;
  return row >= 0 && row <= MAX_ROW_INDEX;
}

/** Consume a `"..."` string literal (with `""` escape) starting at `start`. */
function consumeString(src: string, start: number): { text: string; end: number } {
  let i = start + 1;
  while (i < src.length) {
    if (src[i] === '"') {
      if (src[i + 1] === '"') {
        i += 2;
        continue;
      }
      i += 1;
      break;
    }
    i += 1;
  }
  return { text: src.slice(start, i), end: i };
}

/** Consume one structured-reference bracket expression, including nested
 * brackets (`Table1[[#Headers],[Amount]]`). In Excel table syntax an
 * apostrophe escapes bracket, `@`, `#`, and apostrophe characters in a header;
 * those escaped brackets must not change the nesting depth. */
function consumeBracketed(src: string, start: number): { text: string; end: number } {
  let depth = 0;
  let i = start;
  while (i < src.length) {
    const ch = src[i] ?? '';
    if (ch === "'") {
      const next = src[i + 1] ?? '';
      if (next === '[' || next === ']' || next === '@' || next === '#' || next === "'") {
        i += 2;
        continue;
      }
    }
    if (ch === '[') {
      depth += 1;
    } else if (ch === ']') {
      depth -= 1;
      if (depth === 0) {
        i += 1;
        break;
      }
    }
    i += 1;
  }
  return { text: src.slice(start, i), end: i };
}

/** Render an atom to A1 text, or null when it falls outside the grid. */
function renderAtom(absCol: boolean, col: number, absRow: boolean, row: number): string | null {
  if (col < 0 || row < 0 || col > MAX_COL_INDEX || row > MAX_ROW_INDEX) return null;
  return `${absCol ? '$' : ''}${colIndexToLabel(col)}${absRow ? '$' : ''}${row + 1}`;
}

function renderWholeColAtom(abs: boolean, col: number): string | null {
  if (col < 0 || col > MAX_COL_INDEX) return null;
  return `${abs ? '$' : ''}${colIndexToLabel(col)}`;
}

function renderWholeRowAtom(abs: boolean, row: number): string | null {
  if (row < 0 || row > MAX_ROW_INDEX) return null;
  return `${abs ? '$' : ''}${row + 1}`;
}

/** The outcome of transforming one endpoint against a structural edit. */
type EndpointResult = { kind: 'keep'; col: number; row: number } | { kind: 'ref' }; // fully inside a deleted band → #REF!
interface AxisEndpoint {
  abs: boolean;
  index: number;
}
type AxisEndpointResult = { kind: 'keep'; index: number } | { kind: 'ref' };

/** Walk `formula`, replacing each reference token via `visit`. The visitor
 *  receives the token and returns the replacement text, or null to emit
 *  `#REF!` for the whole token. String literals, `#REF!` tokens, function
 *  names, and non-reference text are passed through untouched. */
function rewriteRefs(formula: string, visit: (tok: RefToken) => string | null): string {
  if (!formula.startsWith('=')) return formula;
  let out = '';
  let i = 0;
  while (i < formula.length) {
    const ch = formula[i] ?? '';
    if (ch === '"') {
      const lit = consumeString(formula, i);
      out += lit.text;
      i = lit.end;
      continue;
    }
    if (ch === '[') {
      const structured = consumeBracketed(formula, i);
      // An external workbook qualifier (`[Book.xlsx]Sheet1!A1`) is not a
      // structured reference. Keep the whole external token verbatim because
      // this module has no workbook resolver for it.
      const external = matchRefToken(formula, structured.end);
      if (external?.sheetQual) {
        out += formula.slice(i, external.end);
        i = external.end;
        continue;
      }
      out += structured.text;
      i = structured.end;
      continue;
    }
    // Only attempt a match at a boundary (prev char not an identifier
    // continuation), so we never split a function/defined name.
    const prev = i > 0 ? (formula[i - 1] ?? '') : '';
    if (!isIdentChar(prev) && prev !== "'") {
      const tok = matchRefToken(formula, i);
      if (tok) {
        const rep = visit(tok);
        out += rep ?? '#REF!';
        i = tok.end;
        continue;
      }
    }
    out += ch;
    i += 1;
  }
  return out;
}

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
        const col = colLabelToIndex(at.label);
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
      const col = colLabelToIndex(at.label);
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
        { abs: tok.a.abs, index: colLabelToIndex(tok.a.label) },
        { abs: tok.b.abs, index: colLabelToIndex(tok.b.label) },
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
      const col = colLabelToIndex(at.label);
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
 * shifted axis move. Cross-sheet refs are left untouched.
 */
export function adjustFormulaForCellBandShift(
  formula: string,
  affected: { r0: number; c0: number; r1: number; c1: number },
  axis: 'down' | 'right' | 'up' | 'left',
  delta: number,
): string {
  if (delta === 0) return formula;
  const vertical = axis === 'down' || axis === 'up';
  return rewriteRefs(formula, (tok) => {
    if (tok.sheetQual) return renderToken(tok);
    if (tok.kind !== 'cell') return renderToken(tok);
    const shiftAtom = (at: Atom): EndpointResult => {
      const col = colLabelToIndex(at.label);
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
      if (aResult.kind === 'ref') return null;
      return renderAtom(tok.a.absCol, aResult.col, tok.a.absRow, aResult.row);
    }
    return clampRange(
      tok,
      shiftAtom,
      vertical ? 'row' : 'col',
      vertical ? affected.r0 : affected.c0,
      delta < 0,
    );
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
    const col = colLabelToIndex(at.label);
    const row = Number.parseInt(at.rowStr, 10) - 1;
    if (row < source.r0 || row > source.r1 || col < source.c0 || col > source.c1) {
      return renderAtomRaw(at);
    }
    return renderAtom(at.absCol, col + dCol, at.absRow, row + dRow);
  };
  return rewriteRefs(formula, (tok) => {
    if (tok.sheetQual) return renderToken(tok);
    if (tok.kind !== 'cell') return renderToken(tok);
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
      // Partial cell cuts do not move whole-axis references, but a moved
      // formula still needs to retain the sheet binding of an unqualified
      // whole-axis ref when its output sheet changes.
      if (!tok.sheetQual) {
        return `${qualifierForOutputSheet(context.formulaSheet, context)}${renderToken(tok)}`;
      }
      return renderToken(tok);
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
  const aCol = colLabelToIndex(a.label);
  const bRow = Number.parseInt(b.rowStr, 10) - 1;
  const bCol = colLabelToIndex(b.label);
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
    colLabelToIndex(a.label) >= cut.c0 &&
    colLabelToIndex(a.label) <= cut.c1;
  const bInside =
    Number.parseInt(b.rowStr, 10) - 1 >= cut.r0 &&
    Number.parseInt(b.rowStr, 10) - 1 <= cut.r1 &&
    colLabelToIndex(b.label) >= cut.c0 &&
    colLabelToIndex(b.label) <= cut.c1;
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
  let aCol = colLabelToIndex(a.label);
  let aRow = Number.parseInt(a.rowStr, 10) - 1;
  let bCol = colLabelToIndex(b.label);
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
  const col = colLabelToIndex(at.label);
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

function resolveReferenceSheet(sheetQual: string, context: CutPasteSheetContext): number | null {
  if (!sheetQual) return context.formulaSheet;
  const body = sheetQual.slice(0, -1);
  // 3-D references and external workbook qualifiers are deliberately kept
  // verbatim until their own workbook-aware transform exists.
  if (body.includes(':') || body.includes('[') || body.includes(']')) return null;
  const name =
    body.startsWith("'") && body.endsWith("'") ? body.slice(1, -1).replace(/''/g, "'") : body;
  const folded = name.toLocaleLowerCase();
  const index = context.sheetNames.findIndex(
    (sheetName) => sheetName.toLocaleLowerCase() === folded,
  );
  return index >= 0 ? index : null;
}

function qualifierForOutputSheet(sheet: number, context: CutPasteSheetContext): string {
  if (sheet === context.outputSheet) return '';
  const name = context.sheetNames[sheet];
  if (name === undefined) return '';
  const safe = /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name);
  return `${safe ? name : `'${name.replace(/'/g, "''")}'`}!`;
}

/** Re-render a token verbatim from its parsed parts (used to pass through
 *  cross-sheet refs unchanged while keeping normalization consistent). */
function renderToken(tok: RefToken): string {
  if (tok.kind === 'whole-col') {
    return `${tok.sheetQual}${renderWholeColRaw(tok.a)}:${renderWholeColRaw(tok.b)}`;
  }
  if (tok.kind === 'whole-row') {
    return `${tok.sheetQual}${renderWholeRowRaw(tok.a)}:${renderWholeRowRaw(tok.b)}`;
  }
  const a = renderAtomRaw(tok.a);
  if (!tok.b) return `${tok.sheetQual}${a}`;
  return `${tok.sheetQual}${a}:${renderAtomRaw(tok.b)}`;
}

function renderAtomRaw(at: Atom): string {
  return `${at.absCol ? '$' : ''}${at.label}${at.absRow ? '$' : ''}${at.rowStr}`;
}

function renderWholeColRaw(at: WholeColAtom): string {
  return `${at.abs ? '$' : ''}${at.label}`;
}

function renderWholeRowRaw(at: WholeRowAtom): string {
  return `${at.abs ? '$' : ''}${at.rowStr}`;
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
    const col = colLabelToIndex(at.label);
    const row = Number.parseInt(at.rowStr, 10) - 1;
    const otherAxis =
      axis === 'row' ? Number.parseInt(other.rowStr, 10) - 1 : colLabelToIndex(other.label);
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
