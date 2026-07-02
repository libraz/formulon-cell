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

/** A reference token surfaced in a formula: an optional sheet qualifier plus a
 *  cell atom and, for ranges, a second atom. */
interface RefToken {
  /** Raw sheet-qualifier text including the trailing `!` (e.g. `Sheet2!`,
   *  `'My Sheet'!`, `Sheet1:Sheet3!`), or '' when the ref is unqualified. */
  sheetQual: string;
  a: Atom;
  b: Atom | null;
  /** Character offset just past the whole token. */
  end: number;
}

const isLetter = (c: string): boolean => (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z');
const isDigit = (c: string): boolean => c >= '0' && c <= '9';
const isIdentChar = (c: string): boolean => isLetter(c) || isDigit(c) || c === '_';

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
 *  digits and dots after the first char, but must start with a letter or `_`. */
function parseSheetWord(src: string, start: number): number | null {
  let i = start;
  const first = src[i] ?? '';
  if (!(isLetter(first) || first === '_')) return null;
  i += 1;
  while (i < src.length) {
    const c = src[i] ?? '';
    if (isLetter(c) || isDigit(c) || c === '_' || c === '.') i += 1;
    else break;
  }
  return i;
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
  if (isIdentChar(next)) return null;
  // Reject tokens outside the grid (`Year2024`, `SHEET`) — those are defined
  // names / words, not cell references, and must pass through untouched.
  if (!atomInGrid(first.atom)) return null;
  if (b && !atomInGrid(b)) return null;
  return { sheetQual, a: first.atom, b, end: i };
}

/** True when an atom addresses a cell inside the grid (a real ref, not a name
 *  like `Year2024` whose "column" exceeds the last column). */
function atomInGrid(at: Atom): boolean {
  const col = colLabelToIndex(at.label);
  const row = Number.parseInt(at.rowStr, 10) - 1;
  return col >= 0 && col <= MAX_COL_INDEX && row >= 0 && row <= MAX_ROW_INDEX;
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

/** Render an atom to A1 text, or null when it falls outside the grid. */
function renderAtom(absCol: boolean, col: number, absRow: boolean, row: number): string | null {
  if (col < 0 || row < 0 || col > MAX_COL_INDEX || row > MAX_ROW_INDEX) return null;
  return `${absCol ? '$' : ''}${colIndexToLabel(col)}${absRow ? '$' : ''}${row + 1}`;
}

/** The outcome of transforming one endpoint against a structural edit. */
type EndpointResult = { kind: 'keep'; col: number; row: number } | { kind: 'ref' }; // fully inside a deleted band → #REF!

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
 * are left as their original text so the engine surfaces `#REF!`.
 */
export function shiftFormulaRefs(formula: string, dRow: number, dCol: number): string {
  if (!formula.startsWith('=') || (dRow === 0 && dCol === 0)) return formula;
  return rewriteRefs(formula, (tok) => {
    const shiftAtom = (at: Atom): string | null => {
      const col = colLabelToIndex(at.label);
      const row = Number.parseInt(at.rowStr, 10) - 1;
      const nc = at.absCol ? col : col + dCol;
      const nr = at.absRow ? row : row + dRow;
      return renderAtom(at.absCol, nc, at.absRow, nr);
    };
    const aTxt = shiftAtom(tok.a);
    // Out-of-grid shift: keep the original text (the engine surfaces #REF! when
    // it re-parses) rather than eagerly rewriting.
    if (aTxt === null) return renderToken(tok);
    if (!tok.b) return `${tok.sheetQual}${aTxt}`;
    const bTxt = shiftAtom(tok.b);
    if (bTxt === null) return renderToken(tok);
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
    const adjust = (at: Atom): EndpointResult => {
      const col = colLabelToIndex(at.label);
      const row = Number.parseInt(at.rowStr, 10) - 1;
      if (axis === 'row') {
        if (at.absRow || row < split) return { kind: 'keep', col, row };
        if (delta < 0 && row < split - delta) return { kind: 'ref' };
        return { kind: 'keep', col, row: row + delta };
      }
      if (at.absCol || col < split) return { kind: 'keep', col, row };
      if (delta < 0 && col < split - delta) return { kind: 'ref' };
      return { kind: 'keep', col: col + delta, row };
    };
    if (!tok.b) {
      const r = adjust(tok.a);
      if (r.kind === 'ref') return null;
      return `${renderAtom(tok.a.absCol, r.col, tok.a.absRow, r.row)}`;
    }
    // Range: clamp endpoints rather than dropping the whole ref for a partial
    // deletion.
    return clampRange(tok, adjust, axis, split);
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
    const shiftAtom = (at: Atom): string | null => {
      const col = colLabelToIndex(at.label);
      const row = Number.parseInt(at.rowStr, 10) - 1;
      let nr = row;
      let nc = col;
      if (
        vertical &&
        !at.absRow &&
        row >= affected.r0 &&
        col >= affected.c0 &&
        col <= affected.c1
      ) {
        nr = row + delta;
      } else if (
        !vertical &&
        !at.absCol &&
        col >= affected.c0 &&
        row >= affected.r0 &&
        row <= affected.r1
      ) {
        nc = col + delta;
      }
      return renderAtom(at.absCol, nc, at.absRow, nr);
    };
    const aTxt = shiftAtom(tok.a);
    if (aTxt === null) return null;
    if (!tok.b) return aTxt;
    const bTxt = shiftAtom(tok.b);
    if (bTxt === null) return null;
    return `${aTxt}:${bTxt}`;
  });
}

/**
 * Update formulas outside a cut/paste payload so references that pointed at
 * the moved cells follow them to the destination. Unlike copy/fill shifting,
 * absolute markers do not pin a moved-cell reference: `$A$1` should become
 * `$C$3` when the cell it points at is cut from A1 to C3.
 */
export function adjustFormulaForCutPasteMove(
  formula: string,
  source: { r0: number; c0: number; r1: number; c1: number },
  dest: { r0: number; c0: number },
): string {
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
    const aTxt = moveAtom(tok.a);
    if (aTxt === null) return null;
    if (!tok.b) return aTxt;
    const bTxt = moveAtom(tok.b);
    if (bTxt === null) return null;
    return `${aTxt}:${bTxt}`;
  });
}

/** Re-render a token verbatim from its parsed parts (used to pass through
 *  cross-sheet refs unchanged while keeping normalization consistent). */
function renderToken(tok: RefToken): string {
  const a = renderAtomRaw(tok.a);
  if (!tok.b) return `${tok.sheetQual}${a}`;
  return `${tok.sheetQual}${a}:${renderAtomRaw(tok.b)}`;
}

function renderAtomRaw(at: Atom): string {
  return `${at.absCol ? '$' : ''}${at.label}${at.absRow ? '$' : ''}${at.rowStr}`;
}

/** Clamp a range against a row/column edit so a partial deletion keeps the
 *  surviving span instead of turning the whole reference into `#REF!`. */
function clampRange(
  tok: RefToken,
  adjust: (at: Atom) => EndpointResult,
  axis: 'row' | 'col',
  split: number,
): string | null {
  const b = tok.b as Atom;
  const ra = adjust(tok.a);
  const rb = adjust(b);
  if (ra.kind === 'ref' && rb.kind === 'ref') return null; // whole range deleted
  // One endpoint deleted → clamp it to the boundary that survives.
  const resolve = (r: EndpointResult, at: Atom): { col: number; row: number } => {
    if (r.kind === 'keep') return { col: r.col, row: r.row };
    // Deleted endpoint → clamp to the first surviving line at/after the band.
    const col = colLabelToIndex(at.label);
    const row = Number.parseInt(at.rowStr, 10) - 1;
    if (axis === 'row') return { col, row: split };
    return { col: split, row };
  };
  const pa = resolve(ra, tok.a);
  const pb = resolve(rb, b);
  const aTxt = renderAtom(tok.a.absCol, pa.col, tok.a.absRow, pa.row);
  const bTxt = renderAtom(b.absCol, pb.col, b.absRow, pb.row);
  if (aTxt === null || bTxt === null) return null;
  return `${aTxt}:${bTxt}`;
}
