/**
 * A1 reference tokenizer behind every transform in `formula-refs.ts`: token
 * shapes, the formula scanner that visits each reference, and the renderers
 * that turn transformed endpoints back into A1 text.
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
export interface Atom {
  absCol: boolean;
  label: string;
  absRow: boolean;
  rowStr: string;
}

/** A cell reference token surfaced in a formula. */
export interface CellRefToken {
  kind: 'cell';
  /** Raw sheet-qualifier text including the trailing `!` (e.g. `Sheet2!`,
   *  `'My Sheet'!`, `Sheet1:Sheet3!`), or '' when the ref is unqualified. */
  sheetQual: string;
  a: Atom;
  b: Atom | null;
  /** Character offset just past the whole token. */
  end: number;
}

export interface WholeColAtom {
  abs: boolean;
  label: string;
}

export interface WholeRowAtom {
  abs: boolean;
  rowStr: string;
}

export interface WholeColRefToken {
  kind: 'whole-col';
  sheetQual: string;
  a: WholeColAtom;
  b: WholeColAtom;
  end: number;
}

export interface WholeRowRefToken {
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
export function renderAtom(
  absCol: boolean,
  col: number,
  absRow: boolean,
  row: number,
): string | null {
  if (col < 0 || row < 0 || col > MAX_COL_INDEX || row > MAX_ROW_INDEX) return null;
  return `${absCol ? '$' : ''}${colIndexToLabel(col)}${absRow ? '$' : ''}${row + 1}`;
}

export function renderWholeColAtom(abs: boolean, col: number): string | null {
  if (col < 0 || col > MAX_COL_INDEX) return null;
  return `${abs ? '$' : ''}${colIndexToLabel(col)}`;
}

export function renderWholeRowAtom(abs: boolean, row: number): string | null {
  if (row < 0 || row > MAX_ROW_INDEX) return null;
  return `${abs ? '$' : ''}${row + 1}`;
}

/** Walk `formula`, replacing each reference token via `visit`. The visitor
 *  receives the token and returns the replacement text, or null to emit
 *  `#REF!` for the whole token. String literals, `#REF!` tokens, function
 *  names, and non-reference text are passed through untouched. */
export function rewriteRefs(formula: string, visit: (tok: RefToken) => string | null): string {
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

/** Re-render a token verbatim from its parsed parts (used to pass through
 *  cross-sheet refs unchanged while keeping normalization consistent). */
export function renderToken(tok: RefToken): string {
  return `${tok.sheetQual}${renderTokenBody(tok)}`;
}

export function renderTokenBody(tok: RefToken): string {
  if (tok.kind === 'whole-col') {
    return `${renderWholeColRaw(tok.a)}:${renderWholeColRaw(tok.b)}`;
  }
  if (tok.kind === 'whole-row') {
    return `${renderWholeRowRaw(tok.a)}:${renderWholeRowRaw(tok.b)}`;
  }
  const a = renderAtomRaw(tok.a);
  if (!tok.b) return a;
  return `${a}:${renderAtomRaw(tok.b)}`;
}

export function renderAtomRaw(at: Atom): string {
  return `${at.absCol ? '$' : ''}${at.label}${at.absRow ? '$' : ''}${at.rowStr}`;
}

function renderWholeColRaw(at: WholeColAtom): string {
  return `${at.abs ? '$' : ''}${at.label}`;
}

function renderWholeRowRaw(at: WholeRowAtom): string {
  return `${at.abs ? '$' : ''}${at.rowStr}`;
}
