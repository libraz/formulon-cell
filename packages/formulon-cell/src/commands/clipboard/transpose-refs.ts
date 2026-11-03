import { colFromLetters } from '../../engine/address.js';
import type { Range } from '../../engine/types.js';
import { type Atom, renderAtom, renderToken, rewriteRefs } from '../formula-ref-lexer.js';
import { shiftFormulaRefs } from '../formula-refs.js';

const atomRow = (at: Atom): number => Number.parseInt(at.rowStr, 10) - 1;
const atomCol = (at: Atom): number => colFromLetters(at.label);

const insideBlock = (at: Atom, block: Range): boolean => {
  const row = atomRow(at);
  const col = atomCol(at);
  return row >= block.r0 && row <= block.r1 && col >= block.c0 && col <= block.c1;
};

/**
 * Rewrite a formula copied with Transpose. A reference whose cells all lie in
 * the copied block follows those cells to their transposed position in the
 * pasted tile — `$` markers included, so `=$A$1` from A1:D4 pasted at N100
 * reads `=$N$100`, and `=SUM(A1:A2)` reads `=SUM(N100:O100)`. Every other
 * reference shifts by the cell's own paste offset, as an ordinary copy does.
 */
export function transposeFormulaRefs(
  formula: string,
  block: Range,
  from: { row: number; col: number },
  to: { row: number; col: number },
  tileOrigin: { row: number; col: number },
): string {
  if (!formula.startsWith('=')) return formula;
  const dRow = to.row - from.row;
  const dCol = to.col - from.col;
  const transposed = (at: Atom): string | null =>
    renderAtom(
      at.absCol,
      tileOrigin.col + (atomRow(at) - block.r0),
      at.absRow,
      tileOrigin.row + (atomCol(at) - block.c0),
    );
  return rewriteRefs(formula, (tok) => {
    const inBlock =
      tok.kind === 'cell' &&
      tok.sheetQual === '' &&
      insideBlock(tok.a, block) &&
      (tok.b === null || insideBlock(tok.b, block));
    if (!inBlock || tok.kind !== 'cell') {
      return shiftFormulaRefs(`=${renderToken(tok)}`, dRow, dCol).slice(1);
    }
    if (!tok.b) return transposed(tok.a);
    // Both corners swap axes, so the top-left corner stays top-left.
    const first = transposed(tok.a);
    const second = transposed(tok.b);
    if (first === null || second === null) return null;
    return `${first}:${second}`;
  });
}
