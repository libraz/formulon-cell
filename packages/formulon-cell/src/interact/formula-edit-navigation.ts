import { rewriteRefs } from '../commands/formula-ref-lexer.js';
import { resolveFormulaReferenceInsertionSpan } from '../commands/formula-reference-input.js';
import { r1c1AtomToA1 } from '../commands/refs.js';
import { MAX_COL, MAX_ROW, parseA1Atom } from '../engine/address.js';
import type { Addr } from '../engine/types.js';
import type { State } from '../store/types.js';
import {
  createFormulaReferenceGesture,
  type FormulaReferenceGesture,
} from './formula-reference-gesture.js';
import { rangeRefOf, refOf } from './range-insert.js';
import { stepVisibleWithMerge } from './visible-navigation.js';

export type FormulaEditArrow = 'up' | 'down' | 'left' | 'right';

export interface FormulaReferenceNavigationState {
  readonly gesture: FormulaReferenceGesture;
  readonly value: string;
  readonly caret: number;
  readonly anchor: Addr;
  readonly tip: Addr;
  readonly r1c1: boolean;
}

export interface NavigateFormulaReferenceInput {
  readonly state: State;
  readonly text: string;
  readonly selectionStart: number;
  readonly selectionEnd: number;
  readonly editedCell: Addr;
  readonly direction: FormulaEditArrow;
  readonly extend: boolean;
  readonly current?: FormulaReferenceNavigationState | null;
  readonly r1c1?: boolean;
}

const R1C1_QUALIFIED = /^((?:'(?:[^']|'')*'|[\p{L}_][\p{L}\p{N}_.]*)!)?([^!]+)$/u;

/** Read an unqualified `ref` as endpoints on the edited cell's sheet. Returns
 *  null for anything else, including a sheet-qualified reference. */
function parsePickedReference(
  ref: string,
  editedCell: Addr,
  r1c1: boolean,
): { anchor: Addr; tip: Addr } | null {
  let atoms: string[] | null = null;
  if (r1c1) {
    const match = R1C1_QUALIFIED.exec(ref);
    if (!match || match[1]) return null;
    const parts = (match[2] ?? '').split(/\s*:\s*/u);
    if (parts.length > 2) return null;
    atoms = parts.map((part) => r1c1AtomToA1(part, editedCell));
  } else {
    rewriteRefs(`=${ref}`, (token, start) => {
      if (start !== 1 || token.end !== ref.length + 1 || token.sheetQual) return null;
      if (token.kind !== 'cell') return null;
      atoms = [`${token.a.label}${token.a.rowStr}`];
      if (token.b) atoms.push(`${token.b.label}${token.b.rowStr}`);
      return null;
    });
  }
  const cells = (atoms ?? []).map((atom) => parseA1Atom(atom));
  const [first, second] = cells;
  if (!first || (cells.length === 2 && !second)) return null;
  const toAddr = (cell: { row: number; col: number }): Addr => ({
    ...cell,
    sheet: editedCell.sheet,
  });
  return { anchor: toAddr(first), tip: toAddr(second ?? first) };
}

/** Navigation state for a reference just inserted by a pointer pick, so the
 *  next arrow continues from the picked cell or range. Null when `ref` is not
 *  a plain reference on the edited cell's sheet. */
export function pickedReferenceNavigation(input: {
  gesture: FormulaReferenceGesture;
  inserted: { text: string; caret: number };
  ref: string;
  editedCell: Addr;
  r1c1: boolean;
}): FormulaReferenceNavigationState | null {
  const { gesture, inserted, ref, editedCell, r1c1 } = input;
  if (inserted.text.slice(inserted.caret - ref.length, inserted.caret) !== ref) return null;
  const picked = parsePickedReference(ref, editedCell, r1c1);
  if (!picked) return null;
  return { gesture, value: inserted.text, caret: inserted.caret, ...picked, r1c1 };
}

function axisStep(input: NavigateFormulaReferenceInput): { dRow: number; dCol: number } {
  if (input.direction === 'up') return { dRow: -1, dCol: 0 };
  if (input.direction === 'down') return { dRow: 1, dCol: 0 };
  const visualDirection = input.direction === 'right' ? 1 : -1;
  return {
    dRow: 0,
    dCol: input.state.ui.rightToLeft ? -visualDirection : visualDirection,
  };
}

/** Move the active endpoint of a formula reference without changing the
 *  workbook selection. Each fresh gesture starts at the edited cell, and
 *  repeated moves replace the formula span captured by the first move. */
export function navigateFormulaReference(
  input: NavigateFormulaReferenceInput,
): FormulaReferenceNavigationState | null {
  const { text, selectionStart, selectionEnd, editedCell, state } = input;
  const caret = Math.max(0, Math.min(selectionStart, text.length));
  const current = input.current ?? null;
  if (current) {
    if (selectionStart !== selectionEnd || current.value !== text || current.caret !== caret)
      return null;
  }

  const r1c1 = input.r1c1 === true;
  let gesture: FormulaReferenceGesture;
  let anchor: Addr;
  let tip: Addr;
  if (current) {
    gesture = current.gesture;
    anchor = current.anchor;
    tip = current.tip;
  } else {
    const options = { r1c1, requireExpressionSelection: true };
    const span = resolveFormulaReferenceInsertionSpan(text, selectionStart, selectionEnd, options);
    if (!span) return null;
    gesture = createFormulaReferenceGesture(text, selectionStart, selectionEnd, options);
    anchor = editedCell;
    tip = editedCell;
  }

  const { dRow, dCol } = axisStep(input);
  const destination = stepVisibleWithMerge(state, tip, dRow, dCol, MAX_ROW, MAX_COL);
  if (
    destination.row === tip.row &&
    destination.col === tip.col &&
    current &&
    (input.extend ||
      (current.anchor.sheet === current.tip.sheet &&
        current.anchor.row === current.tip.row &&
        current.anchor.col === current.tip.col))
  )
    return current;

  if (input.extend) {
    tip = destination;
  } else {
    anchor = destination;
    tip = destination;
  }

  const mode = { r1c1, base: editedCell };
  const ref =
    anchor.row === tip.row && anchor.col === tip.col
      ? refOf(tip.row, tip.col, mode)
      : rangeRefOf(anchor, tip, mode);
  const inserted = gesture.insert(ref);
  return {
    gesture,
    value: inserted.text,
    caret: inserted.caret,
    anchor,
    tip,
    r1c1,
  };
}
