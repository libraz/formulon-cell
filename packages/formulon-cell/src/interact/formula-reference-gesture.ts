import {
  type FormulaReferenceInputOptions,
  insertFormulaReference,
} from '../commands/formula-reference-input.js';

export interface FormulaReferenceGesture {
  /** Replaces the original formula span with the current reference or range. */
  insert(ref: string): { text: string; caret: number };
}

/** Capture the text and selection at the start of a pointer gesture. Repeated
 *  drag updates must all use this snapshot so the selected endpoint's original
 *  prefix and suffix remain intact. */
export function createFormulaReferenceGesture(
  text: string,
  selectionStart: number,
  selectionEnd: number,
  options: FormulaReferenceInputOptions = {},
): FormulaReferenceGesture {
  const originalText = text;
  const originalSelectionStart = selectionStart;
  const originalSelectionEnd = selectionEnd;
  const originalOptions = { ...options };

  return {
    insert(ref) {
      return insertFormulaReference(
        originalText,
        originalSelectionStart,
        originalSelectionEnd,
        ref,
        originalOptions,
      );
    },
  };
}
