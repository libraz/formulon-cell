import type { WorkbookHandle } from '../../engine/workbook-handle.js';
import { isWholeColumnRange, isWholeRowRange } from '../../store/selection-geometry.js';
import type { State } from '../../store/store.js';
import { inputLocaleForWorkbook, localizeFormulaForClipboard } from '../coerce-input.js';
import { type CopyResult, copy } from './copy.js';

/**
 * Capture the selection as a cut payload. The source stays intact until the
 * payload is successfully pasted; this is how desktop spreadsheets let Escape
 * cancel a cut and how overlapping cut moves preserve the captured values.
 */
export function cut(state: State, wb: WorkbookHandle): CopyResult | null {
  const range = state.selection.range;
  // A whole-sheet move is rejected: its trimmed payload cannot stand in for the full source clear.
  if (isWholeRowRange(range) && isWholeColumnRange(range)) return null;
  // A cut is deliberately capture only; pasteSpecial performs the deferred
  // source clear in one transaction. Keep its TSV spelling aligned with copy
  // while retaining the workbook argument used by existing callers.
  return copy(state, inputLocaleForWorkbook(wb), localizeFormulaForClipboard(wb));
}
