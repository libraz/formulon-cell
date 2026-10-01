import type { WorkbookHandle } from '../../engine/workbook-handle.js';
import type { State } from '../../store/store.js';
import { type CopyResult, copy } from './copy.js';

/**
 * Capture the selection as a cut payload. The source stays intact until the
 * payload is successfully pasted; this is how Excel permits Escape to cancel
 * a cut and how overlapping cut moves preserve the captured values.
 */
export function cut(state: State, wb: WorkbookHandle): CopyResult | null {
  // `wb` remains in the signature for compatibility with command callers and
  // future native clipboard implementations. A cut is deliberately capture
  // only; pasteSpecial performs the deferred source clear in one transaction.
  void wb;
  return copy(state);
}
