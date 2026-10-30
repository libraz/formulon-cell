import { addrKey } from '../../engine/address.js';
import type { CellPatch, WorkbookHandle } from '../../engine/workbook-handle.js';
import type { History } from '../history.js';
import { snapshotToPatch } from '../prepared-change.js';

/** Apply `patches` in one engine transaction and, when anything changed,
 *  record a single undo/redo child on `history`. */
export function applyPatchesAtomic(
  wb: WorkbookHandle,
  patches: readonly CellPatch[],
  history: History | null,
): void {
  const atomic = wb.applyCellPatchAtomic(patches);
  if (atomic.changed.length === 0 || !history || history.isReplaying()) return;
  const changed = new Set(atomic.changed.map(addrKey));
  const before = atomic.before
    .filter((snapshot) => changed.has(addrKey(snapshot.addr)))
    .map(snapshotToPatch);
  const after = atomic.after
    .filter((snapshot) => changed.has(addrKey(snapshot.addr)))
    .map(snapshotToPatch);
  history.push({
    undo: () => {
      wb.applyCellPatchAtomic(before);
    },
    redo: () => {
      wb.applyCellPatchAtomic(after);
    },
  });
}
