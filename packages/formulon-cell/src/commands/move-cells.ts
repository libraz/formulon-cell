import { MAX_COL, MAX_ROW, parseAddrKey } from '../engine/address.js';
import type { Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { mutators, type SpreadsheetStore, type State } from '../store/store.js';
import { cut } from './clipboard/cut.js';
import { createCopyFormatResolver } from './clipboard/format-resolver.js';
import { pasteBatchWrite } from './clipboard/paste-actions.js';
import { pasteSpecial } from './clipboard/paste-special.js';
import { captureSnapshotFromCopyResult } from './clipboard/snapshot.js';
import type { History } from './history.js';

export interface MoveCellsDeps {
  store: SpreadsheetStore;
  wb: WorkbookHandle;
  history: History | null;
  /** The gesture's command id; the move is authorized as an `editor` paste under it. */
  commandId: string;
  /** Asked before the move overwrites data outside the source. */
  confirmOverwrite: () => Promise<boolean>;
}

/** The source shifted so its top-left lands on `topLeft`. */
export const moveTargetRange = (src: Range, topLeft: { row: number; col: number }): Range => ({
  sheet: src.sheet,
  r0: topLeft.row,
  c0: topLeft.col,
  r1: topLeft.row + (src.r1 - src.r0),
  c1: topLeft.col + (src.c1 - src.c0),
});

/** Whether `range` lies entirely inside the sheet grid. */
export const moveTargetFits = (range: Range): boolean => range.r1 <= MAX_ROW && range.c1 <= MAX_COL;

/** Whether `dest` holds a value or formula outside `src`. */
const overwritesData = (state: State, src: Range, dest: Range): boolean => {
  for (const [key, cell] of state.data.cells) {
    if (cell.formula === null && cell.value.kind === 'blank') continue;
    const addr = parseAddrKey(key);
    if (!addr || addr.sheet !== dest.sheet) continue;
    if (addr.row < dest.r0 || addr.row > dest.r1 || addr.col < dest.c0 || addr.col > dest.c1) {
      continue;
    }
    if (addr.row >= src.r0 && addr.row <= src.r1 && addr.col >= src.c0 && addr.col <= src.c1) {
      continue;
    }
    return true;
  }
  return false;
};

/**
 * Drag-move the selection so its top-left lands on `topLeft`: one cut-paste
 * transaction, so formulas pointing at the moved cells follow them and Undo
 * restores both ranges. It is authorized as a paste over the target and the cleared source. Dropping onto data asks first. Resolves true when
 * the block moved.
 */
export async function moveSelectionTo(
  deps: MoveCellsDeps,
  topLeft: { row: number; col: number },
): Promise<boolean> {
  const { store, wb, history } = deps;
  const state = store.getState();
  if ((state.selection.extraRanges?.length ?? 0) > 0) return false;
  const src = state.selection.range;
  const dest = moveTargetRange(src, topLeft);
  if (dest.r0 === src.r0 && dest.c0 === src.c0) return false;
  if (!moveTargetFits(dest)) return false;
  if (overwritesData(state, src, dest) && !(await deps.confirmOverwrite())) return false;

  const current = store.getState();
  const captured = cut(current, wb);
  if (!captured) return false;
  const snap = captureSnapshotFromCopyResult(
    current,
    captured,
    'cut',
    createCopyFormatResolver(wb, current),
  );
  if (!snap) return false;
  mutators.setActive(store, { sheet: dest.sheet, row: dest.r0, col: dest.c0 });
  const written = pasteBatchWrite(
    { store, history, origin: 'editor', commandId: deps.commandId },
    { destination: dest, cutSource: src },
    () =>
      pasteSpecial(
        store.getState(),
        store,
        wb,
        snap,
        { what: 'all', operation: 'none', skipBlanks: false, transpose: false },
        history,
      ),
  ).range;
  if (!written) {
    mutators.setActive(store, { sheet: src.sheet, row: src.r0, col: src.c0 });
    mutators.setRange(store, src);
    return false;
  }
  mutators.replaceCells(store, wb.cells(dest.sheet));
  mutators.setRange(store, written);
  return true;
}
