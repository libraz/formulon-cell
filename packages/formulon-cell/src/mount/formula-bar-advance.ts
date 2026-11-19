/** Selection movement after a formula-bar commit, and Tab ownership for the bar. */
import { inputOperation } from '../commands/cell-input-commit.js';
import { interactionControllerFor } from '../commands/interaction-controller.js';
import { isCellWritable } from '../commands/protection.js';
import {
  isNavigationAddrAllowed,
  navigationBoundsFor,
  navigationSelectionBoundsFor,
  nextTabStop,
} from '../interact/navigation-policy.js';
import { advanceAfterCommit, type CommitAdvance } from '../interact/selection-navigation.js';
import type { SpreadsheetStore } from '../store/store.js';
import { mutators } from '../store/store.js';

export function moveAfterCommit(
  store: SpreadsheetStore,
  advance: CommitAdvance,
  mac: boolean,
): void {
  if (advance === 'none') return;
  const controller = interactionControllerFor(store);
  if (controller?.policy?.selection === false) return;
  const hasTabStops =
    controller?.policy !== undefined ||
    navigationBoundsFor(store) !== undefined ||
    navigationSelectionBoundsFor(store) !== undefined;
  // Mac Tab honours explicit tab stops ahead of the selected rectangle.
  if (mac && (advance === 'right' || advance === 'left') && hasTabStops) {
    const target = nextTabStop(store, store.getState().selection.active, advance === 'left');
    if (target && isNavigationAddrAllowed(store, target)) mutators.setActive(store, target);
    return;
  }
  advanceAfterCommit(store, advance, mac);
}

/** Whether the bar owns Tab: false when the active cell cannot be edited from here. */
export function tabCaptured(store: SpreadsheetStore, input: HTMLTextAreaElement): boolean {
  const controller = interactionControllerFor(store);
  if (controller?.policy?.selection === false || input.readOnly) return false;
  const state = store.getState();
  const a = state.selection.active;
  if (!isCellWritable(state, a)) return false;
  if (controller?.policy === undefined) return true;
  return controller.canExecute({
    operation: inputOperation(input.value),
    origin: 'editor',
    effects: [{ kind: 'cells', cells: [a] }],
  }).allowed;
}
