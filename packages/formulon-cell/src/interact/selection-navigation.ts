import { interactionControllerFor } from '../commands/interaction-controller.js';
import { mergeAt, stepWithMerge } from '../commands/merge.js';
import { MAX_COL, MAX_ROW } from '../engine/address.js';
import type { Addr, Range } from '../engine/types.js';
import { mutators, type SpreadsheetStore, type State } from '../store/store.js';
import {
  clampNavigationAddr,
  isNavigationAddrAllowed,
  navigationBoundsFor,
  navigationPolicyFor,
} from './navigation-policy.js';

export type SelectionNavigationDirection = 'down' | 'up' | 'right' | 'left';

const contains = (range: Range, addr: Addr): boolean =>
  range.sheet === addr.sheet &&
  addr.row >= range.r0 &&
  addr.row <= range.r1 &&
  addr.col >= range.c0 &&
  addr.col <= range.c1;

const sameAddr = (a: Addr, b: Addr): boolean =>
  a.sheet === b.sheet && a.row === b.row && a.col === b.col;

const nextAddress = (range: Range, addr: Addr, direction: SelectionNavigationDirection): Addr => {
  if (direction === 'down' || direction === 'up') {
    const height = range.r1 - range.r0 + 1;
    const width = range.c1 - range.c0 + 1;
    const offset = (addr.col - range.c0) * height + (addr.row - range.r0);
    const area = height * width;
    const next = (offset + (direction === 'down' ? 1 : area - 1)) % area;
    return {
      sheet: range.sheet,
      row: range.r0 + (next % height),
      col: range.c0 + Math.floor(next / height),
    };
  }

  const height = range.r1 - range.r0 + 1;
  const width = range.c1 - range.c0 + 1;
  const offset = (addr.row - range.r0) * width + (addr.col - range.c0);
  const area = height * width;
  const next = (offset + (direction === 'right' ? 1 : area - 1)) % area;
  return {
    sheet: range.sheet,
    row: range.r0 + Math.floor(next / width),
    col: range.c0 + (next % width),
  };
};

const mergeIsContained = (range: Range, merge: Range): boolean =>
  merge.sheet === range.sheet &&
  merge.r0 >= range.r0 &&
  merge.r1 <= range.r1 &&
  merge.c0 >= range.c0 &&
  merge.c1 <= range.c1;

/** Jump over the rest of a merge in the requested logical traversal order. */
const afterMerge = (
  range: Range,
  merge: Range,
  direction: SelectionNavigationDirection,
  candidate: Addr,
): Addr => {
  if (direction === 'down') {
    if (merge.r1 < range.r1) return { sheet: range.sheet, row: merge.r1 + 1, col: candidate.col };
    if (merge.r0 === range.r0 && merge.r1 === range.r1) {
      return {
        sheet: range.sheet,
        row: range.r0,
        col: merge.c1 < range.c1 ? merge.c1 + 1 : range.c0,
      };
    }
    if (candidate.col < range.c1) {
      return { sheet: range.sheet, row: range.r0, col: candidate.col + 1 };
    }
    return { sheet: range.sheet, row: range.r0, col: range.c0 };
  }
  if (direction === 'up') {
    if (merge.r0 > range.r0) return { sheet: range.sheet, row: merge.r0 - 1, col: candidate.col };
    if (candidate.col > range.c0) {
      return { sheet: range.sheet, row: range.r1, col: candidate.col - 1 };
    }
    return { sheet: range.sheet, row: range.r1, col: range.c1 };
  }
  if (direction === 'right') {
    if (merge.c1 < range.c1) return { sheet: range.sheet, row: candidate.row, col: merge.c1 + 1 };
    if (merge.c0 === range.c0 && merge.c1 === range.c1) {
      return {
        sheet: range.sheet,
        row: merge.r1 < range.r1 ? merge.r1 + 1 : range.r0,
        col: range.c0,
      };
    }
    if (candidate.row < range.r1) {
      return { sheet: range.sheet, row: candidate.row + 1, col: range.c0 };
    }
    return { sheet: range.sheet, row: range.r0, col: range.c0 };
  }
  if (merge.c0 > range.c0) return { sheet: range.sheet, row: candidate.row, col: merge.c0 - 1 };
  if (candidate.row > range.r0) {
    return { sheet: range.sheet, row: candidate.row - 1, col: range.c1 };
  }
  return { sheet: range.sheet, row: range.r1, col: range.c1 };
};

/**
 * Return the next logical cell inside the primary selected rectangle.
 * Return/Shift+Return use column-major order; Tab/Shift+Tab use row-major
 * order. Traversal wraps and skips merged-cell bodies without materializing
 * the selected area.
 */
export function nextWithinSelection(
  state: State,
  direction: SelectionNavigationDirection,
): Addr | null {
  const range = state.selection.range;
  if (!contains(range, state.selection.active)) return null;
  if (range.r0 === range.r1 && range.c0 === range.c1) return null;

  const activeMerge = mergeAt(state, state.selection.active);
  const start = activeMerge
    ? { sheet: activeMerge.sheet, row: activeMerge.r0, col: activeMerge.c0 }
    : state.selection.active;
  if (!contains(range, start)) return null;

  // A selection occupied by exactly one merge has only one logical stop.
  if (
    activeMerge &&
    activeMerge.r0 === range.r0 &&
    activeMerge.r1 === range.r1 &&
    activeMerge.c0 === range.c0 &&
    activeMerge.c1 === range.c1
  ) {
    return null;
  }

  let candidate = nextAddress(range, start, direction);
  // Each jump clears at least one merged cell, so a full-sheet selection never scans its area.
  const maxMergeJumps = state.merges.byCell.size + state.merges.byAnchor.size + 1;
  for (let mergeJumps = 0; mergeJumps <= maxMergeJumps; mergeJumps += 1) {
    if (sameAddr(candidate, start)) return null;
    const merge = mergeAt(state, candidate);
    if (!merge) return candidate;

    const mergeAnchor = { sheet: merge.sheet, row: merge.r0, col: merge.c0 };
    if (
      mergeIsContained(range, merge) &&
      (sameAddr(candidate, mergeAnchor) ||
        (direction === 'up' && candidate.col === merge.c0) ||
        (direction === 'left' && candidate.row === merge.r0))
    ) {
      return mergeAnchor;
    }

    const clippedMerge: Range = {
      sheet: range.sheet,
      r0: Math.max(range.r0, merge.r0),
      c0: Math.max(range.c0, merge.c0),
      r1: Math.min(range.r1, merge.r1),
      c1: Math.min(range.c1, merge.c1),
    };
    candidate = afterMerge(range, clippedMerge, direction, candidate);
  }
  throw new Error('nextWithinSelection: merge traversal did not converge');
}

const STEP: Record<SelectionNavigationDirection, readonly [number, number]> = {
  down: [1, 0],
  up: [-1, 0],
  right: [0, 1],
  left: [0, -1],
};

/** Where Return/Tab moves the active cell, and whether the selected area stays intact. */
export interface SelectionAdvance {
  readonly addr: Addr;
  readonly preserveSelection: boolean;
}

/**
 * Resolve one Return/Tab step: inside the selected rectangle when
 * `traverseSelection` is set and the rectangle has another stop, otherwise a
 * merge-aware single step clamped to `bounds` (the sheet limits by default).
 */
export function nextAdvanceTarget(
  state: State,
  direction: SelectionNavigationDirection,
  traverseSelection: boolean,
  bounds?: Range,
): SelectionAdvance {
  if (traverseSelection) {
    const within = nextWithinSelection(state, direction);
    if (within) return { addr: within, preserveSelection: true };
  }
  const [dRow, dCol] = STEP[direction];
  return {
    addr: stepWithMerge(
      state,
      state.selection.active,
      dRow,
      dCol,
      bounds?.r1 ?? MAX_ROW,
      bounds?.c1 ?? MAX_COL,
    ),
    preserveSelection: false,
  };
}

/**
 * Move the active cell after a committed edit. Selected-rectangle traversal is
 * a Mac convention and yields to any interaction or navigation policy; plain
 * steps respect merges, navigation bounds, and disabled selection.
 */
export function advanceAfterCommit(
  store: SpreadsheetStore,
  direction: SelectionNavigationDirection,
  mac: boolean,
): void {
  const controller = interactionControllerFor(store);
  if (controller?.policy?.selection === false) return;
  const traverse =
    mac && controller?.policy === undefined && navigationPolicyFor(store)?.options === undefined;
  const next = nextAdvanceTarget(store.getState(), direction, traverse, navigationBoundsFor(store));
  if (next.preserveSelection) {
    mutators.setActivePreservingSelection(store, next.addr);
    return;
  }
  const clamped = clampNavigationAddr(store, next.addr);
  if (clamped && isNavigationAddrAllowed(store, clamped)) mutators.setActive(store, clamped);
}
