import type { Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { addMergeToMaps } from '../store/merge-maps.js';
import { isSingleCell } from '../store/selection-geometry.js';
import { withSparklineMembers } from '../store/sparkline-slice.js';
import {
  type ConditionalRule,
  mutators,
  type SpreadsheetStore,
  type State,
} from '../store/store.js';
import {
  shiftFilterCriteria,
  shiftFilterHiddenButtons,
  shiftFilterSorts,
  shiftRangeAxis,
} from './axis-shift.js';
import { recordFilterChange } from './filter-history.js';
import type { History } from './history.js';
import {
  recordConditionalRulesChange,
  recordMergesChange,
  recordMergesChangeWithEngine,
  recordSparklineChange,
} from './slice-history.js';
import { shiftSparklinesForAxisEdit } from './sparkline.js';

function mergeRangesForSheet(state: State, sheet: number): Range[] {
  return [...state.merges.byAnchor.values()]
    .filter((range) => range.sheet === sheet)
    .map((range) => ({ ...range }));
}

function sameMergeRanges(a: readonly Range[], b: readonly Range[]): boolean {
  if (a.length !== b.length) return false;
  return a.every((left, index) => {
    const right = b[index];
    return (
      right !== undefined &&
      left.sheet === right.sheet &&
      left.r0 === right.r0 &&
      left.c0 === right.c0 &&
      left.r1 === right.r1 &&
      left.c1 === right.c1
    );
  });
}

function syncMergesToEngine(wb: WorkbookHandle, sheet: number, ranges: readonly Range[]): void {
  if (!wb.capabilities.merges) return;
  wb.engineClearMerges(sheet);
  for (const range of ranges) wb.engineAddMerge(sheet, range);
}

/** Re-point merges, conditional-format ranges, and the autofilter region after
 *  a row/col insert or delete so they track the cells they annotate. Each
 *  concern records its own history entry inside the surrounding transaction. */
export function shiftAnchoredRanges(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  history: History | null,
  sheet: number,
  axis: 'row' | 'col',
  split: number,
  delta: number,
  nativeAxisOp: boolean,
): void {
  // Merges: rebuild both lookup maps from the shifted anchors. A merge fully
  // inside a deleted band, or collapsed to a single cell, is dropped.
  const mutateMerges = (): void => {
    store.setState((s) => {
      const byAnchor = new Map<string, Range>();
      const byCell = new Map<string, string>();
      for (const merge of s.merges.byAnchor.values()) {
        const shifted = merge.sheet === sheet ? shiftRangeAxis(merge, axis, split, delta) : merge;
        if (!shifted) continue;
        if (isSingleCell(shifted)) continue;
        addMergeToMaps(byAnchor, byCell, shifted);
      }
      return { ...s, merges: { byAnchor, byCell } };
    });
  };
  // A native axis edit already moves merges inside the engine. During undo the
  // merge history entry runs before the inverse axis operation, so syncing the
  // engine here would make the native operation move the restored ranges a
  // second time. The native axis history entry restores the exact engine merge
  // snapshot after its inverse; the JS fallback still mirrors the store here.
  if (nativeAxisOp) {
    recordMergesChange(history, store, mutateMerges);
    const afterMerges = mergeRangesForSheet(store.getState(), sheet);
    // Native engines generally move merges correctly, but a merge reduced to
    // one cell is still reported by some versions. Excel removes that merge,
    // so reconcile the native result with the store's range normalization.
    if (wb.capabilities.merges && !sameMergeRanges(wb.getMerges(sheet), afterMerges)) {
      syncMergesToEngine(wb, sheet, afterMerges);
      if (history && !history.isReplaying()) {
        history.push({
          undo: () => {},
          redo: () => syncMergesToEngine(wb, sheet, afterMerges),
        });
      }
    }
  } else {
    recordMergesChangeWithEngine(history, store, wb, sheet, mutateMerges);
  }

  // Conditional formats: shift each rule's range; drop rules whose range is
  // fully consumed by a deletion.
  recordConditionalRulesChange(history, store, () => {
    store.setState((s) => {
      const rules: ConditionalRule[] = [];
      for (const rule of s.conditional.rules) {
        if (rule.range.sheet !== sheet) {
          rules.push(rule);
          continue;
        }
        const shifted = shiftRangeAxis(rule.range, axis, split, delta);
        if (!shifted) continue;
        rules.push({ ...rule, range: shifted });
      }
      return { ...s, conditional: { rules } };
    });
  });

  // Sparklines: host cells move with their rows/columns and sources follow data.
  // The engine shifts its own groups, and the native op re-reads them.
  if (!(nativeAxisOp && wb.capabilities.sparklines)) {
    recordSparklineChange(history, store, () => {
      store.setState((s) => {
        if (s.sparkline.sparklines.size === 0) return s;
        const sparklines = shiftSparklinesForAxisEdit(
          s.sparkline.sparklines,
          sheet,
          wb.sheetName(sheet),
          axis,
          split,
          delta,
        );
        return { ...s, sparkline: withSparklineMembers(s.sparkline, sparklines) };
      });
    });
  }

  // Copy marquee. Transient UI state that history never captures, so it is
  // shifted outside the recorders — otherwise the outline would keep pointing
  // at the source's pre-shift indices. A cut payload keeps its original source
  // coordinates until paste; a successful structure edit invalidates that
  // payload, so Excel cancels the cut marquee instead of shifting it.
  const copyUi = store.getState().ui;
  if (copyUi.copyMode === 'cut') {
    if (copyUi.copyRanges) mutators.setCopyRanges(store, null);
    else mutators.setCopyRange(store, null);
  } else {
    store.setState((s) => {
      const shiftMarquee = (range: Range): Range | null =>
        range.sheet === sheet ? shiftRangeAxis(range, axis, split, delta) : range;
      const copyRange = s.ui.copyRange ? shiftMarquee(s.ui.copyRange) : null;
      const copyRanges = s.ui.copyRanges
        ? s.ui.copyRanges.map(shiftMarquee).filter((r): r is Range => r !== null)
        : null;
      if (copyRange === s.ui.copyRange && copyRanges === s.ui.copyRanges) return s;
      return {
        ...s,
        ui: { ...s.ui, copyRange, copyRanges: copyRanges?.length ? copyRanges : null },
      };
    });
  }

  // Autofilter region + per-column criteria.
  recordFilterChange(history, store, () => {
    store.setState((s) => {
      if (!s.ui.filterRanges.some((range) => range.sheet === sheet)) return s;
      // A filter whose rows or columns were all deleted goes away with them.
      const filterRanges = s.ui.filterRanges.flatMap((range) => {
        if (range.sheet !== sheet) return [range];
        const shifted = shiftRangeAxis(range, axis, split, delta);
        return shifted ? [shifted] : [];
      });
      return {
        ...s,
        ui: {
          ...s.ui,
          filterRanges,
          filterCriteria: shiftFilterCriteria(s.ui.filterCriteria, sheet, axis, split, delta),
          filterSorts: shiftFilterSorts(s.ui.filterSorts, sheet, axis, split, delta),
          filterHiddenButtons: shiftFilterHiddenButtons(
            s.ui.filterHiddenButtons,
            sheet,
            axis,
            split,
            delta,
          ),
        },
      };
    });
  });
}
