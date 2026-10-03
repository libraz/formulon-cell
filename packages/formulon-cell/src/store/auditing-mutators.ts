import { addrKey } from '../engine/address.js';
import type { Addr, Range } from '../engine/types.js';
import { rangeArea } from './selection-geometry.js';
import type { SpreadsheetStore } from './store.js';
import type { TraceArrow } from './types.js';

const MAX_WATCH_RANGE_CELLS = 10_000;

/** Formula-auditing state: watch window, trace arrows, ignored errors and
 *  invalid-data circles. `addWatchRange` stays in `store.ts` because it
 *  dispatches through the composed `mutators` object. */
export const auditingMutators = {
  /** Pin `addr` to the Watch Window. No-op when the same sheet/row/col is
   *  already watched — duplicate entries would render redundant rows. */
  addWatch(store: SpreadsheetStore, addr: Addr): void {
    store.setState((s) => {
      const exists = s.watch.watches.some(
        (w) => w.sheet === addr.sheet && w.row === addr.row && w.col === addr.col,
      );
      if (exists) return s;
      return {
        ...s,
        watch: { watches: [...s.watch.watches, { ...addr }] },
      };
    });
  },

  /** Pin every cell in each range to the Watch Window, preserving row-major
   *  order across ranges while de-duplicating already watched addresses. */
  addWatchRanges(store: SpreadsheetStore, ranges: readonly Range[]): void {
    if (ranges.some((range) => rangeArea(range) > MAX_WATCH_RANGE_CELLS)) return;
    store.setState((s) => {
      const seen = new Set(s.watch.watches.map((w) => `${w.sheet}:${w.row}:${w.col}`));
      const next: Addr[] = [...s.watch.watches];
      let changed = false;
      for (const range of ranges) {
        for (let row = range.r0; row <= range.r1; row += 1) {
          for (let col = range.c0; col <= range.c1; col += 1) {
            const key = `${range.sheet}:${row}:${col}`;
            if (seen.has(key)) continue;
            seen.add(key);
            next.push({ sheet: range.sheet, row, col });
            changed = true;
          }
        }
      }
      if (!changed) return s;
      return { ...s, watch: { watches: next } };
    });
  },

  /** Unpin `addr` from the Watch Window. No-op when not present. */
  removeWatch(store: SpreadsheetStore, addr: Addr): void {
    store.setState((s) => {
      const next = s.watch.watches.filter(
        (w) => !(w.sheet === addr.sheet && w.row === addr.row && w.col === addr.col),
      );
      if (next.length === s.watch.watches.length) return s;
      return { ...s, watch: { watches: next } };
    });
  },

  /** Drop every watched cell. */
  clearWatches(store: SpreadsheetStore): void {
    store.setState((s) => (s.watch.watches.length === 0 ? s : { ...s, watch: { watches: [] } }));
  },

  /** Replace the Watch Window entries. */
  setWatches(store: SpreadsheetStore, watches: readonly Addr[]): void {
    store.setState((s) => ({
      ...s,
      watch: { watches: watches.map((addr) => ({ ...addr })) },
    }));
  },

  /** Show or hide the Watch Window panel. */
  setWatchPanelOpen(store: SpreadsheetStore, open: boolean): void {
    store.setState((s) =>
      s.ui.watchPanelOpen === open ? s : { ...s, ui: { ...s.ui, watchPanelOpen: open } },
    );
  },

  /** Append a trace arrow to the visible set. Duplicates (same kind +
   *  identical endpoints) are dropped so repeated `tracePrecedents()` calls
   *  on the same active cell don't pile up overlapping arrows. */
  addTrace(store: SpreadsheetStore, item: TraceArrow): void {
    store.setState((s) => {
      const exists = s.traces.items.some(
        (t) =>
          t.kind === item.kind &&
          t.from.sheet === item.from.sheet &&
          t.from.row === item.from.row &&
          t.from.col === item.from.col &&
          t.to.sheet === item.to.sheet &&
          t.to.row === item.to.row &&
          t.to.col === item.to.col,
      );
      if (exists) return s;
      return {
        ...s,
        traces: { items: [...s.traces.items, { kind: item.kind, from: item.from, to: item.to }] },
      };
    });
  },

  /** Empty the trace-arrow set. */
  clearTraces(store: SpreadsheetStore): void {
    store.setState((s) => (s.traces.items.length === 0 ? s : { ...s, traces: { items: [] } }));
  },

  /** Suppress the error-indicator triangle for `addr` for the rest of the
   *  session. Idempotent — re-ignoring a cell is a no-op. NOT history-tracked. */
  ignoreError(store: SpreadsheetStore, addr: Addr): void {
    store.setState((s) => {
      const key = addrKey(addr);
      if (s.errorIndicators.ignoredErrors.has(key)) return s;
      const next = new Set(s.errorIndicators.ignoredErrors);
      next.add(key);
      return { ...s, errorIndicators: { ...s.errorIndicators, ignoredErrors: next } };
    });
  },

  /** Re-enable the error-indicator triangle for `addr` when it was ignored. */
  unignoreError(store: SpreadsheetStore, addr: Addr): void {
    store.setState((s) => {
      const key = addrKey(addr);
      if (!s.errorIndicators.ignoredErrors.has(key)) return s;
      const next = new Set(s.errorIndicators.ignoredErrors);
      next.delete(key);
      return { ...s, errorIndicators: { ...s.errorIndicators, ignoredErrors: next } };
    });
  },

  /** Drop every ignored-error suppression. */
  clearIgnoredErrors(store: SpreadsheetStore): void {
    store.setState((s) =>
      s.errorIndicators.ignoredErrors.size === 0
        ? s
        : { ...s, errorIndicators: { ...s.errorIndicators, ignoredErrors: new Set() } },
    );
  },

  /** Replace the current ignored-error suppressions. */
  setIgnoredErrors(store: SpreadsheetStore, keys: Set<string>): void {
    store.setState((s) => ({
      ...s,
      errorIndicators: { ...s.errorIndicators, ignoredErrors: new Set(keys) },
    }));
  },

  /** Replace the current "Circle Invalid Data" marks. */
  setValidationCircles(store: SpreadsheetStore, keys: Set<string>): void {
    store.setState((s) => ({
      ...s,
      errorIndicators: { ...s.errorIndicators, validationCircles: new Set(keys) },
    }));
  },

  /** Clear every visible "Circle Invalid Data" mark. */
  clearValidationCircles(store: SpreadsheetStore): void {
    store.setState((s) =>
      s.errorIndicators.validationCircles.size === 0
        ? s
        : { ...s, errorIndicators: { ...s.errorIndicators, validationCircles: new Set() } },
    );
  },
};
