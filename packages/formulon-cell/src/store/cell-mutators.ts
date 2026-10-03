import { addrKey } from '../engine/address.js';
import type { Addr, CellValue, Range } from '../engine/types.js';
import { navigationBoundsFor, syncNavigationViewport } from '../interact/navigation-policy.js';
import type { SpreadsheetStore } from './store.js';
import type { CellFormat, CustomCellStyle } from './types.js';

/** Cell values, the active sheet's data, per-cell formats and merges. */
export const cellMutators = {
  setCell(
    store: SpreadsheetStore,
    addr: Addr,
    value: CellValue,
    formula: string | null = null,
  ): void {
    store.setState((s) => {
      const cells = new Map(s.data.cells);
      if (value.kind === 'blank' && !formula) cells.delete(addrKey(addr));
      else cells.set(addrKey(addr), { value, formula });
      return { ...s, data: { ...s.data, cells } };
    });
  },

  replaceCells(
    store: SpreadsheetStore,
    entries: Iterable<{ addr: Addr; value: CellValue; formula: string | null }>,
  ): void {
    const cells = new Map<string, { value: CellValue; formula: string | null }>();
    for (const e of entries) cells.set(addrKey(e.addr), { value: e.value, formula: e.formula });
    store.setState((s) => ({ ...s, data: { ...s.data, cells } }));
  },

  /** Switch the active sheet index. Cells must be re-hydrated separately
   *  via `replaceCells` after calling this. Resets selection to A1 on the
   *  new sheet. */
  setSheetIndex(store: SpreadsheetStore, idx: number): void {
    const fixedRange = navigationBoundsFor(store);
    if (fixedRange && fixedRange.sheet !== idx) return;
    store.setState((s) => ({
      ...s,
      data: { ...s.data, sheetIndex: idx, cells: new Map() },
      ui: { ...s.ui, pendingFormat: null },
      selection: {
        active: { sheet: idx, row: 0, col: 0 },
        anchor: { sheet: idx, row: 0, col: 0 },
        range: { sheet: idx, r0: 0, c0: 0, r1: 0, c1: 0 },
      },
    }));
    syncNavigationViewport(store);
  },

  /** Merge a partial format into the cell at `addr`. Pass `null` to clear. */
  setCellFormat(store: SpreadsheetStore, addr: Addr, patch: Partial<CellFormat> | null): void {
    store.setState((s) => {
      const formats = new Map(s.format.formats);
      const key = addrKey(addr);
      if (patch === null) {
        formats.delete(key);
      } else {
        const prev = formats.get(key) ?? {};
        const next: CellFormat = { ...prev, ...patch };
        if (patch.borders) next.borders = { ...(prev.borders ?? {}), ...patch.borders };
        formats.set(key, next);
      }
      return { ...s, format: { ...s.format, formats } };
    });
  },

  /** Apply `patch` to every cell in `range`. Pass `null` to clear.
   *  Skips no-op when the range is huge (full-row/column or selectAll) — until
   *  we have row/column-level format storage, painting per-cell entries for
   *  millions of empty cells would OOM. */
  setRangeFormat(store: SpreadsheetStore, range: Range, patch: Partial<CellFormat> | null): void {
    const area = (range.r1 - range.r0 + 1) * (range.c1 - range.c0 + 1);
    if (area > 100_000) return;
    store.setState((s) => {
      const formats = new Map(s.format.formats);
      const sheet = range.sheet;
      for (let r = range.r0; r <= range.r1; r += 1) {
        for (let c = range.c0; c <= range.c1; c += 1) {
          const key = addrKey({ sheet, row: r, col: c });
          if (patch === null) {
            formats.delete(key);
          } else {
            const prev = formats.get(key) ?? {};
            const next: CellFormat = { ...prev, ...patch };
            if (patch.borders) next.borders = { ...(prev.borders ?? {}), ...patch.borders };
            formats.set(key, next);
          }
        }
      }
      return { ...s, format: { ...s.format, formats } };
    });
  },

  /** Merge a range into a single cell. The top-left becomes the anchor;
   *  every other cell in the range is mapped back to that anchor. If the
   *  range is 1×1 it's a no-op. */
  mergeRange(store: SpreadsheetStore, range: Range): void {
    if (range.r0 === range.r1 && range.c0 === range.c1) return;
    store.setState((s) => {
      const byAnchor = new Map(s.merges.byAnchor);
      const byCell = new Map(s.merges.byCell);
      // Strip any existing merges that touch the range.
      for (const [anchorKey, r] of byAnchor) {
        if (
          r.sheet === range.sheet &&
          !(r.r1 < range.r0 || r.r0 > range.r1 || r.c1 < range.c0 || r.c0 > range.c1)
        ) {
          byAnchor.delete(anchorKey);
          for (let row = r.r0; row <= r.r1; row += 1) {
            for (let col = r.c0; col <= r.c1; col += 1) {
              byCell.delete(addrKey({ sheet: r.sheet, row, col }));
            }
          }
        }
      }
      const anchor = { sheet: range.sheet, row: range.r0, col: range.c0 };
      const ak = addrKey(anchor);
      byAnchor.set(ak, range);
      for (let row = range.r0; row <= range.r1; row += 1) {
        for (let col = range.c0; col <= range.c1; col += 1) {
          if (row === range.r0 && col === range.c0) continue;
          byCell.set(addrKey({ sheet: range.sheet, row, col }), ak);
        }
      }
      return { ...s, merges: { byAnchor, byCell } };
    });
  },

  upsertCustomCellStyle(store: SpreadsheetStore, style: CustomCellStyle): void {
    store.setState((s) => {
      const next = (s.format.customCellStyles ?? []).filter((item) => item.id !== style.id);
      return {
        ...s,
        format: {
          ...s.format,
          customCellStyles: [...next, { ...style, format: { ...style.format } }],
        },
      };
    });
  },

  /** Remove any merges that intersect the range. */
  unmergeRange(store: SpreadsheetStore, range: Range): void {
    store.setState((s) => {
      const byAnchor = new Map(s.merges.byAnchor);
      const byCell = new Map(s.merges.byCell);
      for (const [anchorKey, r] of byAnchor) {
        if (
          r.sheet === range.sheet &&
          !(r.r1 < range.r0 || r.r0 > range.r1 || r.c1 < range.c0 || r.c0 > range.c1)
        ) {
          byAnchor.delete(anchorKey);
          for (let row = r.r0; row <= r.r1; row += 1) {
            for (let col = r.c0; col <= r.c1; col += 1) {
              byCell.delete(addrKey({ sheet: r.sheet, row, col }));
            }
          }
        }
      }
      return { ...s, merges: { byAnchor, byCell } };
    });
  },
};
