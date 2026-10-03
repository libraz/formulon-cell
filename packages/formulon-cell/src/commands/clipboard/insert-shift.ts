import { addrKey, MAX_COL, MAX_ROW } from '../../engine/address.js';
import type { Addr, Range } from '../../engine/types.js';
import { writeCell } from '../../engine/value.js';
import type { WorkbookHandle } from '../../engine/workbook-handle.js';
import { addMergeToMaps, removeIntersectingMerges } from '../../store/merge-maps.js';
import type { CellFormat, SpreadsheetStore } from '../../store/store.js';
import type { History } from '../history.js';
import { isCellWritable } from '../protection.js';
import { shiftFormulaRefs } from '../refs.js';
import { recordFormatChange, recordMergesChangeWithEngine } from '../slice-history.js';
import type { ClipboardSnapshot } from './snapshot.js';

export function writeSnapshotIntoInsertedRange(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  history: History | null,
  snapshot: ClipboardSnapshot,
  origin: Addr,
): void {
  const formatWrites: { key: string; format: CellFormat | null }[] = [];
  for (let r = 0; r < snapshot.rows; r += 1) {
    for (let c = 0; c < snapshot.cols; c += 1) {
      const src = snapshot.cells[r]?.[c];
      if (!src) continue;
      const addr: Addr = { sheet: origin.sheet, row: origin.row + r, col: origin.col + c };
      if (!isCellWritable(store.getState(), addr)) continue;
      if (src.formula) {
        const formula =
          snapshot.mode === 'cut'
            ? src.formula
            : shiftFormulaRefs(
                src.formula,
                addr.row - (snapshot.range.r0 + r),
                addr.col - (snapshot.range.c0 + c),
              );
        wb.setFormula(addr, formula);
      } else {
        writeCell(wb, addr, src.value, null);
      }
      formatWrites.push({
        key: addrKey(addr),
        format: src.format
          ? { ...src.format, borders: src.format.borders ? { ...src.format.borders } : undefined }
          : null,
      });
    }
  }
  if (formatWrites.length === 0) return;
  recordFormatChange(history, store, () => {
    store.setState((s) => {
      const formats = new Map(s.format.formats);
      for (const { key, format } of formatWrites) {
        if (format) formats.set(key, format);
        else formats.delete(key);
      }
      return { ...s, format: { ...s.format, formats } };
    });
  });
}

export function copySnapshotMerges(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  history: History | null,
  origin: Addr,
  snapshot: ClipboardSnapshot | null | undefined,
): void {
  if (!snapshot || snapshot.merges?.length === 0) return;
  const sourceMerges = snapshot.merges ?? [];
  recordMergesChangeWithEngine(history, store, wb, origin.sheet, () => {
    store.setState((s) => {
      const byAnchor = new Map(s.merges.byAnchor);
      const byCell = new Map(s.merges.byCell);
      for (const merge of sourceMerges) {
        if (
          !Number.isInteger(merge.r0) ||
          !Number.isInteger(merge.c0) ||
          !Number.isInteger(merge.r1) ||
          !Number.isInteger(merge.c1) ||
          merge.r0 < 0 ||
          merge.c0 < 0 ||
          merge.r1 < merge.r0 ||
          merge.c1 < merge.c0 ||
          merge.r1 >= snapshot.rows ||
          merge.c1 >= snapshot.cols ||
          (merge.sheet !== undefined && merge.sheet !== snapshot.range.sheet)
        ) {
          continue;
        }
        const next: Range = {
          sheet: origin.sheet,
          r0: origin.row + merge.r0,
          c0: origin.col + merge.c0,
          r1: origin.row + merge.r1,
          c1: origin.col + merge.c1,
        };
        if (next.r1 > MAX_ROW || next.c1 > MAX_COL) continue;
        removeIntersectingMerges(byAnchor, byCell, next);
        addMergeToMaps(byAnchor, byCell, next);
      }
      return { ...s, merges: { byAnchor, byCell } };
    });
  });
}
