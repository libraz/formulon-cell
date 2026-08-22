import type { Addr, PhoneticRun } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';

/** The guide's readings run together, which is what a single-field editor
 *  shows and what OOXML's whole-cell form carries. */
export const phoneticReading = (runs: readonly PhoneticRun[] | undefined): string =>
  runs === undefined ? '' : runs.map((run) => run.text).join('');

/** The guide on `addr`, as the readings a single-field editor opens on. */
export function phoneticReadingAt(store: SpreadsheetStore, addr: Addr): string {
  const key = `${addr.sheet}:${addr.row}:${addr.col}`;
  return phoneticReading(store.getState().format.formats.get(key)?.phonetic);
}

/**
 * Apply a whole-cell reading typed into the phonetic editor. `reading` covers
 * the entire cell text, so writing it replaces every span the guide had —
 * which is right when the user edited the field and wrong when they did not.
 * A reading equal to `initial` is therefore left alone, so opening the editor
 * on a partially annotated cell and dismissing it with OK keeps the spans.
 *
 * Returns true when the guide changed, false when it was left as it was or the
 * engine refused the write.
 */
export function setPhoneticReading(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  addr: Addr,
  reading: string,
  initial: string,
): boolean {
  if (reading === initial) return false;
  // The engine spans the cell text itself, so the whole-cell entry point is
  // what turns a typed reading into a run without the host measuring anything.
  if (!wb.setCellPhonetic(addr.sheet, addr.row, addr.col, reading)) return false;
  mutators.setCellFormat(store, addr, { phonetic: readBackGuide(store, wb, addr, reading) });
  return true;
}

/** Prefer the engine's own view of the guide it just stored. An engine with no
 *  per-run surface leaves the host to spell the same thing itself: one run over
 *  the whole cell text. */
function readBackGuide(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  addr: Addr,
  reading: string,
): readonly PhoneticRun[] | undefined {
  const runs = wb.getCellPhoneticRuns(addr.sheet, addr.row, addr.col);
  if (runs !== null) return runs.length > 0 ? runs : undefined;
  if (!reading) return undefined;
  const cell = store.getState().data.cells.get(`${addr.sheet}:${addr.row}:${addr.col}`);
  const text = cell?.value.kind === 'text' ? cell.value.value : '';
  return [{ start: 0, end: text.length, text: reading }];
}
