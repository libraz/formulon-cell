import { addrKey, parseAddrKey } from '../engine/address.js';
import {
  type PhoneticAlignment,
  phoneticAlignment,
  phoneticKanaForm,
  phoneticPropertiesFor,
} from '../engine/phonetic-sync.js';
import type { Addr, PhoneticRun, Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { KanaForm } from '../format/kana.js';
import { rangeContainsAddr } from '../store/selection-geometry.js';
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

/** Kana form and alignment chosen in the phonetic settings dialog. */
export interface PhoneticSettings {
  readonly form: Exclude<KanaForm, 'asStored'>;
  readonly alignment: PhoneticAlignment;
}

/** The settings the dialog opens on: the active cell's, else the desktop
 *  default of full-width katakana, left aligned. */
export function phoneticSettingsAt(store: SpreadsheetStore, addr: Addr): PhoneticSettings {
  const properties = store.getState().format.formats.get(addrKey(addr))?.phoneticProperties;
  const form = phoneticKanaForm(properties);
  const alignment = phoneticAlignment(properties);
  if (!properties || form === 'asStored') return { form: 'katakana', alignment: 'left' };
  return { form, alignment };
}

/**
 * Apply `settings` to every cell in `range` that carries a phonetic guide; the
 * properties are only observable on a cell with runs. Returns the number of
 * cells changed — zero when the engine has no property surface.
 */
export function setPhoneticSettings(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  range: Range,
  settings: PhoneticSettings,
): number {
  if (!wb.capabilities.phoneticProperties) return 0;
  let changed = 0;
  for (const [key, format] of store.getState().format.formats) {
    if (!format.phonetic?.length) continue;
    const addr = parseAddrKey(key);
    if (!addr || !rangeContainsAddr(range, addr)) continue;
    const properties = phoneticPropertiesFor(
      settings.form,
      settings.alignment,
      format.phoneticProperties?.fontId ??
        wb.getCellPhoneticProperties(addr.sheet, addr.row, addr.col)?.fontId,
    );
    if (!wb.setCellPhoneticProperties(addr.sheet, addr.row, addr.col, properties)) continue;
    mutators.setCellFormat(store, addr, { phoneticProperties: properties });
    changed++;
  }
  return changed;
}
