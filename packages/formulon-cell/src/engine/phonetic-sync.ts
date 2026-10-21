import type { KanaForm } from '../format/kana.js';
import type { CellFormat } from '../store/store.js';
import type { Addr, PhoneticProperties } from './types.js';
import type { WorkbookHandle } from './workbook-handle.js';

/** Kana forms by the engine's `<phoneticPr type>` code. */
export const PHONETIC_KANA_FORMS: readonly KanaForm[] = [
  'halfKatakana',
  'katakana',
  'hiragana',
  'asStored',
];

export type PhoneticAlignment = 'normal' | 'left' | 'center' | 'distributed';

/** Alignments by the engine's `<phoneticPr alignment>` code. */
export const PHONETIC_ALIGNMENTS: readonly PhoneticAlignment[] = [
  'normal',
  'left',
  'center',
  'distributed',
];

/** The kana form a guide is shown in; unset takes the `<phoneticPr>` schema
 *  default, full-width katakana. */
export const phoneticKanaForm = (properties: PhoneticProperties | undefined): KanaForm =>
  PHONETIC_KANA_FORMS[properties?.type ?? 1] ?? 'asStored';

/** How a guide sits over its span; unset takes the schema default, left. */
export const phoneticAlignment = (properties: PhoneticProperties | undefined): PhoneticAlignment =>
  PHONETIC_ALIGNMENTS[properties?.alignment ?? 1] ?? 'normal';

/** Font size of a guide over text drawn at `fontSize`. */
export const phoneticGuideSize = (fontSize: number): number =>
  Math.max(6, Math.round(fontSize * 0.52));

/** Engine properties for a kana form and alignment, keeping `fontId`. */
export const phoneticPropertiesFor = (
  form: Exclude<KanaForm, 'asStored'>,
  alignment: PhoneticAlignment,
  fontId = 0,
): PhoneticProperties => ({
  fontId,
  type: PHONETIC_KANA_FORMS.indexOf(form),
  alignment: PHONETIC_ALIGNMENTS.indexOf(alignment),
});

/** Read how a cell's guide renders; a guide with no `<phoneticPr>` hydrates as
 *  "unset" rather than as the all-zero triple the engine resolves it to. */
export function readPhoneticProperties(
  wb: WorkbookHandle,
  addr: Addr,
): PhoneticProperties | undefined {
  if (!wb.capabilities.phoneticProperties) return undefined;
  const properties = wb.getCellPhoneticProperties(addr.sheet, addr.row, addr.col);
  if (properties === null) return undefined;
  const { present, fontId, type, alignment } = properties;
  return present ? { fontId, type, alignment } : undefined;
}

/** Add the cell's Show Phonetic flag to a hydrate patch; off is the absence. */
export function withShowPhonetic(
  wb: WorkbookHandle,
  addr: Addr,
  patch: Partial<CellFormat>,
): Partial<CellFormat> {
  if (!wb.capabilities.showPhonetic) return patch;
  if (wb.getCellShowPhonetic(addr.sheet, addr.row, addr.col) === true) patch.showPhonetic = true;
  return patch;
}

/** Add the cell's guide properties to a hydrate patch that carries its runs. */
export function withPhoneticProperties(
  wb: WorkbookHandle,
  addr: Addr,
  patch: Partial<CellFormat>,
): Partial<CellFormat> {
  const properties = readPhoneticProperties(wb, addr);
  if (properties) patch.phoneticProperties = properties;
  return patch;
}

/** What a new guide gets: the schema default the desktop spreadsheet writes.
 *  The engine's unset state is half-width katakana, so it is stated outright. */
export const DEFAULT_PHONETIC_PROPERTIES: PhoneticProperties = phoneticPropertiesFor(
  'katakana',
  'left',
);

/**
 * Write a cell's guide properties after its runs, the default ones when the
 * cell has none. Returns false only when the engine refused a write; a cell
 * without runs is a no-op.
 */
export function writePhoneticProperties(
  wb: WorkbookHandle,
  addr: Addr,
  format: Pick<CellFormat, 'phonetic' | 'phoneticProperties'>,
): boolean {
  if (!format.phonetic?.length || !wb.capabilities.phoneticProperties) return true;
  const properties = format.phoneticProperties ?? DEFAULT_PHONETIC_PROPERTIES;
  return wb.setCellPhoneticProperties(addr.sheet, addr.row, addr.col, properties);
}

/**
 * Bring a cell's Show Phonetic flag in line with its format. The engine keeps
 * the flag on constants and blanks only, so a formula cell is left alone.
 * Returns false only when the engine refused a write.
 */
export function writeShowPhonetic(
  wb: WorkbookHandle,
  addr: Addr,
  format: Pick<CellFormat, 'showPhonetic'>,
  isFormula: boolean,
): boolean {
  if (isFormula || !wb.capabilities.showPhonetic) return true;
  const show = format.showPhonetic === true;
  if (wb.getCellShowPhonetic(addr.sheet, addr.row, addr.col) === show) return true;
  return wb.setCellShowPhonetic(addr.sheet, addr.row, addr.col, show);
}
