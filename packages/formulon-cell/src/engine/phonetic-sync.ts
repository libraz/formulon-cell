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

/**
 * Read how a cell's guide renders. The all-zero triple is what the engine
 * reports for a guide with no `<phoneticPr>`, so it hydrates as "unset".
 */
export function readPhoneticProperties(
  wb: WorkbookHandle,
  addr: Addr,
): PhoneticProperties | undefined {
  if (!wb.capabilities.phoneticProperties) return undefined;
  const properties = wb.getCellPhoneticProperties(addr.sheet, addr.row, addr.col);
  if (properties === null) return undefined;
  const { fontId, type, alignment } = properties;
  return fontId === 0 && type === 0 && alignment === 0 ? undefined : properties;
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

/**
 * Write a cell's guide properties after its runs. Returns false only when the
 * engine refused a write; a cell with nothing to state is a no-op.
 */
export function writePhoneticProperties(
  wb: WorkbookHandle,
  addr: Addr,
  format: Pick<CellFormat, 'phonetic' | 'phoneticProperties'>,
): boolean {
  const properties = format.phoneticProperties;
  if (!properties || !format.phonetic?.length || !wb.capabilities.phoneticProperties) return true;
  return wb.setCellPhoneticProperties(addr.sheet, addr.row, addr.col, properties);
}
