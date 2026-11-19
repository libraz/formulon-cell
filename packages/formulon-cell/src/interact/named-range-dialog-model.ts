import {
  type DefinedNameEntry,
  definedNameValueText,
  deleteDefinedName,
  isSameStoredEntry,
  isValidDefinedName,
  resolveDefinedNameKey,
  sanitizeName,
  upsertDefinedName,
} from '../commands/named-ranges.js';
import type { Addr, Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { Strings } from '../i18n/strings.js';
import { formatSheetAbsoluteRange } from '../wrappers/toolbar-a1.js';

export type NameFilter = 'all' | 'errors' | 'noErrors' | 'workbook';
export type NameSortKey = 'name' | 'value' | 'formula' | 'scope' | 'comment';

const hasFormulaError = (formula: string): boolean =>
  /#(?:REF|NAME|VALUE|DIV\/0|N\/A|NULL|NUM)!?/i.test(formula);

/** Map a defined-name mutation failure to its inline message. */
export function errorMessageFor(
  reason: 'empty-name' | 'invalid-name' | 'empty-formula' | 'unsupported' | 'engine-failed',
  t: Strings['namedRangeDialog'],
  refusal: Strings['commandRefusal'],
): string {
  if (reason === 'invalid-name') return t.errorInvalidName;
  if (reason === 'empty-name') return t.errorEmptyName;
  if (reason === 'empty-formula') return t.errorEmptyFormula;
  return refusal.engineFailed;
}

export const scopeLabel = (
  wb: WorkbookHandle,
  t: Strings['namedRangeDialog'],
  localSheetId: number,
): string => (localSheetId < 0 ? t.workbookScope : wb.sheetName(localSheetId));

export const scopeOptions = (
  wb: WorkbookHandle,
  t: Strings['namedRangeDialog'],
): { value: string; label: string }[] => [
  { value: '-1', label: t.workbookScope },
  ...Array.from({ length: wb.sheetCount }, (_, sheet) => ({
    value: String(sheet),
    label: wb.sheetName(sheet),
  })),
];

const nameSortValue = (
  wb: WorkbookHandle,
  t: Strings['namedRangeDialog'],
  entry: DefinedNameEntry,
  key: NameSortKey,
): string => {
  switch (key) {
    case 'formula':
      return entry.formula;
    case 'scope':
      return scopeLabel(wb, t, entry.localSheetId);
    case 'value':
      return definedNameValueText(wb, entry);
    case 'comment':
      return entry.comment ?? '';
    default:
      return wb.localizeDefinedName(entry.name);
  }
};

/** Hidden names already removed by the caller; filters, then sorts as the list shows them. */
export function selectNameRows(
  wb: WorkbookHandle,
  t: Strings['namedRangeDialog'],
  entries: readonly DefinedNameEntry[],
  filter: NameFilter,
  key: NameSortKey,
  dir: 'asc' | 'desc',
): DefinedNameEntry[] {
  return entries
    .filter((entry) => {
      if (filter === 'errors') return hasFormulaError(entry.formula);
      if (filter === 'noErrors') return !hasFormulaError(entry.formula);
      if (filter === 'workbook') return entry.localSheetId < 0;
      return true;
    })
    .sort((a, b) => {
      const left = nameSortValue(wb, t, a, key);
      const right = nameSortValue(wb, t, b, key);
      const result = left.localeCompare(right, undefined, { numeric: true, sensitivity: 'base' });
      return dir === 'asc' ? result : -result;
    });
}

/** Where the New Name form reads the selection from. */
export interface NameSeedSource {
  getSelection?: () => { active: Addr; range: Range };
  getSelectedRangeFormula?: () => string;
}

export function newNameSeed(
  wb: WorkbookHandle,
  source: NameSeedSource,
): { name: string; formula: string } {
  const selection = source.getSelection?.();
  if (!selection) return { name: '', formula: source.getSelectedRangeFormula?.() ?? '' };
  const { active, range } = selection;
  const formula = `=${formatSheetAbsoluteRange(wb.sheetName(range.sheet), range)}`;
  const textAt = (row: number, col: number): string => {
    if (row < 0 || col < 0) return '';
    const v = wb.getValue({ sheet: active.sheet, row, col });
    return v.kind === 'text' ? v.value.trim() : '';
  };
  const suggest = (text: string): string => {
    if (!text) return '';
    if (isValidDefinedName(text)) return text;
    const cleaned = sanitizeName(text);
    return isValidDefinedName(cleaned) ? cleaned : '';
  };
  const name =
    suggest(textAt(active.row, active.col)) ||
    suggest(textAt(active.row, active.col - 1)) ||
    suggest(textAt(active.row - 1, active.col));
  return { name, formula };
}

/** Add `name`; when renaming, drop the entry it replaces so no copy is left behind. */
export function replaceDefinedName(
  wb: WorkbookHandle,
  previous: DefinedNameEntry | null,
  name: string,
  formula: string,
  scope: number,
  comment: string,
): ReturnType<typeof upsertDefinedName> {
  const key = resolveDefinedNameKey(wb, name, formula, scope);
  const result = upsertDefinedName(wb, name, formula, scope, { comment });
  // The entry the write landed on is never removed, whatever spelling it was typed in.
  if (result.ok && previous !== null && !isSameStoredEntry(previous, key)) {
    const removed = deleteDefinedName(wb, previous.name, previous.localSheetId);
    if (!removed.ok) {
      deleteDefinedName(wb, result.entry.name, result.entry.localSheetId);
      const status = removed.reason === 'engine-failed' ? removed.status : null;
      return { ok: false, reason: 'engine-failed', status };
    }
  }
  return result;
}
