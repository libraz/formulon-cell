import { inputLocaleForWorkbook } from '../commands/coerce-input.js';
import type { WorkbookHandle } from './workbook-handle.js';

/**
 * Text comparison used by cell-layer operations when the native workbook
 * cannot evaluate the operation itself.  Keeping the locale explicit is
 * deliberate: `localeCompare()` without a locale inherits the browser/host
 * language, which makes the same workbook sort differently after it moves
 * between hosts.
 */
export interface ProfileTextSemantics {
  readonly locale: string;
  readonly compare: (left: string, right: string) => number;
  readonly equals: (left: string, right: string) => boolean;
  /** A linear-time key for checklist membership. It folds case and Unicode
   * width while retaining the selected locale. The native workbook route is
   * preferred whenever its dynamic-array evaluator is available. */
  readonly key: (value: string) => string;
}

const safeLocale = (locale: string | undefined): string => {
  const requested = locale?.trim();
  if (!requested) return 'en-US';
  try {
    // Intl.Locale is not available in every embedded browser, while the
    // Collator constructor is. This validation also prevents a malformed
    // host locale from throwing during a sort action.
    new Intl.Collator(requested);
    return requested;
  } catch {
    return 'en-US';
  }
};

export const localeForWorkbook = (wb: WorkbookHandle | null | undefined): string => {
  if (!wb) return 'en-US';
  try {
    return safeLocale(inputLocaleForWorkbook(wb));
  } catch {
    return 'en-US';
  }
};

export function createProfileTextSemantics(locale = 'en-US'): ProfileTextSemantics {
  const resolved = safeLocale(locale);
  const collator = new Intl.Collator(resolved, {
    usage: 'sort',
    numeric: false,
    sensitivity: 'base',
  });
  const key = (value: string): string => {
    // NFKC folds full-width/half-width forms (a common spreadsheet checklist
    // expectation) before applying the locale's case mapping. Collator is
    // still the authority for ordering/equality; this key only keeps large
    // value checklists linear instead of comparing every pair.
    return value.normalize('NFKC').toLocaleLowerCase(resolved);
  };
  return {
    locale: resolved,
    compare: (left, right) => collator.compare(left, right),
    equals: (left, right) => collator.compare(left, right) === 0,
    key,
  };
}

export const profileTextKey = (value: string, locale = 'en-US'): string =>
  createProfileTextSemantics(locale).key(value);
