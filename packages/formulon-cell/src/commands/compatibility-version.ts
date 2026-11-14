import type { CompatibilityVersionCode } from '../engine/calc-mode.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { Strings } from '../i18n/strings.js';
import type { History } from './history.js';

/** Calculation Options › Compatibility Version entries, in menu order. */
export const COMPATIBILITY_VERSIONS: readonly CompatibilityVersionCode[] = [1, 2, 3];

type CompatibilityVersionTemplateKey =
  | 'compatibilityVersionItem'
  | 'compatibilityVersionRecommended'
  | 'compatibilityVersionUnsupported';

/** The label template each entry takes. */
const VERSION_TEMPLATES: Readonly<
  Record<CompatibilityVersionCode, CompatibilityVersionTemplateKey>
> = {
  1: 'compatibilityVersionItem',
  2: 'compatibilityVersionRecommended',
  3: 'compatibilityVersionUnsupported',
};

/** The drawn label of a Compatibility Version entry; `shown` is the number it names, which for
 *  the unsupported entry is the book's own stamp. */
export const compatibilityVersionLabel = (
  text: Pick<Strings['ribbonMenu'], CompatibilityVersionTemplateKey>,
  version: CompatibilityVersionCode,
  shown: number = version,
): string => text[VERSION_TEMPLATES[version]].replace('{n}', String(shown));

/** How a Compatibility Version entry shows while the workbook holds the stamp `current`. */
export interface CompatibilityVersionRow {
  /** Version 1 and 2 always; the unsupported entry only on a book stamped 3 or higher. */
  readonly listed: boolean;
  readonly checked: boolean;
  /** The number the entry names: the book's own stamp for the unsupported entry. */
  readonly shown: number;
  /** See {@link compatibilityStampLocked}. */
  readonly locked: boolean;
}

/** Whether the stamp `current` is one the setter cannot set back (above 3), so a change away
 *  from it could not be undone. */
export const compatibilityStampLocked = (current: number | null): boolean =>
  current !== null && current > 3;

export const compatibilityVersionRow = (
  version: CompatibilityVersionCode,
  current: number | null,
): CompatibilityVersionRow => {
  const unsupported = current !== null && current >= 3;
  return {
    listed: version !== 3 || unsupported,
    checked: version === 3 ? unsupported : current === version,
    shown: version === 3 && unsupported ? current : version,
    locked: compatibilityStampLocked(current),
  };
};

/** Chooses a Compatibility Version as one undo step; undo sets the previous value back. The held
 *  value records nothing. Returns whether the workbook changed. */
export function chooseCompatibilityVersion(
  wb: WorkbookHandle,
  history: History | null,
  version: CompatibilityVersionCode,
): boolean {
  const previous = wb.compatibilityVersion();
  if (previous === null || previous === version) return false;
  // The setter cannot restore a stamp above 3.
  if (previous !== 1 && previous !== 2 && previous !== 3) return false;
  if (!wb.setCompatibilityVersion(version)) return false;
  if (history && !history.isReplaying()) {
    history.push({
      undo: () => {
        wb.setCompatibilityVersion(previous);
      },
      redo: () => {
        wb.setCompatibilityVersion(version);
      },
    });
  }
  return true;
}
