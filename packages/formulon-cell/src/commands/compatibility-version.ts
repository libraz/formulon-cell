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

/** The drawn label of a Compatibility Version entry. */
export const compatibilityVersionLabel = (
  text: Pick<Strings['ribbonMenu'], CompatibilityVersionTemplateKey>,
  version: CompatibilityVersionCode,
): string => text[VERSION_TEMPLATES[version]].replace('{n}', String(version));

/** Whether `version` is listed while the workbook holds `current`: version 3 only on a book
 *  stamped 3. */
export const compatibilityVersionListed = (
  version: CompatibilityVersionCode,
  current: CompatibilityVersionCode | null,
): boolean => version !== 3 || current === 3;

/** Chooses a Compatibility Version as one undo step; undo sets the previous value back. The held
 *  value records nothing. Returns whether the workbook changed. */
export function chooseCompatibilityVersion(
  wb: WorkbookHandle,
  history: History | null,
  version: CompatibilityVersionCode,
): boolean {
  const previous = wb.compatibilityVersion();
  if (previous === null || previous === version) return false;
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
