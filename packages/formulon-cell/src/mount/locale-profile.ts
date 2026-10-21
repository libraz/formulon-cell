import {
  engineProfileIdFor,
  engineProfileToPublic,
  PROFILE_SUFFIX_FOR_LOCALE,
  type ProfileLocale,
  type SpreadsheetPlatform,
} from '../engine/spreadsheet-profile.js';
import type { SpreadsheetProfileId } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { builtInLocaleFor } from '../i18n/locales.js';

export type LocaleProfilePlatform = SpreadsheetPlatform;

/**
 * Resolve a UI locale to the locale suffix used by Excel profile ids.
 * Unknown custom dictionaries deliberately return `null`: a host supplied
 * dictionary must not silently change formula semantics just because its
 * language tag happens to be unfamiliar to this adapter.
 */
export function profileLocaleForUiLocale(locale: string): ProfileLocale | null {
  const mapped = builtInLocaleFor(locale);
  if (!mapped) return null;

  const normalized = locale.trim().replaceAll('_', '-');
  if (normalized.toLowerCase() === mapped) return PROFILE_SUFFIX_FOR_LOCALE[mapped];

  // Keep this validation aligned with the locale switcher: a built-in
  // language may carry a region, but script and variant tags remain custom.
  try {
    const parsed = new Intl.Locale(normalized);
    const region = parsed.region;
    if (!region || parsed.script) return null;
    if (parsed.baseName.toLowerCase() !== `${mapped}-${region.toLowerCase()}`) return null;
    return PROFILE_SUFFIX_FOR_LOCALE[mapped];
  } catch {
    return null;
  }
}

/** Return the public profile selected by a UI locale and platform. */
export function spreadsheetProfileForLocale(
  locale: string,
  platform: LocaleProfilePlatform,
): SpreadsheetProfileId | null {
  if (!profileLocaleForUiLocale(locale)) return null;
  return engineProfileToPublic(engineProfileIdFor(locale, platform));
}

export interface SyncLocaleProfileOptions {
  /**
   * Throw when a profile-capable engine cannot commit the selected profile.
   * Callers that change UI state can use this to roll back their own state;
   * unsupported engines remain an intentional no-op regardless of this flag.
   */
  throwOnFailure?: boolean;
}

const profileFailure = (profile: SpreadsheetProfileId, message: string): Error =>
  new Error(`setSpreadsheetProfileId(${profile}) failed: ${message}`);

type ProfileWriteResult = { ok: boolean; message?: string } | boolean;

const profileWriteOk = (result: ProfileWriteResult): boolean =>
  typeof result === 'boolean' ? result : result.ok;

const profileWriteMessage = (result: ProfileWriteResult): string =>
  typeof result === 'boolean'
    ? 'engine rejected the profile'
    : result.message || 'engine rejected the profile';

const errorMessage = (error: unknown): string =>
  error instanceof Error ? error.message : String(error);

/**
 * Apply a locale profile to a workbook and recalculate exactly once when the
 * effective profile changes. Returns `true` only after the new profile was
 * read back successfully and the recalc completed.
 *
 * Profile selection is an optional capability. A stub or older package is a
 * supported no-op, while a profile-capable engine that rejects a known id can
 * be made strict by passing `throwOnFailure`.
 */
export function syncLocaleProfile(
  wb: WorkbookHandle,
  locale: string,
  platform: LocaleProfilePlatform,
  options: SyncLocaleProfileOptions = {},
): boolean {
  const target = spreadsheetProfileForLocale(locale, platform);
  if (!target || !wb.capabilities.spreadsheetProfile) return false;

  const current = wb.spreadsheetProfileId();
  if (current === target) return false;

  // A profile write is a transaction from the mount's point of view. The
  // engine status API normally leaves the old profile untouched on failure,
  // but restore defensively when a package reports success and then gives an
  // unexpected readback or recalc error.
  const restoreCurrentProfile = (): string | null => {
    if (current === null) return 'previous profile is unavailable for restoration';
    try {
      if (wb.spreadsheetProfileId() !== current) {
        const restoreResult = wb.setSpreadsheetProfileId(current) as ProfileWriteResult;
        if (!profileWriteOk(restoreResult)) return profileWriteMessage(restoreResult);
        if (wb.spreadsheetProfileId() !== current) {
          return `restoration readback was ${wb.spreadsheetProfileId() ?? 'unavailable'}`;
        }
        wb.recalc();
      }
      return null;
    } catch (error) {
      return errorMessage(error);
    }
  };

  const reject = (message: string): boolean => {
    const restoration = restoreCurrentProfile();
    const detail = restoration ? `${message}; restoration failed: ${restoration}` : message;
    if (options.throwOnFailure) throw profileFailure(target, detail);
    return false;
  };

  let result: ProfileWriteResult;
  try {
    result = wb.setSpreadsheetProfileId(target) as ProfileWriteResult;
  } catch (error) {
    return reject(errorMessage(error));
  }
  if (!profileWriteOk(result)) return reject(profileWriteMessage(result));

  let effective: SpreadsheetProfileId | null;
  try {
    effective = wb.spreadsheetProfileId();
  } catch (error) {
    return reject(`engine profile readback failed: ${errorMessage(error)}`);
  }
  if (effective !== target) {
    return reject(`engine readback was ${effective === null ? 'unavailable' : effective}`);
  }

  try {
    wb.recalc();
  } catch (error) {
    return reject(`recalc failed: ${errorMessage(error)}`);
  }
  return current !== effective;
}
