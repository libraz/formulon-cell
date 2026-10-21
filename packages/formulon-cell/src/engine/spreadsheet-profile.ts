import { type BuiltInLocale, builtInLocaleFor } from '../i18n/locales.js';
import type { ExcelProfileId, SpreadsheetProfileId } from './types.js';

/** Raw ids are strings at the compatibility boundary so older engines can be
 * probed and unknown ids can be reported as `null` instead of rejected by
 * the adapter's type system. */
export type EngineSpreadsheetProfileId = string;

const PROFILE_VERSION = '365' as const;

const PUBLIC_TO_ENGINE: Readonly<Record<SpreadsheetProfileId, ExcelProfileId>> = {
  'windows-ja_JP': `win-${PROFILE_VERSION}-ja_JP`,
  'windows-en_US': `win-${PROFILE_VERSION}-en_US`,
  'windows-de_DE': `win-${PROFILE_VERSION}-de_DE`,
  'windows-fr_FR': `win-${PROFILE_VERSION}-fr_FR`,
  'windows-zh_CN': `win-${PROFILE_VERSION}-zh_CN`,
  'windows-ko_KR': `win-${PROFILE_VERSION}-ko_KR`,
  'windows-th_TH': `win-${PROFILE_VERSION}-th_TH`,
  'mac-ja_JP': `mac-${PROFILE_VERSION}-ja_JP`,
  'mac-en_US': `mac-${PROFILE_VERSION}-en_US`,
  'mac-de_DE': `mac-${PROFILE_VERSION}-de_DE`,
  'mac-fr_FR': `mac-${PROFILE_VERSION}-fr_FR`,
  'mac-zh_CN': `mac-${PROFILE_VERSION}-zh_CN`,
  'mac-ko_KR': `mac-${PROFILE_VERSION}-ko_KR`,
  'mac-th_TH': `mac-${PROFILE_VERSION}-th_TH`,
};

const ENGINE_TO_PUBLIC: ReadonlyMap<EngineSpreadsheetProfileId, SpreadsheetProfileId> = new Map(
  Object.entries(PUBLIC_TO_ENGINE).map(([publicId, engineId]) => [
    engineId,
    publicId as SpreadsheetProfileId,
  ]),
);

export const engineProfileToPublic = (
  profileId: EngineSpreadsheetProfileId,
): SpreadsheetProfileId | null => {
  return ENGINE_TO_PUBLIC.get(profileId) ?? null;
};

export const publicProfileToEngine = (profileId: SpreadsheetProfileId): ExcelProfileId =>
  PUBLIC_TO_ENGINE[profileId];

/** Desktop platform whose defaults and screen geometry a workbook follows. */
export type SpreadsheetPlatform = 'mac' | 'default';

/** The locale suffixes that have a corresponding formulon profile. */
export type ProfileLocale = 'ja_JP' | 'en_US' | 'de_DE' | 'fr_FR' | 'zh_CN' | 'ko_KR' | 'th_TH';

export const PROFILE_SUFFIX_FOR_LOCALE: Readonly<Record<BuiltInLocale, ProfileLocale>> = {
  ja: 'ja_JP',
  en: 'en_US',
  de: 'de_DE',
  fr: 'fr_FR',
  zh: 'zh_CN',
  ko: 'ko_KR',
  th: 'th_TH',
};

/** Engine profile id for `locale` on `platform`. A tag outside the built-in
 *  languages resolves by its language subtag, else to en_US, so every locale
 *  has an id for new-workbook defaults. */
export const engineProfileIdFor = (
  locale: string,
  platform: SpreadsheetPlatform,
): ExcelProfileId => {
  const language = builtInLocaleFor(locale);
  const suffix = language ? PROFILE_SUFFIX_FOR_LOCALE[language] : 'en_US';
  const prefix = platform === 'mac' ? 'mac' : 'windows';
  return publicProfileToEngine(`${prefix}-${suffix}`);
};
