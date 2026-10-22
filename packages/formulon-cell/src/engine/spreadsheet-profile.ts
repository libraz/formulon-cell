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
  'windows-zh_TW': `win-${PROFILE_VERSION}-zh_TW`,
  'windows-ko_KR': `win-${PROFILE_VERSION}-ko_KR`,
  'windows-th_TH': `win-${PROFILE_VERSION}-th_TH`,
  'mac-ja_JP': `mac-${PROFILE_VERSION}-ja_JP`,
  'mac-en_US': `mac-${PROFILE_VERSION}-en_US`,
  'mac-de_DE': `mac-${PROFILE_VERSION}-de_DE`,
  'mac-fr_FR': `mac-${PROFILE_VERSION}-fr_FR`,
  'mac-zh_CN': `mac-${PROFILE_VERSION}-zh_CN`,
  'mac-zh_TW': `mac-${PROFILE_VERSION}-zh_TW`,
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
export type ProfileLocale =
  | 'ja_JP'
  | 'en_US'
  | 'de_DE'
  | 'fr_FR'
  | 'zh_CN'
  | 'zh_TW'
  | 'ko_KR'
  | 'th_TH';

export const PROFILE_SUFFIX_FOR_LOCALE: Readonly<Record<BuiltInLocale, ProfileLocale>> = {
  ja: 'ja_JP',
  en: 'en_US',
  de: 'de_DE',
  fr: 'fr_FR',
  zh: 'zh_CN',
  ko: 'ko_KR',
  th: 'th_TH',
};

const TRADITIONAL_REGIONS: ReadonlySet<string> = new Set(['TW', 'HK', 'MO']);

/** True when a `zh` tag denotes Traditional Chinese: script `Hant`, or region
 *  TW / HK / MO without an explicit `Hans` script. Unparseable tags are not. */
const isTraditionalChinese = (locale: string): boolean => {
  try {
    const { script, region } = new Intl.Locale(locale.trim().replaceAll('_', '-'));
    if (script) return script === 'Hant';
    return region !== undefined && TRADITIONAL_REGIONS.has(region);
  } catch {
    return false;
  }
};

/** Profile locale suffix for `locale`: the built-in language's suffix, with
 *  Traditional Chinese tags resolving to zh_TW; else en_US. */
export const profileLocaleFor = (locale: string): ProfileLocale => {
  const language = builtInLocaleFor(locale);
  if (!language) return 'en_US';
  if (language === 'zh' && isTraditionalChinese(locale)) return 'zh_TW';
  return PROFILE_SUFFIX_FOR_LOCALE[language];
};

/** Engine profile id for `locale` on `platform`. A tag outside the built-in
 *  languages resolves by its language subtag, else to en_US, so every locale
 *  has an id for new-workbook defaults. */
export const engineProfileIdFor = (
  locale: string,
  platform: SpreadsheetPlatform,
): ExcelProfileId => {
  const suffix = profileLocaleFor(locale);
  const prefix = platform === 'mac' ? 'mac' : 'windows';
  return publicProfileToEngine(`${prefix}-${suffix}`);
};
