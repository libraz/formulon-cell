import type {
  FunctionMetadataEntry,
  FunctionMetadataResult,
  MergedFunctionMetadataResult,
} from './types.js';

/**
 * BCP-47 display tag for each engine locale ordinal. Mirrors the ordinal
 * convention used across the function-locale surface (0 = en-US, 1 = ja-JP,
 * 2 = de-DE, 3 = fr-FR)
 * and matches the keys a host uses in a {@link FunctionMetadataProvider}'s
 * `aliases` / `localized` maps.
 */
export const LOCALE_TAGS = ['en-US', 'ja-JP', 'de-DE', 'fr-FR'] as const;

/** UI locale tags for the seven built-in dictionaries. These are deliberately
 * separate from {@link LOCALE_TAGS}: the native function-name API still uses
 * the historical four-value ordinal and the Chinese, Korean, and Thai
 * profiles keep canonical English function names. */
export const UI_LOCALE_TAGS = [
  'en-US',
  'ja-JP',
  'de-DE',
  'fr-FR',
  'zh-CN',
  'ko-KR',
  'th-TH',
] as const;

export type UiLocaleTag = (typeof UI_LOCALE_TAGS)[number];

/** Normalize a host locale into the canonical tag used by metadata providers. */
export function displayLocaleTag(locale: string | null | undefined): string {
  const language = (locale ?? '').trim().toLowerCase().replace(/_/g, '-').split('-')[0];
  switch (language) {
    case 'ja':
      return 'ja-JP';
    case 'de':
      return 'de-DE';
    case 'fr':
      return 'fr-FR';
    case 'zh':
      return 'zh-CN';
    case 'ko':
      return 'ko-KR';
    case 'th':
      return 'th-TH';
    case 'en':
      return 'en-US';
    default:
      return locale?.trim() || 'en-US';
  }
}

/** Legacy native-function ordinal for a provider tag, or `null` when the
 * tag intentionally has no localized native aliases. */
export function localeOrdinalForTag(tag: string): 0 | 1 | 2 | 3 | null {
  switch (displayLocaleTag(tag)) {
    case 'ja-JP':
      return 1;
    case 'de-DE':
      return 2;
    case 'fr-FR':
      return 3;
    case 'en-US':
      return 0;
    default:
      return null;
  }
}

/** Translate a locale ordinal to its BCP-47 tag, defaulting to en-US. */
export function localeTag(locale: number): string {
  return LOCALE_TAGS[locale] ?? LOCALE_TAGS[0];
}

/**
 * Merge a host-supplied {@link FunctionMetadataEntry} over the engine's
 * structural `functionMetadata()` result.
 *
 * The engine package exports the same pure helper, but only from the entry
 * point that also pulls in the generated WASM glue. Keeping a copy here lets a
 * host merge metadata without dragging the engine module into its bundle; the
 * two agree on `docs/function-metadata-schema.md`.
 *
 * Field precedence (first non-nullish wins):
 *   - `signatureTemplate`: `entry.localized[locale].signature` →
 *     `entry.signature` → `base.signatureTemplate`
 *   - `description`: `entry.localized[locale].description` →
 *     `entry.description` → `base.description`
 *   - `localizedName`: `entry.aliases[locale]` → `base.name`
 *
 * When `entry` is `undefined`, `base` is returned verbatim (no `localizedName`
 * is attached). `locale` is a BCP-47 display tag matching the keys in
 * `aliases` / `localized`.
 */
export function mergeFunctionMetadata(
  base: FunctionMetadataResult,
  entry: FunctionMetadataEntry | undefined,
  locale: string,
): MergedFunctionMetadataResult {
  if (!entry) return base;
  const loc = entry.localized?.[locale];
  const signatureTemplate = loc?.signature ?? entry.signature ?? base.signatureTemplate;
  const description = loc?.description ?? entry.description ?? base.description;
  const localizedName = entry.aliases?.[locale] ?? base.name;
  return {
    ...base,
    ...(signatureTemplate !== undefined ? { signatureTemplate } : {}),
    ...(description !== undefined ? { description } : {}),
    ...(localizedName !== undefined ? { localizedName } : {}),
  };
}
