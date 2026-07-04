import type {
  FunctionMetadataEntry,
  FunctionMetadataResult,
  MergedFunctionMetadataResult,
} from './types.js';

/**
 * BCP-47 display tag for each engine locale ordinal. Mirrors the ordinal
 * convention used across the function-locale surface (0 = en-US, 1 = ja-JP)
 * and matches the keys a host uses in a {@link FunctionMetadataProvider}'s
 * `aliases` / `localized` maps.
 */
export const LOCALE_TAGS = ['en-US', 'ja-JP'] as const;

/** Translate a locale ordinal to its BCP-47 tag, defaulting to en-US. */
export function localeTag(locale: number): string {
  return LOCALE_TAGS[locale] ?? LOCALE_TAGS[0];
}

/**
 * Merge a host-supplied {@link FunctionMetadataEntry} over the engine's
 * structural `functionMetadata()` result. This is the pure helper the WASM
 * module documents but does not ship at runtime (its generated JS exports no
 * `mergeFunctionMetadata`), reimplemented per `docs/function-metadata-schema.md`.
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
