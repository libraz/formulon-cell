/**
 * Normalize a short locale identifier (e.g. `'ja'`, `'en'`) to the BCP-47 tag
 * used by `Intl.NumberFormat` / `Intl.DateTimeFormat`. Empty / unknown values
 * fall back to `en-US` so callers never see a runtime locale error.
 *
 * Kept as a tiny shared utility so the renderer (`render/grid/hit-state.ts`)
 * and the format dialog (`interact/format-dialog-model.ts`) agree on the
 * same canonicalisation — earlier they each carried their own copy.
 */
export const normalizeFormatLocale = (locale: string): string => {
  if (locale === 'ja') return 'ja-JP';
  if (locale === 'en') return 'en-US';
  return locale || 'en-US';
};
