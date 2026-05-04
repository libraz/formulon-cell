// Locale entry — sub-path export (`@libraz/formulon-cell/i18n/ja`).
//
// Splitting locales out per file lets consumers ship only the dictionaries
// they actually use. The full strings registry still lives in `strings.ts`
// so we don't duplicate keys.
export { ja as default, ja } from './strings.js';
export type { Strings } from './strings.js';
