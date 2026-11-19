/** Escape every RegExp metacharacter so `text` matches literally. */
export const escapeRegExp = (text: string): string => text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

/** Escape `text` for HTML element content and quoted attribute values. */
export const escapeHtml = (text: string): string =>
  text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');

/** Escape a dynamic value for a CSS string. `<` and `>` are escaped even
 * inside quotes because the complete stylesheet is an HTML raw-text node. */
export const cssString = (value: string): string =>
  `"${value
    .replaceAll('\\', '\\\\')
    .replaceAll('"', '\\"')
    .replace(/\r\n?|\n/g, '\\A ')
    .replaceAll('&', '\\26 ')
    .replaceAll('<', '\\3C ')
    .replaceAll('>', '\\3E ')}"`;
