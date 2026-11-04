/**
 * The delimiters a plain-text paste splits columns on. Like the desktop, a
 * paste splits on exactly the delimiter set of the last delimited Text to
 * Columns run in this session (tab until there has been one); the setting
 * belongs to the session, not to a workbook, so it outlives the one it was
 * set in.
 */

const DEFAULT: readonly string[] = ['\t'];

let current: readonly string[] = DEFAULT;

/** The delimiters a plain-text paste splits on now. */
export const pasteDelimiters = (): readonly string[] => current;

/** Record the delimiters of a delimited Text to Columns run. */
export function rememberTextToColumnsDelimiters(delimiters: string | readonly string[]): void {
  const list = (typeof delimiters === 'string' ? [delimiters] : delimiters).filter((d) => d !== '');
  current = list.length > 0 ? [...new Set(list)] : DEFAULT;
}

/** Back to the session default (tab). */
export function resetPasteDelimiters(): void {
  current = DEFAULT;
}
