/**
 * Tag for tests that write or read the system clipboard. Every browser context on a machine shares
 * the OS pasteboard, so a copy in one worker can land between another worker's copy and paste.
 * Tagged tests run in the `<browser>-clipboard` projects, one at a time.
 */
export const SYSTEM_CLIPBOARD = '@system-clipboard';
