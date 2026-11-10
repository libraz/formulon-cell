import type { RibbonToolbarConfig } from '../api/spreadsheet/types.js';

export type { RibbonToolbarConfig } from '../api/spreadsheet/types.js';

/** The ribbon toolbar, for `mount(host, { toolbar: ribbon() })` and `update({ toolbar })`. */
export const ribbon = (): RibbonToolbarConfig => ({ kind: 'ribbon' });
