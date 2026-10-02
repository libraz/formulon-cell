import { ALIGN_ICONS } from './align.js';
import { CHART_ICONS } from './chart.js';
import { CLEAR_ICONS } from './clear.js';
import { CLIPBOARD_ICONS } from './clipboard.js';
import { FILE_ICONS } from './file.js';
import { FILL_ICONS } from './fill.js';
import { FIND_ICONS } from './find.js';
import { FORMAT_ICONS } from './format.js';
import { FORMULA_ICONS } from './formulas.js';
import { MAC_ICONS } from './mac.js';
import { MEDIA_ICONS } from './media.js';
import { REVIEW_ICONS } from './review.js';
import { SORT_FILTER_ICONS } from './sort-filter.js';
import { STRUCTURE_ICONS } from './structure.js';
import { TABLE_ICONS } from './table.js';
import { VIEW_ICONS } from './view.js';

/** Every ribbon glyph, composed from the icon kit. */
export const RIBBON_GLYPHS = {
  ...CLIPBOARD_ICONS,
  ...FORMAT_ICONS,
  ...ALIGN_ICONS,
  ...FILL_ICONS,
  ...CLEAR_ICONS,
  ...SORT_FILTER_ICONS,
  ...FIND_ICONS,
  ...STRUCTURE_ICONS,
  ...TABLE_ICONS,
  ...FORMULA_ICONS,
  ...CHART_ICONS,
  ...MEDIA_ICONS,
  ...MAC_ICONS,
  ...REVIEW_ICONS,
  ...VIEW_ICONS,
  ...FILE_ICONS,
};
