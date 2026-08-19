/** Save, export, share and extension icons. */

import { circle, join, line, poly, rect, roundRect } from '../path.js';
import { arrow, badge, bars, filled, outlined, stroked } from '../primitives.js';
import { PALETTE } from '../tokens.js';
import type { IconDefinition } from '../types.js';

/** The save body: a diskette with its shutter and label. */
const diskette = (box: { x: number; y: number; w: number; h: number }) => {
  const notch = box.w * 0.22;
  return [
    outlined(
      poly(
        [
          [box.x, box.y],
          [box.x + box.w - notch, box.y],
          [box.x + box.w, box.y + notch],
          [box.x + box.w, box.y + box.h],
          [box.x, box.y + box.h],
        ],
        true,
      ),
      PALETTE.paper,
      PALETTE.ink,
      'regular',
    ),
    filled(rect(box.x + box.w * 0.24, box.y, box.w * 0.44, box.h * 0.34), PALETTE.info),
    outlined(
      rect(box.x + box.w * 0.18, box.y + box.h * 0.52, box.w * 0.64, box.h * 0.48),
      PALETTE.mute,
      PALETTE.grid,
      'thin',
    ),
  ];
};

export const FILE_ICONS = {
  save: diskette({ x: 2.6, y: 3.4, w: 18.8, h: 17.2 }),
  saveAs: [
    ...diskette({ x: 2.6, y: 4.4, w: 16, h: 15.2 }),
    ...badge({ glyph: 'pencil', corner: 'br', tone: PALETTE.alt }),
  ],
  autosave: [
    outlined(roundRect(2.4, 7.4, 19.2, 9.2, 4.6), PALETTE.accent, PALETTE.accentDeep, 'regular'),
    filled(circle(17, 12, 3.2), PALETTE.paper),
  ],

  // Share: content leaving the sheet, upward and out.
  share: [
    stroked('M4.4 13.4v7.2h15.2v-7.2', PALETTE.ink, 'regular'),
    ...arrow({ from: [12, 16.4], to: [12, 3.4], color: PALETTE.accent, head: 'solid' }),
  ],

  pdf: [
    outlined(
      poly(
        [
          [4.4, 2.4],
          [15.2, 2.4],
          [19.6, 6.8],
          [19.6, 21.6],
          [4.4, 21.6],
        ],
        true,
      ),
      PALETTE.paper,
      PALETTE.ink,
      'regular',
    ),
    stroked(
      poly([
        [15.2, 2.4],
        [15.2, 6.8],
        [19.6, 6.8],
      ]),
      PALETTE.ink,
      'thin',
      { join: 'miter' },
    ),
    filled(rect(6.4, 11.4, 11.2, 6.4), PALETTE.danger),
    ...bars({
      x: 8.2,
      y: 14.6,
      width: 7.6,
      widths: [1],
      gap: 0,
      color: PALETTE.paper,
      weight: 'regular',
    }),
  ],

  script: [
    outlined(rect(2.4, 4.4, 19.2, 15.2), PALETTE.paper, PALETTE.ink, 'regular'),
    stroked(
      join(
        poly([
          [9, 9.4],
          [6.4, 12],
          [9, 14.6],
        ]),
        poly([
          [15, 9.4],
          [17.6, 12],
          [15, 14.6],
        ]),
      ),
      PALETTE.info,
      'bold',
    ),
    stroked(line([13.2, 8.4], [10.8, 15.6]), PALETTE.accent, 'bold'),
  ],
  addIn: [
    outlined(rect(2.6, 2.6, 8.4, 8.4), PALETTE.paper, PALETTE.ink, 'regular'),
    outlined(rect(13, 2.6, 8.4, 8.4), PALETTE.paper, PALETTE.ink, 'regular'),
    outlined(rect(2.6, 13, 8.4, 8.4), PALETTE.paper, PALETTE.ink, 'regular'),
    filled(roundRect(13, 13, 8.4, 8.4, 1.4), PALETTE.accent),
    stroked(join('M15.2 17.2h4', 'M17.2 15.2v4'), PALETTE.paper, 'heavy'),
  ],
} satisfies Record<string, IconDefinition>;
