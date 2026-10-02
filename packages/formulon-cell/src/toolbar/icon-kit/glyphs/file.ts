/** Save, export, share and extension icons. */

import { circle, join, line, poly, rect, roundRect } from '../path.js';
import { arrow, badge, bars, doc, filled, outlined, stroked } from '../primitives.js';
import { PALETTE } from '../tokens.js';
import type { IconDefinition } from '../types.js';

/** The save body: a diskette with its shutter and label. */
const diskette = (box: { x: number; y: number; w: number; h: number }) => {
  const notch = box.w * 0.22;
  const body = poly(
    [
      [box.x, box.y],
      [box.x + box.w - notch, box.y],
      [box.x + box.w, box.y + notch],
      [box.x + box.w, box.y + box.h],
      [box.x, box.y + box.h],
    ],
    true,
  );
  const shutter = {
    x: box.x + box.w * 0.25,
    y: box.y + box.h * 0.08,
    w: box.w * 0.42,
    h: box.h * 0.28,
  };
  const label = {
    x: box.x + box.w * 0.18,
    y: box.y + box.h * 0.51,
    w: box.w * 0.64,
    h: box.h * 0.35,
  };
  const slot = roundRect(
    shutter.x + shutter.w * 0.18,
    shutter.y + shutter.h * 0.6,
    shutter.w * 0.64,
    0.7,
    0.35,
  );

  return [
    filled(body, PALETTE.paper),
    filled(roundRect(shutter.x, shutter.y, shutter.w, shutter.h, 0.55), PALETTE.info),
    stroked(
      roundRect(shutter.x, shutter.y, shutter.w, shutter.h, 0.55),
      PALETTE.infoDeep,
      'hairline',
      { cap: 'butt' },
    ),
    filled(slot, PALETTE.infoDeep),
    filled(roundRect(label.x, label.y, label.w, label.h, 0.7), PALETTE.mute),
    stroked(roundRect(label.x, label.y, label.w, label.h, 0.7), PALETTE.grid, 'hairline', {
      cap: 'butt',
    }),
    ...bars({
      x: label.x + label.w * 0.16,
      y: label.y + label.h * 0.3,
      width: label.w * 0.68,
      widths: [0.9, 0.62],
      gap: label.h * 0.28,
      color: PALETTE.grid,
      weight: 'hairline',
    }),
    // Paint the outer edge after every inset detail so the silhouette stays crisp.
    stroked(body, PALETTE.ink, 'regular', { cap: 'butt', join: 'miter' }),
  ];
};

export const FILE_ICONS = {
  manage: [
    ...doc({ x: 3, y: 2.4, w: 15, h: 18, fold: 3, lines: 2 }),
    ...badge({ glyph: 'plus', corner: 'br', shape: 'circle' }),
  ],

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
