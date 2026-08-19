/** Picture, screenshot and drawing-shape icons. */

import { circle, ellipse, poly, rect, roundRect } from '../path.js';
import { arrow, badge, filled, monitor, outlined, stroked } from '../primitives.js';
import { PALETTE } from '../tokens.js';
import type { IconDefinition } from '../types.js';

/** The photo: frame, sun and hills, sized to whatever box it is given. */
const photo = (box: { x: number; y: number; w: number; h: number }, framed = true) => {
  const { x, y, w, h } = box;
  const base = y + h;
  return [
    ...(framed ? [outlined(rect(x, y, w, h), PALETTE.paper, PALETTE.ink, 'regular')] : []),
    filled(circle(x + w * 0.72, y + h * 0.28, Math.min(w, h) * 0.11), PALETTE.warn),
    filled(
      poly(
        [
          [x + w * 0.08, base - h * 0.1],
          [x + w * 0.38, base - h * 0.55],
          [x + w * 0.62, base - h * 0.1],
        ],
        true,
      ),
      PALETTE.accent,
    ),
    filled(
      poly(
        [
          [x + w * 0.45, base - h * 0.1],
          [x + w * 0.68, base - h * 0.42],
          [x + w * 0.92, base - h * 0.1],
        ],
        true,
      ),
      PALETTE.accentDeep,
    ),
  ];
};

const PHOTO_FULL = { x: 2.4, y: 4.4, w: 19.2, h: 15.2 } as const;
const PHOTO_BADGED = { x: 2.4, y: 5.4, w: 16.2, h: 13.2 } as const;

export const MEDIA_ICONS = {
  picture: photo(PHOTO_FULL),
  devicePicture: [...monitor(), ...photo({ x: 4.4, y: 6, w: 15.2, h: 10.4 }, false)],
  onlinePicture: [
    ...photo(PHOTO_BADGED),
    ...badge({ glyph: 'dots', corner: 'br', tone: PALETTE.info }),
  ],
  stockPicture: [
    ...photo(PHOTO_BADGED),
    ...badge({ glyph: 'star', corner: 'br', tone: PALETTE.violet }),
  ],

  screenshot: [
    ...monitor(),
    ...arrow({ from: [7.4, 14], to: [16.6, 7.4], color: PALETTE.accent, head: 'solid' }),
  ],
  screenshotWindow: [
    ...monitor(),
    outlined(rect(5.4, 7, 13.2, 8), PALETTE.mute, PALETTE.grid, 'thin'),
    stroked('M5.4 9.6h13.2', PALETTE.grid, 'thin'),
  ],
  screenClipping: [
    stroked(rect(2.6, 4.6, 18.8, 14.8), PALETTE.ink, 'bold', { dash: '3 2.4', cap: 'butt' }),
    outlined(rect(6.4, 8, 11.2, 8), PALETTE.infoSoft, PALETTE.info, 'regular'),
  ],

  shapes: [
    outlined(rect(2.6, 3.4, 8, 8), PALETTE.paper, PALETTE.info, 'bold'),
    outlined(
      poly(
        [
          [17.4, 2.6],
          [21.8, 11.4],
          [13, 11.4],
        ],
        true,
      ),
      PALETTE.paper,
      PALETTE.accent,
      'bold',
    ),
    outlined(circle(6.6, 17.4, 4.2), PALETTE.paper, PALETTE.warnDeep, 'bold'),
    outlined(roundRect(13, 13.6, 8.8, 7.6, 1.4), PALETTE.paper, PALETTE.alt, 'bold'),
  ],
  shapeLine: [stroked('M3.4 20.6 20.6 3.4', PALETTE.info, 'heavy')],
  shapeArrow: [
    ...arrow({
      from: [3.4, 20.6],
      to: [20.6, 3.4],
      color: PALETTE.info,
      head: 'solid',
      weight: 'heavy',
    }),
  ],
  shapeRectangle: [stroked(rect(2.6, 5.4, 18.8, 13.2), PALETTE.info, 'heavy', { join: 'miter' })],
  shapeRoundedRectangle: [stroked(roundRect(2.6, 5.4, 18.8, 13.2, 3.4), PALETTE.accent, 'heavy')],
  shapeOval: [stroked(ellipse(12, 12, 9.4, 6.8), PALETTE.alt, 'heavy')],
  shapeTriangle: [
    stroked(
      poly(
        [
          [12, 3.2],
          [21.4, 19.8],
          [2.6, 19.8],
        ],
        true,
      ),
      PALETTE.accent,
      'heavy',
    ),
  ],
  shapeDiamond: [
    stroked(
      poly(
        [
          [12, 2.6],
          [21.4, 12],
          [12, 21.4],
          [2.6, 12],
        ],
        true,
      ),
      PALETTE.info,
      'heavy',
    ),
  ],
} satisfies Record<string, IconDefinition>;
