/** Title-bar quick access and search-panel glyphs. */

import { arc, circle, poly, roundRect } from '../path.js';
import { filled, loopArrow, outlined, stroked } from '../primitives.js';
import { PALETTE } from '../tokens.js';
import type { IconDefinition } from '../types.js';

export const SHELL_ICONS = {
  // A gabled house with its door.
  home: [
    outlined(
      poly(
        [
          [3.4, 11.4],
          [12, 3.8],
          [20.6, 11.4],
          [18.6, 11.4],
          [18.6, 20.2],
          [5.4, 20.2],
          [5.4, 11.4],
        ],
        true,
      ),
      PALETTE.paper,
      PALETTE.ink,
      'regular',
    ),
    filled(roundRect(10, 13.6, 4, 6.6, 0.8), PALETTE.infoSoft),
    stroked(
      poly([
        [10, 20.2],
        [10, 14.4],
        [14, 14.4],
        [14, 20.2],
      ]),
      PALETTE.info,
      'thin',
    ),
  ],
  // Do the last action again: a clockwise loop closing on its own head.
  repeat: loopArrow({ cx: 12, cy: 12, r: 7.9, color: PALETTE.ink, weight: 'bold' }),
  // The customise-toolbar menu: three dots on a line.
  moreCommands: [
    filled(circle(4.4, 12, 1.8), PALETTE.ink),
    filled(circle(12, 12, 1.8), PALETTE.ink),
    filled(circle(19.6, 12, 1.8), PALETTE.ink),
  ],
  // A person speaking: head and shoulders with a bubble at the upper right.
  feedback: [
    stroked(circle(8.4, 12.4, 3.2), PALETTE.ink, 'regular'),
    stroked(arc(8.4, 22.4, 6.4, 6.4, 200, 340), PALETTE.ink, 'regular'),
    outlined(
      poly(
        [
          [14.2, 3.4],
          [21.4, 3.4],
          [21.4, 9.6],
          [17.2, 9.6],
          [15, 11.8],
          [15, 9.6],
          [14.2, 9.6],
        ],
        true,
      ),
      PALETTE.infoSoft,
      PALETTE.info,
      'regular',
    ),
  ],
} satisfies Record<string, IconDefinition>;
