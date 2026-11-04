/** Title-bar quick access and search-panel glyphs. */

import { arc, circle, hLine, type Point, poly, roundRect } from '../path.js';
import { arcArrow, filled, loopArrow, outlined, stroked } from '../primitives.js';
import { PALETTE } from '../tokens.js';
import type { IconDefinition } from '../types.js';

/** A clock whose rim turns counter-clockwise, its gap and head at the left. */
function historyClock(cx: number, cy: number, r: number): IconDefinition {
  const at = (deg: number): Point => {
    const rad = (deg * Math.PI) / 180;
    return [cx + r * Math.cos(rad), cy + r * Math.sin(rad)];
  };
  const end = 205;
  const rad = (end * Math.PI) / 180;
  return [
    ...arcArrow({
      from: at(150),
      to: at(end),
      radius: r,
      sweep: 0,
      largeArc: 1,
      headDir: [Math.sin(rad), -Math.cos(rad)],
      color: PALETTE.ink,
      weight: 'regular',
      head: 'solid',
    }),
    stroked(
      poly([
        [cx, cy - 4.4],
        [cx, cy],
        [cx + 3, cy + 1.8],
      ]),
      PALETTE.ink,
      'regular',
    ),
  ];
}

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
  // The title-bar comments toggle: an outline speech bubble with its tail low left.
  comments: [
    outlined(
      poly(
        [
          [3.4, 4.6],
          [20.6, 4.6],
          [20.6, 16.2],
          [11.2, 16.2],
          [7.2, 19.8],
          [7.2, 16.2],
          [3.4, 16.2],
        ],
        true,
      ),
      PALETTE.paper,
      PALETTE.ink,
      'regular',
    ),
  ],
  // A file location: a tabbed folder.
  folder: [
    outlined(
      poly(
        [
          [2.8, 5.8],
          [9.2, 5.8],
          [11, 7.8],
          [21.2, 7.8],
          [21.2, 18.8],
          [2.8, 18.8],
        ],
        true,
      ),
      PALETTE.infoMid,
      PALETTE.info,
      'regular',
    ),
    stroked(hLine(2.8, 10.2, 18.4), PALETTE.info, 'thin'),
  ],
  // Earlier versions: a clock face wound back by a counter-clockwise arrow.
  versionHistory: [...historyClock(12.2, 12, 8.6)],
} satisfies Record<string, IconDefinition>;
