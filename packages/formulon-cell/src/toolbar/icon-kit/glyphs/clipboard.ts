/** Clipboard, history and format-painter icons. */

import { hLine, join, line, poly, rect, roundRect } from '../path.js';
import {
  badge,
  clipboard,
  filled,
  headed,
  outlined,
  place,
  stroked,
  textLines,
} from '../primitives.js';
import { PALETTE } from '../tokens.js';
import type { IconDefinition } from '../types.js';

/** Content area inside the clipboard board. */
const BOARD = { x: 6.4, y: 8.4, w: 11.2, h: 11.2 } as const;

/** A mark centred on the clipboard's content area. */
const onBoard = (
  segments: readonly ReturnType<typeof stroked>[],
  box: readonly [number, number, number, number],
) => place(segments, { box, size: 9.6, cx: BOARD.x + BOARD.w / 2, cy: BOARD.y + BOARD.h / 2 });

/** Lowercase f with a multiplication cross — the formula mark. */
const formulaMark = [
  stroked(join('M12.9 6.4c-1.9-.6-3 .3-3.4 2.3l-2 9.6', 'M7.9 10.4h4.4'), PALETTE.violet, 'bold'),
  stroked(
    join(line([13.1, 13.4], [16.3, 17.2]), line([16.3, 13.4], [13.1, 17.2])),
    PALETTE.ink,
    'regular',
  ),
];
const FORMULA_BOX = [6.9, 6.4, 9.4, 10.8] as const;

/** Digits 1-2-3, the values mark. */
const valuesMark = [
  stroked(
    join(
      'M4.4 5.6 5.9 4.4v8.2',
      'M8.6 5.8a1.7 1.7 0 0 1 3.2.8c0 1.9-3.2 3-3.2 5.7h3.4',
      'M14 4.4h2.8l-1.7 2.8a2 2 0 1 1-1.6 3',
    ),
    PALETTE.accent,
    'bold',
  ),
];
const VALUES_BOX = [3.5, 3.5, 14.2, 9.7] as const;

export const CLIPBOARD_ICONS = {
  // The step-back arrows read as "go back", not "rotate": a straight run into a
  // half-turn, with the head square on the shaft's own axis.
  undo: [
    ...headed({
      shaft: 'M3.6 6.2h9.8a6.2 6.2 0 0 1 0 12.4H7.6',
      tip: [3.6, 6.2],
      headDir: [-1, 0],
      color: PALETTE.ink,
      headScale: 1.2,
    }),
  ],
  redo: [
    ...headed({
      shaft: 'M20.4 6.2h-9.8a6.2 6.2 0 0 0 0 12.4h5.8',
      tip: [20.4, 6.2],
      headDir: [1, 0],
      color: PALETTE.ink,
      headScale: 1.2,
    }),
  ],

  paste: [
    ...clipboard(),
    ...textLines({ x: BOARD.x + 1, y: BOARD.y + 1.4, width: BOARD.w - 2, count: 3, gap: 3.2 }),
  ],
  pasteFormulas: [...clipboard(), ...onBoard(formulaMark, FORMULA_BOX)],
  pasteValues: [...clipboard(), ...onBoard(valuesMark, VALUES_BOX)],
  // Rows become columns: one run turns the corner and comes back down.
  pasteTranspose: [
    ...clipboard(),
    ...headed({
      shaft: 'M7.4 11.4h8.4v6.4',
      tip: [15.8, 18.4],
      headDir: [0, 1],
      color: PALETTE.info,
      weight: 'regular',
    }),
    stroked('M7.4 15.6h4.2', PALETTE.gridLight, 'thin'),
  ],
  pasteSpecial: [
    ...clipboard(),
    ...textLines({ x: BOARD.x + 1, y: BOARD.y + 1.2, width: BOARD.w - 2, count: 2, gap: 2.8 }),
    ...badge({ glyph: 'dots', corner: 'br', tone: PALETTE.info }),
  ],

  cut: [
    stroked(
      join(line([5.4, 3.6], [13.4, 14.4]), line([18.6, 3.6], [10.6, 14.4])),
      PALETTE.ink,
      'bold',
    ),
    outlined(roundRect(3.2, 15, 6, 6, 3), PALETTE.paper, PALETTE.info, 'bold'),
    outlined(roundRect(14.8, 15, 6, 6, 3), PALETTE.paper, PALETTE.info, 'bold'),
    outlined(roundRect(10.4, 11.2, 3.2, 3.2, 1.6), PALETTE.paper, PALETTE.ink, 'thin'),
  ],
  copy: [
    outlined(roundRect(2.5, 2.5, 13, 15, 1.4), PALETTE.mute, PALETTE.grid, 'thin'),
    outlined(roundRect(8.5, 6.5, 13, 15, 1.4), PALETTE.paper, PALETTE.ink, 'regular'),
    ...textLines({ x: 11, y: 11, width: 8, count: 3, gap: 3 }),
  ],
  paint: [
    outlined(rect(9.4, 2, 5.2, 6.6), PALETTE.tan, PALETTE.tanDeep, 'thin'),
    outlined(rect(8, 8.6, 8, 2.8), PALETTE.mute, PALETTE.grid, 'thin'),
    outlined(roundRect(8.6, 11.4, 6.8, 5.2, 1), PALETTE.info, PALETTE.ink, 'thin'),
    stroked('M12 16.6v3.2', PALETTE.ink, 'regular'),
    stroked('M6 21.4c4-1.8 8-1.8 12 0', PALETTE.accent, 'bold'),
  ],

  pen: [
    filled(
      poly(
        [
          [3, 21],
          [4.6, 16.4],
          [16.6, 4.4],
          [19.6, 7.4],
          [7.6, 19.4],
        ],
        true,
      ),
      PALETTE.warn,
    ),
    stroked(
      poly([
        [3, 21],
        [4.6, 16.4],
        [16.6, 4.4],
        [19.6, 7.4],
        [7.6, 19.4],
        [3, 21],
      ]),
      PALETTE.ink,
      'thin',
    ),
    filled(
      poly(
        [
          [16.6, 4.4],
          [18.4, 2.6],
          [21.4, 5.6],
          [19.6, 7.4],
        ],
        true,
      ),
      PALETTE.info,
    ),
    stroked(line([4.6, 16.4], [7.6, 19.4]), PALETTE.ink, 'hairline'),
  ],
  eraser: [
    outlined(
      poly(
        [
          [8.2, 3.4],
          [20.6, 9.8],
          [15.4, 17.4],
          [3, 11],
        ],
        true,
      ),
      PALETTE.paper,
      PALETTE.ink,
      'regular',
    ),
    stroked(line([14.4, 6.6], [9.2, 14.2]), PALETTE.grid, 'thin'),
    stroked(hLine(3, 21, 18), PALETTE.info, 'bold'),
  ],
} satisfies Record<string, IconDefinition>;
