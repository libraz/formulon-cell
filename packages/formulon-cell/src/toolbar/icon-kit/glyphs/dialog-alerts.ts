/** Alert marks for the Data Validation error styles: stop, warning, information. */

import { circle, join, line, poly } from '../path.js';
import { filled, stroked } from '../primitives.js';
import { CENTER, OPTICAL, PALETTE } from '../tokens.js';
import type { IconDefinition } from '../types.js';

const R = OPTICAL / 2;
const ARM = R * 0.42;

/** A red disc with a white cross. */
const stop: IconDefinition = [
  filled(circle(CENTER, CENTER, R), PALETTE.danger),
  stroked(
    join(
      line([CENTER - ARM, CENTER - ARM], [CENTER + ARM, CENTER + ARM]),
      line([CENTER + ARM, CENTER - ARM], [CENTER - ARM, CENTER + ARM]),
    ),
    PALETTE.paper,
    'heavy',
  ),
];

/** A yellow triangle with an exclamation mark. */
const TOP = CENTER - R;
const BASE = CENTER + R * 0.85;
const warning: IconDefinition = [
  filled(
    poly(
      [
        [CENTER, TOP],
        [CENTER + R, BASE],
        [CENTER - R, BASE],
      ],
      true,
    ),
    PALETTE.warn,
  ),
  stroked(line([CENTER, CENTER - R * 0.35], [CENTER, CENTER + R * 0.3]), PALETTE.ink, 'heavy'),
  filled(circle(CENTER, CENTER + R * 0.6, 0.9), PALETTE.ink),
];

/** A blue disc with a white letter i. */
const information: IconDefinition = [
  filled(circle(CENTER, CENTER, R), PALETTE.info),
  filled(circle(CENTER, CENTER - R * 0.45, 1.1), PALETTE.paper),
  stroked(line([CENTER, CENTER - R * 0.1], [CENTER, CENTER + R * 0.55]), PALETTE.paper, 'heavy'),
];

export const DIALOG_ALERT_ICONS = {
  alertStop: stop,
  alertWarning: warning,
  alertInformation: information,
} satisfies Record<string, IconDefinition>;
