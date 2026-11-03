/**
 * @file Conditional-format icon-set artwork, drawn with canvas primitives.
 *
 * Every glyph is laid out on a 12-unit grid whose origin is the top-left of
 * the icon box; the caller translates and scales so one unit is one CSS px at
 * 100% zoom. Shapes, colours and slot mappings follow the desktop
 * spreadsheet's built-in sets: discs and signs fill 11 units, block arrows
 * 11x12 or 12x12, and rating bars run one unit past the box on the right.
 * Outlines are 1 unit wide and sit inside the shape.
 */
import type { ConditionalIconSet } from '../../store/store.js';

interface Tone {
  rim: string;
  light: string;
  dark: string;
}

const RED: Tone = { rim: '#9e3c31', light: '#f69b9c', dark: '#f06b6b' };
const YELLOW: Tone = { rim: '#db7d17', light: '#fee89e', dark: '#f8c948' };
const GREEN: Tone = { rim: '#326b2b', light: '#75c48a', dark: '#4caa61' };
const GRAY: Tone = { rim: '#4d4d4d', light: '#b5b5b5', dark: '#959595' };
const BLACK: Tone = { rim: '#0b0b0b', light: '#5a5a5a', dark: '#161616' };

type Point = readonly [number, number];

/** Downward block arrow, 11 wide x 12 long. */
const ARROW_DOWN: readonly Point[] = [
  [3, 0],
  [8, 0],
  [8, 6],
  [11, 6],
  [6, 11],
  [6, 12],
  [5, 12],
  [5, 11],
  [0, 6],
  [3, 6],
];

/** Down-right block arrow, 12x12, head in the bottom-right corner. */
const ARROW_DOWN_RIGHT: readonly Point[] = [
  [0, 2],
  [2, 0],
  [3.5, 0],
  [9.75, 6.25],
  [12, 4],
  [12, 12],
  [4, 12],
  [6.25, 9.75],
  [0, 3.5],
];

type ArrowDir = 'down' | 'downRight' | 'right' | 'upRight' | 'up';

const arrowPoints = (dir: ArrowDir): Point[] => {
  switch (dir) {
    case 'down':
      return [...ARROW_DOWN];
    case 'up':
      return ARROW_DOWN.map(([x, y]) => [x, 12 - y]);
    case 'right':
      return ARROW_DOWN.map(([x, y]) => [y, x]);
    case 'downRight':
      return [...ARROW_DOWN_RIGHT];
    case 'upRight':
      return ARROW_DOWN_RIGHT.map(([x, y]) => [x, 12 - y]);
  }
};

/** Direction and tone per slot, lowest bucket first. */
const ARROW_SLOTS: Readonly<Record<3 | 4 | 5, readonly [ArrowDir, Tone][]>> = {
  3: [
    ['down', RED],
    ['right', YELLOW],
    ['up', GREEN],
  ],
  4: [
    ['down', RED],
    ['downRight', YELLOW],
    ['upRight', YELLOW],
    ['up', GREEN],
  ],
  5: [
    ['down', RED],
    ['downRight', YELLOW],
    ['right', YELLOW],
    ['upRight', YELLOW],
    ['up', GREEN],
  ],
};

const ARROW_SETS: Partial<Record<ConditionalIconSet, { slots: 3 | 4 | 5; gray: boolean }>> = {
  arrows3: { slots: 3, gray: false },
  arrowsGray3: { slots: 3, gray: true },
  arrows4: { slots: 4, gray: false },
  arrowsGray4: { slots: 4, gray: true },
  arrows5: { slots: 5, gray: false },
  arrowsGray5: { slots: 5, gray: true },
};

const pick = <T>(list: readonly T[], slot: number): T =>
  list[Math.max(0, Math.min(list.length - 1, Math.trunc(slot) || 0))] as T;

const polygon = (ctx: CanvasRenderingContext2D, pts: readonly Point[]): void => {
  ctx.beginPath();
  pts.forEach(([x, y], i) => {
    if (i === 0) ctx.moveTo(x, y);
    else ctx.lineTo(x, y);
  });
  ctx.closePath();
};

const disc = (ctx: CanvasRenderingContext2D, cx: number, cy: number, r: number): void => {
  ctx.beginPath();
  ctx.arc(cx, cy, r, 0, Math.PI * 2);
};

const roundedRect = (
  ctx: CanvasRenderingContext2D,
  x: number,
  y: number,
  w: number,
  h: number,
  r: number,
): void => {
  ctx.beginPath();
  ctx.moveTo(x + r, y);
  ctx.lineTo(x + w - r, y);
  ctx.arc(x + w - r, y + r, r, -Math.PI / 2, 0);
  ctx.lineTo(x + w, y + h - r);
  ctx.arc(x + w - r, y + h - r, r, 0, Math.PI / 2);
  ctx.lineTo(x + r, y + h);
  ctx.arc(x + r, y + h - r, r, Math.PI / 2, Math.PI);
  ctx.lineTo(x, y + r);
  ctx.arc(x + r, y + r, r, Math.PI, Math.PI * 1.5);
  ctx.closePath();
};

/** Fill the current path with a top-left → bottom-right gradient, then draw
 *  a `width`-unit rim inside it. */
const shade = (ctx: CanvasRenderingContext2D, tone: Tone, width = 1): void => {
  const gradient = ctx.createLinearGradient(0, 0, 12, 12);
  // Excel's glyphs hold their highlight to about a third of the diagonal and
  // reach their shade by two thirds.
  gradient.addColorStop(0.36, tone.light);
  gradient.addColorStop(0.62, tone.dark);
  ctx.fillStyle = gradient;
  ctx.fill();
  ctx.save();
  ctx.clip();
  ctx.strokeStyle = tone.rim;
  ctx.lineWidth = width * 2;
  ctx.stroke();
  ctx.restore();
};

/** Stroke the current path twice: a wide rim, then a narrower fill. */
const outlinedStroke = (
  ctx: CanvasRenderingContext2D,
  rim: string,
  fill: string,
  width: number,
): void => {
  ctx.lineCap = 'round';
  ctx.lineJoin = 'round';
  ctx.strokeStyle = rim;
  ctx.lineWidth = width;
  ctx.stroke();
  ctx.strokeStyle = fill;
  ctx.lineWidth = width - 1.4;
  ctx.stroke();
};

const TRAFFIC: readonly Tone[] = [RED, YELLOW, GREEN];

function drawRimLight(ctx: CanvasRenderingContext2D, slot: number): void {
  // Black housing with a bevel lighter at the top-left.
  const bevel = ctx.createLinearGradient(0, 0, 12, 12);
  bevel.addColorStop(0, '#9a9999');
  bevel.addColorStop(1, '#353535');
  ctx.fillStyle = bevel;
  roundedRect(ctx, 0.5, 0.5, 11.25, 11.25, 1.5);
  ctx.fill();
  ctx.fillStyle = '#0a0a0a';
  roundedRect(ctx, 1.5, 1.5, 9.25, 9.25, 1);
  ctx.fill();
  const [hi, body, edge] = pick(
    [
      ['#f9c6c2', '#ed4430', '#b23b2a'],
      ['#fce4b9', '#f2a40b', '#c48616'],
      ['#cde6c7', '#46a542', '#2e7a2c'],
    ],
    slot,
  );
  const glow = ctx.createRadialGradient(4.8, 4.6, 0, 6.1, 6.2, 4.2);
  glow.addColorStop(0, hi);
  glow.addColorStop(0.45, body);
  glow.addColorStop(1, edge);
  ctx.fillStyle = glow;
  disc(ctx, 6.1, 6.2, 4);
  ctx.fill();
}

function drawFlag(ctx: CanvasRenderingContext2D, slot: number): void {
  const [edge, body, light] = pick(
    [
      ['#b23324', '#eb5145', '#f37977'],
      ['#b48003', '#f0b51d', '#fbd050'],
      ['#377827', '#58ab3d', '#84c761'],
    ],
    slot,
  );
  const cloth = ctx.createLinearGradient(1.5, 0, 10, 0);
  cloth.addColorStop(0, edge);
  cloth.addColorStop(0.3, body);
  cloth.addColorStop(0.6, light);
  cloth.addColorStop(1, body);
  ctx.fillStyle = cloth;
  polygon(ctx, [
    [1.6, 0],
    [6, 0],
    [6.6, 3],
    [9.8, 3],
    [9.2, 5.8],
    [3.3, 6],
  ]);
  ctx.fill();
  ctx.strokeStyle = '#4a4848';
  ctx.lineWidth = 1.1;
  ctx.lineCap = 'butt';
  ctx.beginPath();
  ctx.moveTo(1.2, 0.3);
  ctx.lineTo(4.6, 11.8);
  ctx.stroke();
}

function drawCircledSymbol(ctx: CanvasRenderingContext2D, slot: number): void {
  const idx = Math.max(0, Math.min(2, Math.trunc(slot) || 0));
  disc(ctx, 5.5, 5.5, 5.5);
  shade(ctx, pick(TRAFFIC, idx));
  ctx.strokeStyle = 'rgba(255,255,255,0.92)';
  ctx.fillStyle = 'rgba(255,255,255,0.92)';
  ctx.lineWidth = 1.5;
  ctx.lineCap = 'butt';
  ctx.beginPath();
  if (idx === 0) {
    ctx.moveTo(3.2, 3.2);
    ctx.lineTo(7.8, 7.8);
    ctx.moveTo(7.8, 3.2);
    ctx.lineTo(3.2, 7.8);
    ctx.stroke();
  } else if (idx === 1) {
    ctx.fillRect(4.75, 2.5, 1.5, 4.3);
    ctx.fillRect(4.75, 8, 1.5, 1.5);
  } else {
    ctx.moveTo(2.6, 5.9);
    ctx.lineTo(4.6, 7.9);
    ctx.lineTo(8.4, 4.1);
    ctx.stroke();
  }
}

function drawPlainSymbol(ctx: CanvasRenderingContext2D, slot: number): void {
  const idx = Math.max(0, Math.min(2, Math.trunc(slot) || 0));
  if (idx === 0) {
    ctx.beginPath();
    ctx.moveTo(2, 2);
    ctx.lineTo(10, 10);
    ctx.moveTo(10, 2);
    ctx.lineTo(2, 10);
    outlinedStroke(ctx, '#bb5147', '#ec8582', 3);
  } else if (idx === 1) {
    polygon(ctx, [
      [5, 0.75],
      [8, 0.75],
      [7.4, 7],
      [5.6, 7],
    ]);
    shade(ctx, { rim: '#db8619', light: '#f7dca8', dark: '#efc27c' });
    roundedRect(ctx, 5, 8.6, 3, 2.4, 0.6);
    shade(ctx, { rim: '#db8619', light: '#f7dca8', dark: '#efc27c' });
  } else {
    ctx.beginPath();
    ctx.moveTo(1.4, 5.9);
    ctx.lineTo(4.4, 8.9);
    ctx.lineTo(10.6, 1.6);
    outlinedStroke(ctx, '#3f7d3b', '#5bb36d', 3);
  }
}

function drawSign(ctx: CanvasRenderingContext2D, slot: number): void {
  const idx = Math.max(0, Math.min(2, Math.trunc(slot) || 0));
  if (idx === 0) {
    polygon(ctx, [
      [6, 0],
      [12, 6],
      [6, 12],
      [0, 6],
    ]);
    shade(ctx, RED);
  } else if (idx === 1) {
    polygon(ctx, [
      [6, 0],
      [11.6, 10.6],
      [0.4, 10.6],
    ]);
    shade(ctx, { rim: '#da8000', light: '#fde796', dark: '#fbcf1c' });
  } else {
    disc(ctx, 6, 6, 5.5);
    shade(ctx, GREEN);
  }
}

/** Four signal-strength bars; `lit` of them in navy, the rest light gray. */
function drawRating(ctx: CanvasRenderingContext2D, lit: number): void {
  const tops = [7.5, 5, 2, 0];
  for (let i = 0; i < 4; i += 1) {
    const on = i < lit;
    const x = i * 3.5;
    const top = tops[i] ?? 0;
    ctx.fillStyle = on ? '#153a6a' : '#a7a7a7';
    roundedRect(ctx, x, top, 2.5, 12 - top, 1.25);
    ctx.fill();
    ctx.fillStyle = on ? '#6d8db1' : '#d6d6d6';
    ctx.fillRect(x + 0.75, top + 1, 1, 10 - top);
  }
}

function drawQuarters(ctx: CanvasRenderingContext2D, slot: number): void {
  const idx = Math.max(0, Math.min(4, Math.trunc(slot) || 0));
  disc(ctx, 5.5, 5.5, 5.5);
  ctx.fillStyle = idx === 4 ? '#000000' : '#ffffff';
  ctx.fill();
  if (idx > 0 && idx < 4) {
    ctx.fillStyle = '#000000';
    ctx.beginPath();
    ctx.moveTo(5.5, 5.5);
    ctx.arc(5.5, 5.5, 5.5, -Math.PI / 2, -Math.PI / 2 + (idx / 4) * Math.PI * 2);
    ctx.closePath();
    ctx.fill();
  }
  disc(ctx, 5.5, 5.5, 5.5);
  ctx.save();
  ctx.clip();
  ctx.strokeStyle = '#0b0b0b';
  ctx.lineWidth = 1.6;
  ctx.stroke();
  ctx.restore();
}

function drawStar(ctx: CanvasRenderingContext2D, slot: number): void {
  const idx = Math.max(0, Math.min(2, Math.trunc(slot) || 0));
  const trace = (): void => {
    ctx.beginPath();
    for (let i = 0; i < 10; i += 1) {
      const angle = (i * Math.PI) / 5 - Math.PI / 2;
      const radius = i % 2 === 0 ? 6 : 2.6;
      const px = 6 + Math.cos(angle) * radius;
      const py = 6.4 + Math.sin(angle) * radius;
      if (i === 0) ctx.moveTo(px, py);
      else ctx.lineTo(px, py);
    }
    ctx.closePath();
  };
  ctx.fillStyle = '#f5b800';
  if (idx === 2) {
    trace();
    ctx.fill();
  } else if (idx === 1) {
    // Left half lit.
    ctx.save();
    ctx.beginPath();
    ctx.rect(0, 0, 6, 12);
    ctx.clip();
    trace();
    ctx.fill();
    ctx.restore();
  }
  trace();
  ctx.strokeStyle = '#c98f00';
  ctx.lineWidth = 0.8;
  ctx.lineJoin = 'round';
  ctx.stroke();
}

function drawTriangle(ctx: CanvasRenderingContext2D, slot: number): void {
  const idx = Math.max(0, Math.min(2, Math.trunc(slot) || 0));
  if (idx === 1) {
    ctx.beginPath();
    ctx.rect(1, 4.5, 10, 3);
    shade(ctx, YELLOW);
    return;
  }
  const down: Point[] = [
    [0.5, 2.5],
    [11.5, 2.5],
    [6, 10],
  ];
  polygon(ctx, idx === 0 ? down : down.map(([x, y]) => [x, 12.5 - y]));
  shade(ctx, idx === 0 ? RED : GREEN);
}

/** Paint one icon on the 12-unit grid at the current transform's origin.
 *  `slot` is 0-based, lowest bucket first; out-of-range slots clamp. */
export function drawConditionalIconArt(
  ctx: CanvasRenderingContext2D,
  kind: ConditionalIconSet,
  slot: number,
): void {
  const arrows = ARROW_SETS[kind];
  if (arrows) {
    const [dir, tone] = pick(ARROW_SLOTS[arrows.slots], slot);
    polygon(ctx, arrowPoints(dir));
    shade(ctx, arrows.gray ? GRAY : tone);
    return;
  }
  switch (kind) {
    case 'traffic3':
      disc(ctx, 5.5, 5.5, 5.5);
      shade(ctx, pick(TRAFFIC, slot));
      return;
    case 'traffic4':
      disc(ctx, 5.5, 5.5, 5.5);
      shade(ctx, pick([BLACK, RED, YELLOW, GREEN], slot));
      return;
    case 'trafficRim3':
      drawRimLight(ctx, slot);
      return;
    case 'redToBlack4':
      // Flat discs, no rim or shading.
      disc(ctx, 5.5, 5.5, 5.5);
      ctx.fillStyle = pick(['#000000', '#bfbfbf', '#fd9088', '#fd2000'], slot);
      ctx.fill();
      return;
    case 'signs3':
      drawSign(ctx, slot);
      return;
    case 'symbolsCircled3':
      drawCircledSymbol(ctx, slot);
      return;
    case 'symbols3':
      drawPlainSymbol(ctx, slot);
      return;
    case 'flags3':
      drawFlag(ctx, slot);
      return;
    case 'ratings4':
      drawRating(ctx, Math.max(0, Math.min(3, Math.trunc(slot) || 0)) + 1);
      return;
    case 'ratings5':
      drawRating(ctx, Math.max(0, Math.min(4, Math.trunc(slot) || 0)));
      return;
    case 'quarters5':
      drawQuarters(ctx, slot);
      return;
    case 'stars3':
      drawStar(ctx, slot);
      return;
    case 'triangles3':
      drawTriangle(ctx, slot);
      return;
    case 'boxes5': {
      // 2x2 squares; the slot sets how many are filled.
      const filled = Math.max(0, Math.min(4, Math.trunc(slot) || 0));
      for (let i = 0; i < 4; i += 1) {
        const x = (i % 2) * 6 + 0.5;
        const y = Math.floor(i / 2) * 6 + 0.5;
        ctx.fillStyle = i < filled ? '#3f6fb6' : '#d6dbe3';
        ctx.fillRect(x, y, 5, 5);
      }
      return;
    }
    case 'bars5': {
      const lit = Math.max(0, Math.min(4, Math.trunc(slot) || 0)) + 1;
      for (let i = 0; i < 5; i += 1) {
        const h = 3 + i * 2.25;
        ctx.fillStyle = i < lit ? '#3f6fb6' : '#d6dbe3';
        ctx.fillRect(i * 2.4, 12 - h, 1.8, h);
      }
      return;
    }
    default:
      return;
  }
}
