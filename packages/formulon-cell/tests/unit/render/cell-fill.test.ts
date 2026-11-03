import { describe, expect, it } from 'vitest';
import { paintCellFill } from '../../../src/render/painters/cell.js';
import type { CellFormat } from '../../../src/store/store.js';
import type { ResolvedTheme } from '../../../src/theme/resolve.js';

/** Excel's 8x8 pattern bitmaps (X = foreground), anchored at sheet pixel (0,0). */
const TILES: Record<string, string[]> = {
  mediumGray: ['X.X.X.X.', '.X.X.X.X'],
  darkGray: ['X.XXX.XX', 'XXX.XXX.'],
  lightGray: ['..X...X.', 'X...X...'],
  darkHorizontal: ['........', '........', 'XXXXXXXX', 'XXXXXXXX'],
  darkVertical: ['..XX..XX'],
  darkDown: ['XX..XX..', '.XX..XX.', '..XX..XX', 'X..XX..X'],
  darkUp: ['..XX..XX', '.XX..XX.', 'XX..XX..', 'X..XX..X'],
  darkGrid: ['X..XX..X', '.XX..XX.', '.XX..XX.', 'X..XX..X'],
  darkTrellis: ['XXXXXXXX', '.XX..XX.', 'XXXXXXXX', 'X..XX..X'],
  lightHorizontal: ['........', '........', 'XXXXXXXX', '........'],
  lightVertical: ['..X...X.'],
  lightDown: ['X...X...', '.X...X..', '..X...X.', '...X...X'],
  lightUp: ['...X...X', '..X...X.', '.X...X..', 'X...X...'],
  lightGrid: ['..X...X.', '..X...X.', 'XXXXXXXX', '..X...X.'],
  lightTrellis: ['X...X...', '.X.X.X.X', '..X...X.', '.X.X.X.X'],
  gray125: ['X...X...', '........', '..X...X.', '........'],
  gray0625: ['......X.', '........', '..X.....', '........'],
};

const PATTERN_NAME: Record<string, NonNullable<CellFormat['fillPattern']>> = {
  mediumGray: 'gray50',
  darkGray: 'gray75',
  lightGray: 'gray25',
  darkHorizontal: 'darkHorizontal',
  darkVertical: 'darkVertical',
  darkDown: 'darkDown',
  darkUp: 'darkUp',
  darkGrid: 'darkGrid',
  darkTrellis: 'darkTrellis',
  lightHorizontal: 'lightHorizontal',
  lightVertical: 'lightVertical',
  lightDown: 'lightDown',
  lightUp: 'lightUp',
  lightGrid: 'lightGrid',
  lightTrellis: 'lightTrellis',
  gray125: 'gray125',
  gray0625: 'gray0625',
};

interface Spy {
  ctx: CanvasRenderingContext2D;
  /** Device-pixel set of the foreground ink. */
  ink: Set<string>;
  /** Background fill rect in CSS px. */
  bg: [number, number, number, number] | null;
  linear: [number, number, number, number][];
  radial: number;
}

const makeSpy = (dpr: number): Spy => {
  const spy: Spy = { ctx: null as never, ink: new Set(), bg: null, linear: [], radial: 0 };
  let style: unknown = '';
  let pending: [number, number, number, number][] = [];
  const mark = (x: number, y: number, w: number, h: number): void => {
    const px0 = Math.round(x * dpr);
    const py0 = Math.round(y * dpr);
    const px1 = Math.round((x + w) * dpr);
    const py1 = Math.round((y + h) * dpr);
    for (let py = py0; py < py1; py++)
      for (let px = px0; px < px1; px++) spy.ink.add(`${px},${py}`);
  };
  const grad = { addColorStop() {} };
  spy.ctx = {
    get fillStyle() {
      return style;
    },
    set fillStyle(v: unknown) {
      style = v;
    },
    save() {},
    restore() {},
    clip() {},
    beginPath() {
      pending = [];
    },
    moveTo() {},
    lineTo() {},
    closePath() {},
    rect(x: number, y: number, w: number, h: number) {
      pending.push([x, y, w, h]);
    },
    fill() {
      if (style === '#112233') for (const r of pending) mark(...r);
    },
    fillRect(x: number, y: number, w: number, h: number) {
      if (style === '#ffff00') spy.bg = [x, y, w, h];
      else if (style === '#112233') mark(x, y, w, h);
    },
    createLinearGradient(x0: number, y0: number, x1: number, y1: number) {
      spy.linear.push([x0, y0, x1, y1]);
      return grad;
    },
    createRadialGradient() {
      spy.radial++;
      return grad;
    },
  } as unknown as CanvasRenderingContext2D;
  return spy;
};

const paint = (
  spy: Spy,
  format: CellFormat,
  dpr: number,
  bounds = { x: 40, y: 20, w: 60, h: 30 },
) =>
  paintCellFill({
    ctx: spy.ctx,
    bounds,
    theme: { bg: '#fefefe' } as ResolvedTheme,
    value: { kind: 'blank' },
    formula: null,
    isActive: false,
    isInRange: false,
    format,
    dpr,
  });

describe('paintCellFill pattern tiles', () => {
  for (const dpr of [1, 2]) {
    for (const [name, rows] of Object.entries(TILES)) {
      it(`draws the ${name} bitmap at one device px per tile px (dpr ${dpr})`, () => {
        const spy = makeSpy(dpr);
        paint(
          spy,
          { fill: '#ffff00', fillPattern: PATTERN_NAME[name], fillPatternColor: '#112233' },
          dpr,
        );
        const bg = spy.bg;
        expect(bg).not.toBeNull();
        const [bx, by, bw, bh] = bg as [number, number, number, number];
        const x0 = Math.round(bx * dpr);
        const y0 = Math.round(by * dpr);
        const x1 = Math.round((bx + bw) * dpr);
        const y1 = Math.round((by + bh) * dpr);
        let mismatches = 0;
        for (let y = y0; y < y1; y++) {
          for (let x = x0; x < x1; x++) {
            const want = rows[y % rows.length]?.[x % 8] === 'X';
            if (spy.ink.has(`${x},${y}`) !== want) mismatches++;
          }
        }
        expect(mismatches).toBe(0);
      });
    }
  }
});

describe('paintCellFill path gradient', () => {
  it('fills a box gradient: one linear gradient per side, none radial', () => {
    const spy = makeSpy(2);
    paint(
      spy,
      {
        fillGradient: {
          type: 'path',
          degree: 0,
          left: 0.5,
          right: 0.5,
          top: 0.5,
          bottom: 0.5,
          stops: [
            { position: 0, color: '#ffffff' },
            { position: 1, color: '#00b050' },
          ],
        },
      },
      2,
    );
    expect(spy.radial).toBe(0);
    // Centre (70,35) to the left, right, top and bottom edges of the fill rect.
    expect(spy.linear).toHaveLength(4);
    const cx = spy.linear[0]?.[0];
    const cy = spy.linear[0]?.[1];
    expect(spy.linear.every(([x0, y0]) => x0 === cx && y0 === cy)).toBe(true);
    const ends = spy.linear.map(([, , x1, y1]) => [x1, y1]);
    expect(ends.filter(([x]) => x === cx).length).toBe(2);
    expect(ends.filter(([, y]) => y === cy).length).toBe(2);
  });
});

describe('paintCellFill pattern background', () => {
  it('clears to the sheet background when the pattern has no background colour', () => {
    const spy = makeSpy(2);
    const filled: string[] = [];
    const fillRect = spy.ctx.fillRect.bind(spy.ctx);
    spy.ctx.fillRect = (x, y, w, h) => {
      filled.push(String(spy.ctx.fillStyle));
      fillRect(x, y, w, h);
    };
    paint(spy, { fillPattern: 'gray50', fillPatternColor: '#112233' }, 2);
    expect(filled[0]).toBe('#fefefe');
  });
});
