import { describe, expect, it } from 'vitest';
import {
  FILL_HANDLE_SIZE,
  paintCellBorders,
  paintFillHandle,
  textBaselineY,
} from '../../../src/render/painters.js';
import type { ResolvedTheme } from '../../../src/theme/resolve.js';

/**
 * Minimal canvas spy. We only need to capture rect-fill calls + the active
 * `fillStyle` at each call so tests can assert paint order and colour.
 */
function makeCtxSpy(): {
  ctx: CanvasRenderingContext2D;
  fills: Array<{ style: string; rect: [number, number, number, number] }>;
} {
  const fills: Array<{ style: string; rect: [number, number, number, number] }> = [];
  let style = '';
  const ctx = {
    get fillStyle(): string {
      return style;
    },
    set fillStyle(v: string) {
      style = v;
    },
    fillRect(x: number, y: number, w: number, h: number): void {
      fills.push({ style, rect: [x, y, w, h] });
    },
    save(): void {},
    restore(): void {},
  } as unknown as CanvasRenderingContext2D;
  return { ctx, fills };
}

function makeStrokeSpy(): {
  ctx: CanvasRenderingContext2D;
  strokes: Array<{ style: string; width: number; dash: number[] }>;
} {
  const strokes: Array<{ style: string; width: number; dash: number[] }> = [];
  let strokeStyle = '';
  let lineWidth = 1;
  let dash: number[] = [];
  const ctx = {
    get strokeStyle(): string {
      return strokeStyle;
    },
    set strokeStyle(v: string) {
      strokeStyle = v;
    },
    get lineWidth(): number {
      return lineWidth;
    },
    set lineWidth(v: number) {
      lineWidth = v;
    },
    save(): void {},
    restore(): void {},
    beginPath(): void {},
    moveTo(): void {},
    lineTo(): void {},
    setLineDash(v: number[]): void {
      dash = [...v];
    },
    stroke(): void {
      strokes.push({ style: strokeStyle, width: lineWidth, dash: [...dash] });
    },
  } as unknown as CanvasRenderingContext2D;
  return { ctx, strokes };
}

const theme = (over: Partial<ResolvedTheme> = {}): ResolvedTheme =>
  ({
    bg: '#fff',
    bgRail: '#eee',
    bgElev: '#fff',
    bgHeader: '#eee',
    fg: '#000',
    fgMute: '#666',
    fgFaint: '#aaa',
    fgStrong: '#000',
    rule: '#ccc',
    ruleStrong: '#999',
    accent: '#0078d4',
    accentSoft: 'rgba(0,120,212,0.1)',
    cellErrorFg: '#c00',
    cellFormulaFg: '#333',
    cellBoolFg: '#060',
    cellNumFg: '#000',
    hoverStripe: 'rgba(0,0,0,0.04)',
    headerFg: '#666',
    headerFgActive: '#000',
    fontUi: 'sans-serif',
    fontMono: 'monospace',
    textCell: 13,
    textHeader: 11.5,
    ...over,
  }) as ResolvedTheme;

describe('paintFillHandle', () => {
  it('exports a 6px size constant', () => {
    expect(FILL_HANDLE_SIZE).toBe(6);
  });

  it('paints a white border ring then an accent square at the cell corner', () => {
    const { ctx, fills } = makeCtxSpy();
    // Cell rect at (10, 20), size 100×30. Bottom-right corner sits at (110, 50).
    const bounds = { x: 10, y: 20, w: 100, h: 30 };
    const hit = paintFillHandle(ctx, bounds, theme({ accent: '#0078d4' }));

    // Expect two fillRect calls — outer white halo, inner accent square.
    expect(fills).toHaveLength(2);
    const [halo, square] = fills;

    // Visible square is centred on the bottom-right corner; half-bleeds outside
    // the cell rect. With FILL_HANDLE_SIZE=6 the inner rect starts at
    //   x = 10 + 100 - 3 = 107
    //   y = 20 +  30 - 3 = 47
    expect(square?.style).toBe('#0078d4');
    expect(square?.rect).toEqual([107, 47, 6, 6]);

    // Halo wraps the square with a 1px white border.
    expect(halo?.style).toBe('#ffffff');
    expect(halo?.rect).toEqual([106, 46, 8, 8]);

    // Returned hit-rect spans the visible (halo) area for comfortable grabbing.
    expect(hit).toEqual({ x: 106, y: 46, w: 8, h: 8 });
  });

  it('falls back to #0078d4 when theme.accent is empty', () => {
    const { ctx, fills } = makeCtxSpy();
    paintFillHandle(ctx, { x: 0, y: 0, w: 50, h: 20 }, theme({ accent: '' }));
    const square = fills[1];
    expect(square?.style).toBe('#0078d4');
  });
});

describe('paintCellBorders', () => {
  it('uses the Excel-like automatic border color instead of the gridline color', () => {
    const { ctx, strokes } = makeStrokeSpy();
    paintCellBorders({
      ctx,
      bounds: { x: 0, y: 0, w: 80, h: 24 },
      theme: theme({ fgStrong: '#111111', ruleStrong: '#999999' }),
      value: { kind: 'blank' },
      formula: null,
      isActive: false,
      isInRange: false,
      format: { borders: { bottom: true } },
    });

    expect(strokes).toHaveLength(1);
    expect(strokes[0]?.style).toBe('#111111');
  });

  it('honors explicit border colors from cell format', () => {
    const { ctx, strokes } = makeStrokeSpy();
    paintCellBorders({
      ctx,
      bounds: { x: 0, y: 0, w: 80, h: 24 },
      theme: theme({ fgStrong: '#111111' }),
      value: { kind: 'blank' },
      formula: null,
      isActive: false,
      isInRange: false,
      format: { borders: { bottom: { style: 'thin', color: '#c00000' } } },
    });

    expect(strokes[0]?.style).toBe('#c00000');
  });
});

describe('textBaselineY', () => {
  it('keeps bottom-aligned text visually inside the cell padding', () => {
    const bounds = { x: 10, y: 20, w: 80, h: 20 };
    const box = { ascent: 10, descent: 3 };
    expect(textBaselineY(bounds, box, 'bottom', 4)).toBe(33);
  });

  it('centers the measured text box for middle alignment', () => {
    const bounds = { x: 0, y: 0, w: 80, h: 24 };
    const box = { ascent: 10, descent: 4 };
    expect(textBaselineY(bounds, box, 'middle', 4)).toBe(15);
  });
});
