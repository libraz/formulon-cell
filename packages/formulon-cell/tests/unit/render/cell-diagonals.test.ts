import { describe, expect, it } from 'vitest';
import { paintCellDiagonals } from '../../../src/render/painters/cell.js';
import type { CellBorders } from '../../../src/store/store.js';
import type { ResolvedTheme } from '../../../src/theme/resolve.js';

const trace = (borders: CellBorders, dpr: number) => {
  const lines: Array<{ from: [number, number]; width: number }> = [];
  let width = 1;
  let from: [number, number] = [0, 0];
  const ctx = {
    set strokeStyle(_v: string) {},
    set lineWidth(v: number) {
      width = v;
    },
    save() {},
    restore() {},
    beginPath() {},
    rect() {},
    clip() {},
    setLineDash() {},
    moveTo(x: number, y: number) {
      from = [x, y];
    },
    lineTo() {
      lines.push({ from, width });
    },
    stroke() {},
  } as unknown as CanvasRenderingContext2D;
  paintCellDiagonals({
    ctx,
    bounds: { x: 60, y: 20, w: 60, h: 24 },
    theme: { cellAutoFg: '#000' } as ResolvedTheme,
    value: { kind: 'blank' },
    formula: null,
    isActive: false,
    isInRange: false,
    format: { borders },
    dpr,
  });
  return lines;
};

describe('paintCellDiagonals rule-wide placement', () => {
  it('sits 1.41 device px right of the corner gridlines with 95% of the width', () => {
    const [thin] = trace({ diagonalDown: { style: 'thin' } }, 2);
    expect(thin?.from[0]).toBeCloseTo(60 - 0.5 + 1.41 / 2, 5);
    expect(thin?.width).toBeCloseTo(0.95, 5);
  });

  it('mirrors the shift for an up diagonal', () => {
    const [up] = trace({ diagonalUp: { style: 'thin' } }, 2);
    expect(up?.from[0]).toBeCloseTo(60 - 0.5 - 1.41 / 2, 5);
  });

  it('leaves medium diagonals on the corner gridlines at full width', () => {
    const [medium] = trace({ diagonalDown: { style: 'medium' } }, 2);
    expect(medium?.from[0]).toBeCloseTo(60 - 0.5, 5);
    expect(medium?.width).toBeCloseTo(2.67, 5);
  });
});
