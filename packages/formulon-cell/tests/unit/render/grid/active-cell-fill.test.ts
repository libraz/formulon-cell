import { describe, expect, it } from 'vitest';

import { buildColLayout, buildRowLayout, layoutForView } from '../../../../src/render/geometry.js';
import { GridRenderer } from '../../../../src/render/grid.js';
import { hasCellFill } from '../../../../src/render/painters/cell.js';
import {
  type CellFormat,
  createSpreadsheetStore,
  mutators,
  type State,
} from '../../../../src/store/store.js';
import type { ResolvedTheme } from '../../../../src/theme/resolve.js';

const theme = {
  accent: '#0078d4',
  accentSoft: 'rgba(0,120,212,0.1)',
  bg: '#ffffff',
  bgElev: '#f5f5f5',
  bgRail: '#eeeeee',
  fg: '#1f1f1f',
  cellAutoFg: '#000000',
  rule: '#cccccc',
  fontUi: 'Arial',
  fontMono: 'monospace',
  textCell: 12,
} as unknown as ResolvedTheme;

/** Records the fillStyle behind every fillRect; every other call is a no-op. */
function recordingCtx(): { ctx: CanvasRenderingContext2D; fillStyles: string[] } {
  const fillStyles: string[] = [];
  const state: Record<string, unknown> = { fillStyle: '', font: '', globalAlpha: 1 };
  const ctx = new Proxy(state, {
    get(target, prop: string) {
      if (prop === 'fillRect') return () => fillStyles.push(String(target.fillStyle));
      if (prop === 'measureText') return () => ({ width: 0 });
      if (prop in target) return target[prop];
      return () => ({ addColorStop() {} });
    },
    set(target, prop: string, value) {
      target[prop] = value;
      return true;
    },
  }) as unknown as CanvasRenderingContext2D;
  return { ctx, fillStyles };
}

function paintActiveCellWith(format: Partial<CellFormat>): string[] {
  const store = createSpreadsheetStore();
  mutators.setCellFormat(store, { sheet: 0, row: 1, col: 1 }, format);
  mutators.setActive(store, { sheet: 0, row: 1, col: 1 });
  const state: State = store.getState();
  const { ctx, fillStyles } = recordingCtx();
  const renderer = new GridRenderer({
    host: { getBoundingClientRect: () => ({ width: 640, height: 480 }) } as unknown as HTMLElement,
    canvas: { getContext: () => ctx, style: {} } as unknown as HTMLCanvasElement,
    getState: () => state,
    getTheme: () => theme,
  });
  const view = { ...state, layout: layoutForView(state) };
  (
    renderer as unknown as {
      paintCells(s: State, t: ResolvedTheme, cols: unknown, rows: unknown): void;
    }
  ).paintCells(
    view,
    theme,
    buildColLayout(view.layout, view.viewport),
    buildRowLayout(view.layout, view.viewport),
  );
  return fillStyles;
}

describe('cell own-fill predicate', () => {
  it('treats fill, fillPattern and fillGradient alike', () => {
    expect(hasCellFill(undefined)).toBe(false);
    expect(hasCellFill({ bold: true })).toBe(false);
    expect(hasCellFill({ fill: '#ff0000' })).toBe(true);
    expect(hasCellFill({ fillPattern: 'gray50' })).toBe(true);
    expect(hasCellFill({ fillGradient: { type: 'linear', angle: 0, stops: [] } as never })).toBe(
      true,
    );
  });

  it('does not paint the active highlight over a pattern-only fill', () => {
    expect(paintActiveCellWith({ fillPattern: 'gray50' })).not.toContain(theme.bgElev);
  });

  it('still paints the active highlight on an unfilled cell', () => {
    expect(paintActiveCellWith({ bold: true })).toContain(theme.bgElev);
  });
});
