import { describe, expect, it } from 'vitest';

import { createPivotTableFromRange } from '../../../../src/commands/pivot-table.js';
import { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { GridRenderer } from '../../../../src/render/grid.js';
import { createSpreadsheetStore, mutators } from '../../../../src/store/store.js';
import type { ResolvedTheme } from '../../../../src/theme/resolve.js';

const ACCENT = '#12ab34';
const theme = {
  accent: ACCENT,
  bg: '#ffffff',
  fg: '#000000',
  cellAutoFg: '#000000',
  rule: '#cccccc',
  fontUi: 'sans-serif',
} as unknown as ResolvedTheme;

interface Op {
  op: 'fillRect' | 'strokeRect';
  style: string;
  args: number[];
}

/** Paint a frame onto a recording context and return the rect operations. */
async function paintEmptyReport(
  select: (store: ReturnType<typeof createSpreadsheetStore>) => void,
) {
  const wb = await WorkbookHandle.createDefault();
  try {
    wb.setText({ sheet: 0, row: 0, col: 0 }, 'Region');
    wb.setText({ sheet: 0, row: 0, col: 1 }, 'Amt');
    wb.setText({ sheet: 0, row: 1, col: 0 }, 'East');
    wb.setNumber({ sheet: 0, row: 1, col: 1 }, 5);
    const created = createPivotTableFromRange(wb, {
      source: { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 },
      destination: { sheet: 0, row: 4, col: 0 },
      rowField: '',
      valueField: '',
    });
    expect(created.ok).toBe(true);
    const store = createSpreadsheetStore();
    select(store);
    const ops: Op[] = [];
    const props: Record<string, unknown> = {};
    const ctx = new Proxy(props, {
      get: (o, key) =>
        key === 'fillRect' || key === 'strokeRect'
          ? (...args: number[]) =>
              ops.push({
                op: key,
                style: String(key === 'fillRect' ? o.fillStyle : o.strokeStyle),
                args,
              })
          : key in o
            ? o[key as string]
            : () => ({ width: 0, addColorStop: () => undefined }),
      set: (o, key, value) => {
        o[key as string] = value;
        return true;
      },
    }) as unknown as CanvasRenderingContext2D;
    const host = {
      getBoundingClientRect: () => ({ width: 640, height: 480 }),
    } as unknown as HTMLElement;
    const renderer = new GridRenderer({
      host,
      canvas: { getContext: () => ctx, style: {} } as unknown as HTMLCanvasElement,
      getState: () => store.getState(),
      getTheme: () => theme,
      getWb: () => wb,
    });
    Object.assign(renderer as unknown as Record<string, number>, { cssWidth: 640, cssHeight: 480 });
    (renderer as unknown as { paintFrame(): void }).paintFrame();
    return ops;
  } finally {
    wb.dispose();
  }
}

/** Whether an op's rect reaches `rect` grown by `slack` px. */
const touches = (
  op: Op,
  rect: { x0: number; y0: number; x1: number; y1: number },
  slack: number,
): boolean => {
  const [x, y, w, h] = op.args as [number, number, number, number];
  return (
    x <= rect.x1 + slack &&
    x + w >= rect.x0 - slack &&
    y <= rect.y1 + slack &&
    y + h >= rect.y0 - slack
  );
};

describe('empty report placeholder and the selection', () => {
  it('shows no cursor chrome over the box when the active cell is its anchor', async () => {
    const ops = await paintEmptyReport((store) =>
      mutators.setActive(store, { sheet: 0, row: 4, col: 0 }),
    );
    const placeholder = ops.find((o) => o.op === 'strokeRect' && o.style === '#999999');
    expect(placeholder).toBeDefined();
    const [px, py, pw, ph] = (placeholder as Op).args as [number, number, number, number];
    const box = { x0: px, y0: py, x1: px + pw, y1: py + ph };
    const accent = ops.filter((o) => o.style === ACCENT && touches(o, box, 3));
    expect(accent).toEqual([]);
  });

  it('keeps the part of a larger selection that lies outside the box', async () => {
    const ops = await paintEmptyReport((store) =>
      mutators.setRange(store, { sheet: 0, r0: 2, c0: 0, r1: 4, c1: 0 }),
    );
    expect(ops.some((o) => o.style === ACCENT)).toBe(true);
  });
});
