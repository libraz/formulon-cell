import { describe, expect, it } from 'vitest';

import { isFillHandleHit } from '../../../../src/interact/pointer-targets.js';
import { buildColLayout, buildRowLayout, layoutForView } from '../../../../src/render/geometry.js';
import { GridRenderer, gridHitState } from '../../../../src/render/grid.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../../src/store/store.js';
import type { ResolvedTheme } from '../../../../src/theme/resolve.js';

const theme = {
  accent: '#0078d4',
  bg: '#ffffff',
  fg: '#1f1f1f',
  rule: '#cccccc',
  fontUi: 'Arial',
  textCell: 12,
} as unknown as ResolvedTheme;

const ctxStub = (): CanvasRenderingContext2D =>
  new Proxy(
    {},
    {
      get: (_target, key) => (key === 'measureText' ? () => ({ width: 0 }) : () => undefined),
      set: () => true,
    },
  ) as unknown as CanvasRenderingContext2D;

/** Paint the active-cell layer of `store` on a fresh grid surface. */
function paintActive(store: SpreadsheetStore): HTMLElement {
  const host = document.createElement('div');
  host.getBoundingClientRect = () => ({ width: 640, height: 480 }) as DOMRect;
  const state = store.getState();
  const renderer = new GridRenderer({
    host,
    canvas: { getContext: () => ctxStub(), style: {} } as unknown as HTMLCanvasElement,
    getState: () => state,
    getTheme: () => theme,
  });
  const view = { ...state, layout: layoutForView(state) };
  (renderer as unknown as { paintActive(...args: unknown[]): void }).paintActive(
    view,
    theme,
    buildColLayout(view.layout, view.viewport),
    buildRowLayout(view.layout, view.viewport),
  );
  return host;
}

describe('two grids on one page', () => {
  it('keep their own fill handle after the other one paints', () => {
    const first = createSpreadsheetStore();
    mutators.setActive(first, { sheet: 0, row: 0, col: 0 });
    const second = createSpreadsheetStore();
    mutators.setActive(second, { sheet: 0, row: 9, col: 4 });

    const firstHost = paintActive(first);
    const secondHost = paintActive(second);

    const firstHandle = gridHitState(firstHost).fillHandle;
    const secondHandle = gridHitState(secondHost).fillHandle;
    if (!firstHandle || !secondHandle) throw new Error('expected both grids to paint a handle');
    expect(firstHandle).not.toEqual(secondHandle);
    expect(isFillHandleHit(gridHitState(firstHost), firstHandle.x, firstHandle.y)).toBe(true);
    expect(isFillHandleHit(gridHitState(firstHost), secondHandle.x, secondHandle.y)).toBe(false);
  });

  it('keep their own validation chevron', () => {
    const listed = createSpreadsheetStore();
    mutators.setCellFormat(
      listed,
      { sheet: 0, row: 0, col: 0 },
      { validation: { kind: 'list', source: ['a', 'b'] } },
    );
    mutators.setActive(listed, { sheet: 0, row: 0, col: 0 });
    const plain = createSpreadsheetStore();

    const listedHost = paintActive(listed);
    const plainHost = paintActive(plain);

    expect(gridHitState(listedHost).validationChevron).toMatchObject({ row: 0, col: 0 });
    expect(gridHitState(plainHost).validationChevron).toBeNull();
  });
});
