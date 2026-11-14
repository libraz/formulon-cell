import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { en, ja } from '../../../src/i18n/strings.js';
import { attachHover } from '../../../src/interact/hover.js';
import { gridHitState } from '../../../src/render/grid/hit-state.js';
import { createSpreadsheetStore, type SpreadsheetStore } from '../../../src/store/store.js';

const tip = (): HTMLElement | null => document.querySelector<HTMLElement>('.fc-hover-tip');

describe('stale value hover tip', () => {
  let grid: HTMLElement;
  let store: SpreadsheetStore;
  let stale: boolean;
  const move = (x: number, y: number): void => {
    grid.dispatchEvent(new PointerEvent('pointermove', { clientX: x, clientY: y, bubbles: true }));
  };

  beforeEach(() => {
    grid = document.createElement('div');
    document.body.appendChild(grid);
    grid.getBoundingClientRect = (): DOMRect =>
      ({
        left: 0,
        top: 0,
        right: 800,
        bottom: 600,
        width: 800,
        height: 600,
        x: 0,
        y: 0,
      }) as DOMRect;
    store = createSpreadsheetStore();
    stale = true;
    gridHitState(grid).errorTriangles = [
      { rect: { x: 100, y: 40, w: 6, h: 6 }, addr: { sheet: 0, row: 1, col: 1 }, kind: 'error' },
    ];
  });
  afterEach(() => document.body.replaceChildren());

  const attach = (strings = ja) =>
    attachHover({
      grid,
      store,
      staleTip: {
        strings: () => strings,
        isStale: () => stale,
      },
    });

  it('shows the bundle heading and body over a stale cell triangle', () => {
    const h = attach();
    move(103, 43);
    expect(tip()?.hidden).toBe(false);
    expect(tip()?.textContent).toBe(
      `${ja.mac.errorMenu.staleHeading}\n${ja.mac.errorMenu.staleBody}`,
    );
    expect(tip()?.textContent).toContain('このセルの値は古くなっています');
    h.detach();
  });

  it('follows the active locale', () => {
    const h = attach(en);
    move(103, 43);
    expect(tip()?.textContent).toBe(
      'Stale Value\nThe value in this cell is stale. Calculate the workbook to see an up-to-date value.',
    );
    h.detach();
  });

  it('stays hidden off the triangle or when the cell is not stale', () => {
    const h = attach();
    move(300, 300);
    expect(tip()?.hidden).toBe(true);
    stale = false;
    move(103, 43);
    expect(tip()?.hidden).toBe(true);
    h.detach();
  });
});
