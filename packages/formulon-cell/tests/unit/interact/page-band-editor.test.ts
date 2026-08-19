import { describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { attachPageBandEditor } from '../../../src/interact/page-band-editor.js';
import type { PageBandHit } from '../../../src/render/grid/page-view.js';
import { createSpreadsheetStore, getPageSetup, mutators } from '../../../src/store/store.js';

const bandEvent = (over: Partial<PageBandHit> = {}): CustomEvent<PageBandHit> =>
  new CustomEvent('fc:editpageband', {
    detail: {
      rect: { x: 40, y: 8, w: 120, h: 22 },
      slot: 'center',
      kind: 'header',
      ...over,
    },
  });

const setup = () => {
  const grid = document.createElement('div');
  document.body.appendChild(grid);
  const store = createSpreadsheetStore();
  const history = new History();
  const handle = attachPageBandEditor({ grid, store, history });
  return { grid, store, history, handle };
};

const field = (grid: HTMLElement): HTMLInputElement | null =>
  grid.querySelector('.fc-pageband-editor');

describe('attachPageBandEditor', () => {
  it('opens an input over the clicked slot, seeded with what the slot holds', () => {
    const { grid, store, handle } = setup();
    mutators.setPageSetup(store, 0, { headerCenter: 'Quarterly report' });

    grid.dispatchEvent(bandEvent());

    const input = field(grid);
    expect(input?.value).toBe('Quarterly report');
    expect(input?.style.left).toBe('40px');
    expect(input?.style.width).toBe('120px');
    expect(input?.style.textAlign).toBe('center');
    expect(handle.isOpen()).toBe(true);
    handle.detach();
  });

  it('commits on Enter and pushes one undoable page-setup change', () => {
    const { grid, store, history, handle } = setup();

    grid.dispatchEvent(bandEvent({ kind: 'footer', slot: 'right' }));
    const input = field(grid);
    if (!input) throw new Error('editor did not open');
    input.value = 'Page &P';
    input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));

    expect(getPageSetup(store.getState(), 0).footerRight).toBe('Page &P');
    expect(field(grid)).toBeNull();
    expect(history.undo()).toBe(true);
    expect(getPageSetup(store.getState(), 0).footerRight).toBeUndefined();
    handle.detach();
  });

  it('discards the edit on Escape', () => {
    const { grid, store, handle } = setup();
    mutators.setPageSetup(store, 0, { headerLeft: 'keep' });

    grid.dispatchEvent(bandEvent({ slot: 'left' }));
    const input = field(grid);
    if (!input) throw new Error('editor did not open');
    input.value = 'discard';
    input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));

    expect(getPageSetup(store.getState(), 0).headerLeft).toBe('keep');
    expect(field(grid)).toBeNull();
    handle.detach();
  });

  it('drops the field entirely when the slot is emptied', () => {
    const { grid, store, handle } = setup();
    mutators.setPageSetup(store, 0, { headerCenter: 'title' });

    grid.dispatchEvent(bandEvent());
    const input = field(grid);
    if (!input) throw new Error('editor did not open');
    input.value = '';
    input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));

    expect(getPageSetup(store.getState(), 0).headerCenter).toBeUndefined();
    handle.detach();
  });

  it('stops listening once detached', () => {
    const { grid, handle } = setup();
    handle.detach();

    grid.dispatchEvent(bandEvent());
    expect(field(grid)).toBeNull();
  });
});
