import { describe, expect, it } from 'vitest';
import { pageBreakHandleAt, updateCursor } from '../../../src/interact/pointer-targets.js';
import { gridHitState } from '../../../src/render/grid.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

const previewStore = () => {
  const store = createSpreadsheetStore();
  store.setState((s) => ({ ...s, ui: { ...s.ui, workbookView: 'pageBreakPreview' } }));
  return store;
};

describe('updateCursor in Page Break Preview', () => {
  it('follows the handle a drag would grab when a manual break overlaps a print-area edge', () => {
    // The print-area row edge is listed first; the manual column break sits within grab
    // range of the same point and is the one pageBreakHandleAt prefers.
    const store = previewStore();
    const host = document.createElement('div');
    const hits = gridHitState(host);
    hits.pageBreakHandles = [
      { axis: 'row', kind: 'printArea', index: 9, position: 100, manual: false },
      { axis: 'col', kind: 'break', index: 3, position: 102, manual: true },
    ];

    expect(pageBreakHandleAt(store.getState(), hits, 101, 101)?.kind).toBe('break');
    updateCursor(host, store, 101, 101);

    expect(host.style.cursor).toBe('col-resize');
  });

  it('uses the row-resize cursor over a row boundary', () => {
    const store = previewStore();
    const host = document.createElement('div');
    gridHitState(host).pageBreakHandles = [
      { axis: 'row', kind: 'break', index: 5, position: 200, manual: false },
    ];

    updateCursor(host, store, 300, 201);
    expect(host.style.cursor).toBe('row-resize');
  });
});
