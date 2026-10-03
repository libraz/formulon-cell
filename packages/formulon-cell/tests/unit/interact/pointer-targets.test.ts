import { describe, expect, it, vi } from 'vitest';
import type { PageBreakHandle } from '../../../src/render/grid/page-view.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

const handles = vi.hoisted(() => ({ list: [] as PageBreakHandle[] }));

vi.mock('../../../src/render/grid/page-view.js', async () => {
  const actual = await vi.importActual<typeof import('../../../src/render/grid/page-view.js')>(
    '../../../src/render/grid/page-view.js',
  );
  return { ...actual, getPageBreakHandles: () => handles.list };
});

const { pageBreakHandleAt, updateCursor } = await import(
  '../../../src/interact/pointer-targets.js'
);

const previewStore = () => {
  const store = createSpreadsheetStore();
  store.setState((s) => ({ ...s, ui: { ...s.ui, workbookView: 'pageBreakPreview' } }));
  return store;
};

describe('updateCursor in Page Break Preview', () => {
  it('follows the handle a drag would grab when a manual break overlaps a print-area edge', () => {
    // The print-area row edge is listed first; the manual column break sits within grab
    // range of the same point and is the one pageBreakHandleAt prefers.
    handles.list = [
      { axis: 'row', kind: 'printArea', index: 9, position: 100, manual: false },
      { axis: 'col', kind: 'break', index: 3, position: 102, manual: true },
    ];
    const store = previewStore();
    const host = document.createElement('div');

    expect(pageBreakHandleAt(store.getState(), 101, 101)?.kind).toBe('break');
    updateCursor(host, store, 101, 101);

    expect(host.style.cursor).toBe('col-resize');
  });

  it('uses the row-resize cursor over a row boundary', () => {
    handles.list = [{ axis: 'row', kind: 'break', index: 5, position: 200, manual: false }];
    const store = previewStore();
    const host = document.createElement('div');

    updateCursor(host, store, 300, 201);
    expect(host.style.cursor).toBe('row-resize');
  });
});
