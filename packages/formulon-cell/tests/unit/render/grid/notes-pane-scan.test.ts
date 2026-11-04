import { describe, expect, it } from 'vitest';
import { buildColLayout, buildRowLayout, layoutForView } from '../../../../src/render/geometry.js';
import { paintVisibleNotes } from '../../../../src/render/grid/notes.js';
import { createSpreadsheetStore, mutators } from '../../../../src/store/store.js';

describe('paintVisibleNotes pane scan', () => {
  it('walks the visible rows once however many notes are shown', () => {
    const store = createSpreadsheetStore();
    for (let i = 0; i < 4; i += 1) {
      mutators.setCellFormat(
        store,
        { sheet: 0, row: 1 + i, col: 1 },
        { comment: 'n', commentVisible: true },
      );
    }
    const state = store.getState();
    const view = { ...state, layout: layoutForView(state) };
    const cols = buildColLayout(view.layout, view.viewport);
    const built = buildRowLayout(view.layout, view.viewport);
    let reads = 0;
    const rows = {
      ...built,
      get visible() {
        reads += 1;
        return built.visible;
      },
    };
    const ctx = new Proxy({} as Record<string, unknown>, {
      get: (t, k: string) => (k in t ? t[k] : () => ({ width: 0 })),
      set: (t, k: string, v) => {
        t[k] = v;
        return true;
      },
    }) as unknown as CanvasRenderingContext2D;

    paintVisibleNotes(ctx, view, cols, rows);

    expect(reads).toBe(1);
  });
});
