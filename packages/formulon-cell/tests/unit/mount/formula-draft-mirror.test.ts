import { afterEach, describe, expect, it } from 'vitest';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { attachFormulaDraftMirror } from '../../../src/mount/formula-draft-mirror.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

describe('formula draft mirror typography', () => {
  afterEach(() => {
    document.body.innerHTML = '';
  });

  it('uses effective cell size and vertical alignment at the current zoom', async () => {
    const grid = document.createElement('div');
    Object.defineProperty(grid, 'clientWidth', { configurable: true, value: 500 });
    Object.defineProperty(grid, 'clientHeight', { configurable: true, value: 400 });
    document.body.append(grid);

    const store = createSpreadsheetStore();
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    const mirror = attachFormulaDraftMirror({ grid, store, wb: workbook });
    try {
      const anchor = { sheet: 0, row: 0, col: 0 };
      mutators.setCellFormat(store, anchor, {
        fontSize: 20,
        fontVertAlign: 'subscript',
        fontFamily: 'Georgia',
      });
      store.setState((state) => ({
        ...state,
        viewport: { ...state.viewport, zoom: 2 },
      }));
      mirror.project(anchor, '=A1');

      const element = grid.querySelector<HTMLElement>('.fc-host__formula-draft-mirror');
      expect(element?.hidden).toBe(false);
      expect(element?.style.fontSize).toBe('28px');
      expect(element?.style.fontFamily).toBe('Georgia, var(--fc-font-ui, system-ui, sans-serif)');
      expect(Number.parseFloat(element?.style.paddingTop ?? '')).toBeGreaterThan(4);
    } finally {
      mirror.detach();
      workbook.dispose();
    }
  });
});
