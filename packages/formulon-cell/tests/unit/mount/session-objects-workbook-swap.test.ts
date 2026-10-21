import { describe, expect, it } from 'vitest';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { mutators } from '../../../src/store/store.js';
import { mountStubSheet } from '../../test-utils/mount.js';

describe('session objects across workbook replacement', () => {
  it('drops sparklines, slicers, watches, traces and error marks of the old workbook', async () => {
    const original = await WorkbookHandle.createDefault();
    const next = await WorkbookHandle.createDefault();
    const sheet = await mountStubSheet({ workbook: original });
    try {
      const { store } = sheet.instance;
      mutators.setSparkline(store, { sheet: 0, row: 0, col: 3 }, { kind: 'line', source: 'A1:C1' });
      mutators.addSlicer(store, { id: 's', tableName: 'T', column: 'X', selected: ['a'] });
      mutators.addWatch(store, { sheet: 0, row: 1, col: 1 });
      store.setState((s) => ({
        ...s,
        traces: {
          items: [
            {
              kind: 'precedent',
              from: { sheet: 0, row: 0, col: 0 },
              to: { sheet: 0, row: 1, col: 1 },
            },
          ],
        },
        errorIndicators: {
          ignoredErrors: new Set(['0:2:2']),
          validationCircles: new Set(['0:3:3']),
        },
      }));

      await sheet.instance.setWorkbook(next);

      const s = store.getState();
      expect(s.sparkline.sparklines.size).toBe(0);
      expect(s.slicers.slicers).toEqual([]);
      expect(s.watch.watches).toEqual([]);
      expect(s.traces.items).toEqual([]);
      expect(s.errorIndicators.ignoredErrors.size).toBe(0);
      expect(s.errorIndicators.validationCircles.size).toBe(0);
    } finally {
      sheet.dispose();
      original.dispose();
      next.dispose();
    }
  });
});
