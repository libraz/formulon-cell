import { afterEach, describe, expect, it, vi } from 'vitest';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createConditionalCache, evaluateConditional } from '../../../src/render/conditional.js';
import { buildColLayout, buildRowLayout, layoutForView } from '../../../src/render/geometry.js';
import { paintCells } from '../../../src/render/grid/cells.js';
import type { ConditionalRule } from '../../../src/store/store.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';
import type { ResolvedTheme } from '../../../src/theme/resolve.js';
import { dateSerial, seedNumber, withConditionalRules } from './conditional-fixtures.js';

const canLoadWasm = (): boolean => typeof WebAssembly !== 'undefined';

const todayRule: ConditionalRule = {
  kind: 'date-occurring',
  range: { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 },
  period: 'today',
  apply: { fill: '#00ff00' },
};

// Time-period rules judge by the day of the last recalculation, as the desktop app does: the
// overlay changes when the values would, on recalc, not at midnight on its own.
describe('date-occurring rules and the recalculation date', () => {
  afterEach(() => {
    vi.useRealTimers();
  });

  it('keep the day of the last recalc until the workbook recalculates', async () => {
    vi.useFakeTimers({ toFake: ['Date'] });
    vi.setSystemTime(new Date(2026, 6, 8, 12));
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    try {
      let s = createSpreadsheetStore().getState();
      s = seedNumber(s, 0, 0, dateSerial(2026, 7, 8));
      s = seedNumber(s, 1, 0, dateSerial(2026, 7, 9));
      s = withConditionalRules(s, [todayRule]);
      const cache = createConditionalCache();
      const fills = () => {
        const overlay = evaluateConditional(s, wb, '', null, cache);
        return [overlay.get('0:0:0')?.fill, overlay.get('0:1:0')?.fill];
      };
      expect(fills()).toEqual(['#00ff00', undefined]);
      vi.setSystemTime(new Date(2026, 6, 9, 12));
      expect(fills()).toEqual(['#00ff00', undefined]);
      wb.recalc();
      expect(fills()).toEqual([undefined, '#00ff00']);
    } finally {
      wb.dispose();
    }
  });

  it.skipIf(!canLoadWasm())('take a pinned clock as the recalc day', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      const native = (wb as unknown as { wb: { setPinnedNow(time: unknown): void } }).wb;
      native.setPinnedNow({ year: 2020, month: 1, day: 2, hour: 9, minute: 0, second: 0 });
      wb.recalc();
      expect(wb.recalcTodaySerial()).toBe(dateSerial(2020, 1, 2));
    } finally {
      wb.dispose();
    }
  });

  it.skipIf(!canLoadWasm())('hand the engine the same day for the rules it evaluates', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      const evaluate = vi.spyOn(wb, 'evaluateCfRange');
      const state = createSpreadsheetStore().getState();
      const view = { ...state, layout: layoutForView(state) };
      const ctx = new Proxy({ measureText: () => ({ width: 8 }) } as Record<string, unknown>, {
        get: (object, property: string) => (property in object ? object[property] : () => {}),
        set: () => true,
      }) as unknown as CanvasRenderingContext2D;
      paintCells(
        {
          ctx,
          wb,
          locale: 'en-US',
          getDisplay: undefined,
          sheetBackgroundImage: null,
          fontEffects: false,
          conditionalCache: createConditionalCache(),
        },
        view,
        { textCell: 13, fontUi: 'sans-serif' } as ResolvedTheme,
        buildColLayout(view.layout, view.viewport),
        buildRowLayout(view.layout, view.viewport),
      );
      expect(evaluate).toHaveBeenCalled();
      expect(evaluate.mock.calls[0]?.[5]).toBe(wb.recalcTodaySerial());
    } finally {
      wb.dispose();
    }
  });
});
