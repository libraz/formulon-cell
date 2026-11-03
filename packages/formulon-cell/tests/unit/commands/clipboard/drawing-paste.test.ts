import { beforeEach, describe, expect, it, vi } from 'vitest';
import { createDrawingClipboard } from '../../../../src/commands/clipboard/drawing.js';
import * as sessionChart from '../../../../src/commands/session-chart.js';
import { createSpreadsheetStore, mutators } from '../../../../src/store/store.js';

vi.mock('../../../../src/commands/session-chart.js', () => ({
  sessionChartById: vi.fn(() => ({ id: 'c1' })),
  copySessionChart: vi.fn(() => ({ sheet: 0, drawing: new Uint8Array(), anchor: {} })),
  clearSessionChart: vi.fn(() => true),
  pasteSessionChart: vi.fn(() => ({ id: 'c2' })),
}));

describe('drawing clipboard paste placement', () => {
  const pasteChart = vi.mocked(sessionChart.pasteSessionChart);
  beforeEach(() => pasteChart.mockClear());

  const setup = () => {
    const store = createSpreadsheetStore();
    mutators.setSelectedObject(store, { id: 'c1', name: 'Chart 1' });
    const clipboard = createDrawingClipboard({ store, wb: {} as never, history: null });
    clipboard.take('copy');
    return { store, clipboard };
  };

  it('anchors at the top-left cell of the selected range', () => {
    const { store, clipboard } = setup();
    mutators.setRange(store, { sheet: 0, r0: 4, c0: 2, r1: 5, c1: 3 });
    mutators.setActive(store, { sheet: 0, row: 5, col: 3 });
    mutators.setRange(store, { sheet: 0, r0: 4, c0: 2, r1: 5, c1: 3 });
    clipboard.paste();
    expect(pasteChart.mock.calls[0]?.[2]).toMatchObject({ row: 4, col: 2 });
  });

  it('offsets each repeat paste at the same anchor by 12 px on both axes', () => {
    const { store, clipboard } = setup();
    mutators.setActive(store, { sheet: 0, row: 19, col: 10 });
    clipboard.paste();
    clipboard.paste();
    clipboard.paste();
    const offsets = pasteChart.mock.calls.map((call) => call[6]);
    expect(offsets).toEqual([
      { x: 0, y: 0 },
      { x: 12, y: 12 },
      { x: 24, y: 24 },
    ]);
  });

  it('restarts the cascade at a different anchor', () => {
    const { store, clipboard } = setup();
    mutators.setActive(store, { sheet: 0, row: 19, col: 10 });
    clipboard.paste();
    clipboard.paste();
    mutators.setActive(store, { sheet: 0, row: 29, col: 5 });
    clipboard.paste();
    expect(pasteChart.mock.calls[2]?.[6]).toEqual({ x: 0, y: 0 });
  });
});
