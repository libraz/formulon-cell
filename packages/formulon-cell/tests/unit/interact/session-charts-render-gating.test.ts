import { AnchorEditAs, AnchorKind } from '@libraz/formulon';
import { describe, expect, it } from 'vitest';
import { attachSessionCharts } from '../../../src/interact/session-charts.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const setup = () => {
  const host = document.createElement('div');
  document.body.appendChild(host);
  const store = createSpreadsheetStore();
  mutators.upsertChart(store, {
    id: 'chart-gate',
    sheet: 0,
    nativeObjectId: 1,
    nativeAnchor: {
      anchorKind: AnchorKind.TwoCell,
      editAs: AnchorEditAs.TwoCell,
      fromRow: 1,
      fromCol: 1,
      fromRowOff: 0,
      fromColOff: 0,
      toRow: 16,
      toCol: 8,
      toRowOff: 0,
      toColOff: 0,
      cx: 4_572_000,
      cy: 2_743_200,
    },
    kind: 'column',
    data: {
      domain: 'category',
      categories: ['A', 'B'],
      categoryIds: ['category-0', 'category-1'],
      series: [{ id: 'series-0', name: 'Value', values: [10, 20] }],
    },
  });
  const handle = attachSessionCharts({ host, store });
  const panel = (): Element | null => host.querySelector('.fc-chart');
  return { host, store, handle, panel };
};

describe('attachSessionCharts render gating', () => {
  it('keeps the panels when only the cursor selection changes', () => {
    const { store, handle, panel } = setup();
    const before = panel();
    expect(before).not.toBeNull();
    mutators.setActive(store, { sheet: 0, row: 4, col: 4 });
    expect(panel()).toBe(before);
    handle.detach();
  });

  it('rebuilds the panels when a chart changes', () => {
    const { store, handle, panel } = setup();
    const before = panel();
    mutators.updateChart(store, 'chart-gate', { title: 'Renamed' });
    expect(panel()).not.toBe(before);
    handle.detach();
  });

  it('rebuilds the panels when the viewport scrolls', () => {
    const { store, handle, panel } = setup();
    const before = panel();
    mutators.scrollBy(store, 3, 0);
    expect(panel()).not.toBe(before);
    handle.detach();
  });
});
