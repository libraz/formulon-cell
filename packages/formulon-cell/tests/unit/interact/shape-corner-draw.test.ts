import { AnchorEditAs, AnchorKind } from '@libraz/formulon';
import { afterEach, describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { updateSessionIllustration } from '../../../src/commands/session-illustration.js';
import { hydrateImagesFromEngine } from '../../../src/engine/images-sync.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { attachSessionIllustrations } from '../../../src/interact/session-illustrations.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const cleanup: (() => void)[] = [];
afterEach(() => {
  for (const dispose of cleanup.splice(0).reverse()) dispose();
});

/** A workbook holding one rounded rectangle with corner guide `adj`, drawn into a host. */
async function drawn(adj: number) {
  const wb = await WorkbookHandle.createDefault();
  cleanup.push(() => wb.dispose());
  wb.insertShape(0, {
    presetGeometry: 'roundRect',
    adjustValues: { adj },
    anchor: {
      kind: AnchorKind.TwoCell,
      editAs: AnchorEditAs.TwoCell,
      fromRow: 2,
      fromCol: 1,
      toRow: 6,
      toCol: 4,
    },
  });
  const store = createSpreadsheetStore();
  hydrateImagesFromEngine(wb, store, 0);
  wb.attachStore(store);
  const history = new History();
  wb.attachHistory(history);
  const host = document.createElement('div');
  document.body.appendChild(host);
  cleanup.push(() => host.remove());
  const handle = attachSessionIllustrations({ host, store, history, getWb: () => wb });
  cleanup.push(() => handle.detach());
  const corner = () => {
    const rect = host.querySelector('svg rect');
    const panel = host.querySelector<HTMLElement>('.fc-illustration');
    return {
      rx: Number(rect?.getAttribute('rx')),
      short: Math.min(
        Number.parseFloat(panel?.style.width ?? '0'),
        Number.parseFloat(panel?.style.height ?? '0'),
      ),
    };
  };
  return { wb, store, history, corner };
}

describe('the drawn corner of a saved rounded rectangle', () => {
  it('is the corner guide share of the shorter side', async () => {
    const { corner } = await drawn(20000);
    const { rx, short } = corner();
    expect(short).toBeGreaterThan(0);
    expect(rx).toBeCloseTo(0.2 * short, 3);
  });

  it('scales with the shape when it is resized, the guide untouched', async () => {
    const { wb, store, history, corner } = await drawn(20000);
    const before = corner();
    const item = store.getState().illustrations.illustrations[0];
    if (!item) throw new Error('no item');

    updateSessionIllustration(
      store,
      item.id,
      { w: (item.w ?? 0) * 2, h: (item.h ?? 0) * 2 },
      history,
      wb,
    );

    expect(wb.listShapes(0)[0]?.adjustValues).toEqual({ adj: 20000 });
    const after = corner();
    expect(after.short).toBeCloseTo(before.short * 2, 1);
    expect(after.rx).toBeCloseTo(before.rx * 2, 1);
  });

  it('scales with the zoom', async () => {
    const { store, corner } = await drawn(20000);
    store.setState((state) => ({ ...state, viewport: { ...state.viewport, zoom: 2 } }));

    // The drawing is laid out unzoomed and scaled with its panel.
    const { rx, short } = corner();
    expect(rx * 2).toBeCloseTo(0.2 * short, 3);
  });
});

describe('a plain rectangle', () => {
  it('is drawn with square corners', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    cleanup.push(() => host.remove());
    const store = createSpreadsheetStore();
    mutators.upsertIllustration(store, {
      id: 'box',
      kind: 'shape',
      shape: 'rectangle',
      sheet: 0,
      x: 40,
      y: 40,
      w: 120,
      h: 60,
      cornerAdj: 30000,
    });
    const handle = attachSessionIllustrations({ host, store });
    cleanup.push(() => handle.detach());

    expect(host.querySelector('svg rect')?.getAttribute('rx')).toBe('0');
  });
});
