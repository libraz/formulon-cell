import {
  AnchorEditAs,
  AnchorKind,
  ChartLineDash,
  ChartLineEndSize,
  ChartLineEndType,
} from '@libraz/formulon';
import { afterEach, describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { hydrateImagesFromEngine } from '../../../src/engine/images-sync.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { attachSessionIllustrations } from '../../../src/interact/session-illustrations.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

const cleanup: (() => void)[] = [];
afterEach(() => {
  for (const dispose of cleanup.splice(0).reverse()) dispose();
});

type End = { type: ChartLineEndType; width?: ChartLineEndSize; length?: ChartLineEndSize };

/** A straight connector drawn into a host, with the ends and dash the file holds. */
async function drawn(line: { headEnd?: End; tailEnd?: End; dash?: ChartLineDash }) {
  const wb = await WorkbookHandle.createDefault();
  cleanup.push(() => wb.dispose());
  wb.insertShape(0, {
    presetGeometry: 'straightConnector1',
    anchor: {
      kind: AnchorKind.TwoCell,
      editAs: AnchorEditAs.TwoCell,
      fromRow: 2,
      fromCol: 1,
      toRow: 6,
      toCol: 4,
    },
    line: { color: '#ff0000', widthPt: 3, ...line },
  });
  const loaded = await WorkbookHandle.loadBytes(wb.save());
  cleanup.push(() => loaded.dispose());
  const store = createSpreadsheetStore();
  hydrateImagesFromEngine(loaded, store, 0);
  loaded.attachStore(store);
  const host = document.createElement('div');
  document.body.appendChild(host);
  cleanup.push(() => host.remove());
  const handle = attachSessionIllustrations({
    host,
    store,
    history: new History(),
    getWb: () => loaded,
  });
  cleanup.push(() => handle.detach());
  const svg = host.querySelector('.fc-illustration svg');
  const item = store.getState().illustrations.illustrations[0];
  if (!svg || !item) throw new Error('nothing drawn');
  return { svg, item };
}

describe('line ends of a loaded line', () => {
  it('draws a triangle at the end of an arrow, and cuts the line short of it', async () => {
    const { svg, item } = await drawn({ tailEnd: { type: ChartLineEndType.Triangle } });

    expect(item).toMatchObject({ shape: 'arrow' });
    expect(svg.querySelectorAll('polygon')).toHaveLength(1);
    const line = svg.querySelector('line');
    const box = svg.getAttribute('viewBox')?.split(' ').map(Number) ?? [];
    expect(Number(line?.getAttribute('x2'))).toBeLessThan(box[2] ?? 0);
  });

  it.each([
    ['stealth', ChartLineEndType.Stealth, 'polygon'],
    ['diamond', ChartLineEndType.Diamond, 'polygon'],
    ['oval', ChartLineEndType.Oval, 'ellipse'],
    ['arrow', ChartLineEndType.Arrow, 'polyline'],
  ] as const)('draws a %s end at either end of a line', async (name, type, tag) => {
    const { svg, item } = await drawn({ headEnd: { type }, tailEnd: { type } });

    expect(item).toMatchObject({ shape: 'line', headEnd: { type: name }, tailEnd: { type: name } });
    expect(svg.querySelectorAll(tag)).toHaveLength(2);
  });

  it('draws a larger end for a larger size', async () => {
    const extent = async (size: ChartLineEndSize) => {
      const { svg } = await drawn({
        tailEnd: { type: ChartLineEndType.Triangle, width: size, length: size },
      });
      const points = (svg.querySelector('polygon')?.getAttribute('points') ?? '')
        .split(' ')
        .map((pair) => pair.split(',').map(Number) as [number, number]);
      const xs = points.map(([x]) => x);
      return Math.max(...xs) - Math.min(...xs);
    };

    expect(await extent(ChartLineEndSize.Large)).toBeGreaterThan(
      await extent(ChartLineEndSize.Small),
    );
  });

  it('draws no end for a line that states none', async () => {
    const { svg, item } = await drawn({});

    expect(item.headEnd).toBeUndefined();
    expect(svg.querySelector('polygon, ellipse, polyline')).toBeNull();
  });
});

describe('dash presets of a loaded line', () => {
  it.each([
    [ChartLineDash.SysDot, '0 8'],
    [ChartLineDash.LgDashDotDot, '32 12 4 12 4 12'],
    [ChartLineDash.SysDashDot, '12 4 4 4'],
  ] as const)('draws dash %i with its own pattern', async (dash, pattern) => {
    const { svg } = await drawn({ dash });

    expect(svg.querySelector('line')?.getAttribute('stroke-dasharray')).toBe(pattern);
  });
});
