import { readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { AnchorEditAs, AnchorKind } from '@libraz/formulon';
import { afterEach, describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { arrangeSessionChart, createSessionChart } from '../../../src/commands/session-chart.js';
import {
  arrangeSessionIllustration,
  createSessionImage,
} from '../../../src/commands/session-illustration.js';
import { hydrateImagesFromEngine } from '../../../src/engine/images-sync.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { attachSessionCharts } from '../../../src/interact/session-charts.js';
import { attachSessionIllustrations } from '../../../src/interact/session-illustrations.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

/** `--fc-z-grid`, where the find bar, quick analysis and slicers sit. */
const GRID_TIER = 16_777_010;

/** `--fc-z-drawing`, where the shared drawing root sits. */
const DRAWING_TIER = 16_777_004;

const styleRoot = resolve(dirname(fileURLToPath(import.meta.url)), '../../../src/styles/core');

/** The cell editor's z-index in the grid's stacking context. */
const EDITOR_Z = Number(
  /\.fc-host__editor \{[^}]*?z-index: (\d+);/.exec(
    readFileSync(resolve(styleRoot, 'surface/grid-editor.css'), 'utf8'),
  )?.[1],
);

/** Classes the stylesheet gives a stacking context, with their z-index; pinned
 *  by the overlay z-index test. */
const CONTEXT_Z: Readonly<Record<string, number>> = { 'fc-drawings': DRAWING_TIER };

/** The z-index `panel` competes with among `parent`'s other children: that of
 *  the outermost ancestor below `parent` forming a stacking context, else its own. */
const competingZ = (panel: HTMLElement, parent: HTMLElement): number => {
  let z = Number(panel.style.zIndex);
  for (let el = panel.parentElement; el && el !== parent; el = el.parentElement) {
    const own =
      el.style.zIndex !== ''
        ? Number(el.style.zIndex)
        : [...el.classList].map((name) => CONTEXT_Z[name]).find((value) => value !== undefined);
    if (own !== undefined) z = own;
  }
  return z;
};

const cleanup: (() => void)[] = [];
afterEach(() => {
  for (const dispose of cleanup.splice(0).reverse()) dispose();
});

/** Two shapes under a chart, both overlays drawn into one layer. */
async function drawn() {
  const wb = await WorkbookHandle.createDefault();
  cleanup.push(() => wb.dispose());
  for (const [name, fromRow] of [
    ['S1', 0],
    ['S2', 10],
  ] as const) {
    wb.insertShape(0, {
      name,
      presetGeometry: 'rect',
      anchor: {
        kind: AnchorKind.TwoCell,
        editAs: AnchorEditAs.TwoCell,
        fromRow,
        fromCol: 1,
        toRow: fromRow + 2,
        toCol: 3,
      },
    });
  }
  const store = createSpreadsheetStore();
  hydrateImagesFromEngine(wb, store, 0);
  wb.attachStore(store);
  const history = new History();
  wb.attachHistory(history);
  wb.setNumber({ sheet: 0, row: 0, col: 0 }, 1);
  const chart = createSessionChart(
    store,
    { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
    { wb, kind: 'column', anchor: { row: 4, col: 4 } },
  );
  if (!chart) throw new Error('chart was not created');
  const host = document.createElement('div');
  document.body.appendChild(host);
  cleanup.push(() => host.remove());
  const charts = attachSessionCharts({ host, store, history, getWb: () => wb });
  const illustrations = attachSessionIllustrations({ host, store, history, getWb: () => wb });
  cleanup.push(
    () => charts.detach(),
    () => illustrations.detach(),
  );
  const shapeIds = store.getState().illustrations.illustrations.map((item) => item.id);
  /** Panel ids, back to front by z-index. */
  const painted = (): string[] =>
    Array.from(host.querySelectorAll<HTMLElement>('.fc-chart, .fc-illustration'))
      .map((panel) => ({
        id: panel.dataset.chartId ?? panel.dataset.illustrationId ?? '',
        z: Number(panel.style.zIndex),
      }))
      .sort((a, b) => a.z - b.z)
      .map((entry) => entry.id);
  const zOf = (id: string): number =>
    Number(
      host.querySelector<HTMLElement>(`[data-chart-id="${id}"], [data-illustration-id="${id}"]`)
        ?.style.zIndex,
    );
  return { wb, store, history, chart, host, shapeIds, painted, zOf };
}

describe('drawing panels share one stacking order', () => {
  it('stacks charts and shapes in the engine order', async () => {
    const { wb, store, history, chart, shapeIds, painted } = await drawn();
    expect(painted()).toEqual([...shapeIds, chart.id]);

    arrangeSessionChart(store, chart.id, 'send-back', history, wb);
    expect(painted()).toEqual([chart.id, ...shapeIds]);

    const [s1, s2] = shapeIds as [string, string];
    arrangeSessionIllustration(store, s1, 'bring-front', history, undefined, wb);
    expect(painted()).toEqual([chart.id, s2, s1]);
  });

  it('raises the selected panel over every kind', async () => {
    const { host, shapeIds, chart, zOf } = await drawn();
    const [s1, s2] = shapeIds as [string, string];
    host.querySelector<HTMLElement>(`[data-illustration-id="${s1}"]`)?.focus();
    expect(zOf(s1)).toBe(4);
    expect(zOf(chart.id)).toBe(3);
    expect(zOf(s2)).toBe(2);
  });

  it('keeps chart panels over the grid and under the in-grid panels', async () => {
    const { chart, zOf } = await drawn();
    expect(zOf(chart.id)).toBeGreaterThan(0);
    expect(zOf(chart.id)).toBeLessThan(GRID_TIER);
  });

  it('keeps a picture over the cell editor whatever the drawing count', () => {
    const over = (count: number): boolean => {
      const store = createSpreadsheetStore();
      for (let index = 0; index < count; index += 1) {
        createSessionImage(
          store,
          { sheet: 0, r0: 0, c0: index, r1: 0, c1: index },
          {
            id: `picture-${index}`,
            src: 'data:image/png;base64,x',
          },
        );
      }
      const host = document.createElement('div');
      document.body.appendChild(host);
      const handle = attachSessionIllustrations({ host, store });
      try {
        const front = host.querySelector<HTMLElement>(
          `[data-illustration-id="picture-${count - 1}"]`,
        );
        if (!front) throw new Error('missing picture panel');
        return competingZ(front, host) > EDITOR_Z;
      } finally {
        handle.detach();
        host.remove();
      }
    };
    expect(EDITOR_Z).toBeGreaterThan(0);
    expect(over(1)).toBe(true);
    expect(over(6)).toBe(true);
  });

  it('shares one drawing root until the last overlay leaves', async () => {
    const store = createSpreadsheetStore();
    const host = document.createElement('div');
    document.body.appendChild(host);
    cleanup.push(() => host.remove());
    const charts = attachSessionCharts({ host, store });
    const illustrations = attachSessionIllustrations({ host, store });
    const roots = host.querySelectorAll(':scope > .fc-drawings');
    expect(roots).toHaveLength(1);
    expect(host.querySelector('.fc-charts')?.parentElement).toBe(roots[0]);
    expect(host.querySelector('.fc-illustrations')?.parentElement).toBe(roots[0]);
    charts.detach();
    expect(host.querySelector('.fc-drawings')).toBe(roots[0]);
    illustrations.detach();
    expect(host.querySelector('.fc-drawings')).toBeNull();
  });
});
