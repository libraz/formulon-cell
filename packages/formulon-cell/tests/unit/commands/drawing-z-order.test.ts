import { AnchorEditAs, AnchorKind } from '@libraz/formulon';
import { afterEach, describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { arrangeSessionChart, createSessionChart } from '../../../src/commands/session-chart.js';
import {
  arrangeSessionIllustration,
  arrangeStackPosition,
  clearSessionIllustration,
  createSessionImage,
  createSessionShape,
  type SessionIllustrationArrangeAction,
  sheetDrawingStack,
} from '../../../src/commands/session-illustration.js';
import { deleteRows } from '../../../src/commands/structure.js';
import { hydrateChartsFromEngine } from '../../../src/engine/charts-sync.js';
import { hydrateImagesFromEngine } from '../../../src/engine/images-sync.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, type SpreadsheetStore } from '../../../src/store/store.js';

const open: WorkbookHandle[] = [];
afterEach(() => {
  for (const wb of open.splice(0)) wb.dispose();
});

const PNG = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=',
  ),
  (char) => char.charCodeAt(0),
);

/** A two-cell rectangle named `name` over rows `fromRow`..`fromRow + 2`. */
const seedShape = (wb: WorkbookHandle, name: string, fromRow: number): void => {
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
};

/** Hydrate `wb` into a bound store with history. */
const bind = (wb: WorkbookHandle) => {
  const store = createSpreadsheetStore();
  hydrateImagesFromEngine(wb, store, 0);
  hydrateChartsFromEngine(wb, store, 0);
  wb.attachStore(store);
  const history = new History();
  wb.attachHistory(history);
  return { store, history };
};

/**
 * Shapes S1, S2, S3, then a picture I and a chart C on top, back to front.
 * S2 alone lies in rows 10-12.
 */
async function sheet() {
  const wb = await WorkbookHandle.createDefault();
  open.push(wb);
  seedShape(wb, 'S1', 0);
  seedShape(wb, 'S2', 10);
  seedShape(wb, 'S3', 20);
  const { store, history } = bind(wb);
  createSessionImage(
    store,
    { sheet: 0, r0: 30, c0: 1, r1: 30, c1: 1 },
    { id: 'I', src: 'data:image/png;base64,x', bytes: PNG, wb, w: 24, h: 24 },
  );
  wb.setNumber({ sheet: 0, row: 40, col: 0 }, 1);
  wb.setNumber({ sheet: 0, row: 41, col: 0 }, 2);
  const chart = createSessionChart(
    store,
    { sheet: 0, r0: 40, c0: 0, r1: 41, c1: 0 },
    { wb, kind: 'column', anchor: { row: 44, col: 4 } },
  );
  if (!chart) throw new Error('chart was not created');
  const labels = new Map<number, string>();
  for (const shape of wb.listShapes(0)) labels.set(shape.objectId, shape.name);
  const image = store.getState().illustrations.illustrations.find((item) => item.id === 'I');
  if (image?.nativeObjectId === undefined) throw new Error('picture was not saved');
  labels.set(image.nativeObjectId, 'I');
  labels.set(chart.nativeObjectId, 'C');
  const order = (): string[] =>
    wb.listDrawingObjects(0).map((object) => labels.get(object.objectId) ?? '?');
  return { wb, store, history, chart, labels, order };
}

const shapeId = (store: SpreadsheetStore, name: string, wb: WorkbookHandle): string => {
  const objectId = wb.listShapes(0).find((shape) => shape.name === name)?.objectId;
  const item = store
    .getState()
    .illustrations.illustrations.find((candidate) => candidate.nativeObjectId === objectId);
  if (!item) throw new Error(`missing shape ${name}`);
  return item.id;
};

describe('one stacking order across drawing kinds', () => {
  it('starts in the order the objects were added', async () => {
    const { order } = await sheet();
    expect(order()).toEqual(['S1', 'S2', 'S3', 'I', 'C']);
  });

  it('brings a shape forward across a picture', async () => {
    const { wb, store, history, order } = await sheet();
    const s3 = shapeId(store, 'S3', wb);
    expect(arrangeSessionIllustration(store, s3, 'bring-forward', history, undefined, wb)).not.toBe(
      null,
    );
    expect(order()).toEqual(['S1', 'S2', 'I', 'S3', 'C']);
    expect(history.undo()).toBe(true);
    expect(order()).toEqual(['S1', 'S2', 'S3', 'I', 'C']);
  });

  it('sends a chart to the back, under the shapes', async () => {
    const { wb, store, history, chart, order } = await sheet();
    expect(arrangeSessionChart(store, chart.id, 'send-back', history, wb)).not.toBeNull();
    expect(order()).toEqual(['C', 'S1', 'S2', 'S3', 'I']);
    expect(history.undo()).toBe(true);
    expect(order()).toEqual(['S1', 'S2', 'S3', 'I', 'C']);
  });

  it('lists the store objects in the engine order, back to front', async () => {
    const { wb, store, history, chart } = await sheet();
    const s1 = shapeId(store, 'S1', wb);
    arrangeSessionChart(store, chart.id, 'send-backward', history, wb);
    expect(sheetDrawingStack(store.getState(), 0).map((entry) => entry.id)).toEqual([
      s1,
      shapeId(store, 'S2', wb),
      shapeId(store, 'S3', wb),
      chart.id,
      'I',
    ]);
  });

  it('follows the measured sequence, undoes it, and saves it', async () => {
    const { wb, store, history, order } = await sheet();
    const arrange = (name: string, action: SessionIllustrationArrangeAction) =>
      arrangeSessionIllustration(store, shapeId(store, name, wb), action, history, undefined, wb);
    const shapes = (): string[] => order().filter((label) => label.startsWith('S'));
    const steps: [string, SessionIllustrationArrangeAction, string[]][] = [
      ['S1', 'bring-forward', ['S2', 'S1', 'S3']],
      ['S1', 'bring-front', ['S2', 'S3', 'S1']],
      ['S3', 'send-back', ['S3', 'S2', 'S1']],
      ['S2', 'send-backward', ['S2', 'S3', 'S1']],
    ];
    const seen = [order()];
    for (const [name, action, expected] of steps) {
      expect(arrange(name, action), `${name} ${action}`).not.toBeNull();
      expect(shapes(), `${name} ${action}`).toEqual(expected);
      seen.push(order());
    }
    expect(order()).toEqual(['S2', 'S3', 'I', 'C', 'S1']);

    for (let step = steps.length - 1; step >= 0; step -= 1) {
      expect(history.undo()).toBe(true);
      expect(order()).toEqual(seen[step]);
    }
    for (let step = 1; step <= steps.length; step += 1) {
      expect(history.redo()).toBe(true);
      expect(order()).toEqual(seen[step]);
    }

    const reopened = await WorkbookHandle.loadBytes(wb.save());
    open.push(reopened);
    const kinds = (handle: WorkbookHandle) =>
      handle.listDrawingObjects(0).map((object) => object.kind);
    expect(kinds(reopened)).toEqual(kinds(wb));
    expect(reopened.listShapes(0).map((shape) => shape.name)).toEqual(['S2', 'S3', 'S1']);
    const { store: reloaded } = bind(reopened);
    expect(sheetDrawingStack(reloaded.getState(), 0)).toHaveLength(5);
  });

  it('reports the ends of the cross-kind stack for the arrange commands', async () => {
    const { wb, store, chart } = await sheet();
    const stack = sheetDrawingStack(store.getState(), 0);
    expect(arrangeStackPosition(stack, chart.id)).toMatchObject({ atFront: true, atBack: false });
    expect(arrangeStackPosition(stack, shapeId(store, 'S1', wb))).toMatchObject({
      atFront: false,
      atBack: true,
    });
  });
});

describe('restored drawings keep their slot', () => {
  it('puts back a shape a row delete took at its old slot', async () => {
    const { wb, store, history, order } = await sheet();
    expect(deleteRows(store, wb, history, 10, 3)).toBe(true);
    expect(order()).toEqual(['S1', 'S3', 'I', 'C']);
    expect(history.undo()).toBe(true);
    expect(order()).toEqual(['S1', 'S2', 'S3', 'I', 'C']);
  });

  it('puts back a deleted shape at its old slot', async () => {
    const { wb, store, history, order } = await sheet();
    expect(clearSessionIllustration(store, shapeId(store, 'S1', wb), history, wb)).toBe(true);
    expect(order()).toEqual(['S2', 'S3', 'I', 'C']);
    expect(history.undo()).toBe(true);
    expect(order()).toEqual(['S1', 'S2', 'S3', 'I', 'C']);
  });
});

describe('session-only objects in the stack', () => {
  it('keeps an arrow, which the file cannot hold, in its place among the drawings', async () => {
    const { wb, store, history, chart } = await sheet();
    createSessionShape(
      store,
      { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
      {
        id: 'arrow',
        shape: 'arrow',
      },
    );
    const ids = () => sheetDrawingStack(store.getState(), 0).map((entry) => entry.id);
    expect(ids().at(-1)).toBe('arrow');
    arrangeSessionChart(store, chart.id, 'send-back', history, wb);
    expect(ids().at(-1)).toBe('arrow');
    expect(
      arrangeSessionIllustration(store, 'arrow', 'send-back', history, undefined, wb),
    ).not.toBe(null);
    expect(ids()).toEqual([
      'arrow',
      chart.id,
      shapeId(store, 'S1', wb),
      shapeId(store, 'S2', wb),
      shapeId(store, 'S3', wb),
      'I',
    ]);
    expect(history.undo()).toBe(true);
    expect(ids().at(-1)).toBe('arrow');
  });
});
