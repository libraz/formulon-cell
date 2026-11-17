import { AnchorEditAs, AnchorKind } from '@libraz/formulon';
import { afterEach, describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import {
  clearSessionIllustration,
  createSessionShape,
  duplicateSessionIllustration,
  updateSessionIllustration,
} from '../../../src/commands/session-illustration.js';
import {
  applySessionShapeFill,
  applySessionShapeKind,
} from '../../../src/commands/shape-format.js';
import { hydrateImagesFromEngine } from '../../../src/engine/images-sync.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import {
  createSpreadsheetStore,
  type SessionIllustration,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

const open: WorkbookHandle[] = [];
afterEach(() => {
  for (const wb of open.splice(0)) wb.dispose();
});

const range = { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 };

const session = async (seed?: (wb: WorkbookHandle) => void, bytes?: Uint8Array) => {
  const wb = bytes ? await WorkbookHandle.loadBytes(bytes) : await WorkbookHandle.createDefault();
  open.push(wb);
  seed?.(wb);
  const store = createSpreadsheetStore();
  hydrateImagesFromEngine(wb, store, 0);
  wb.attachStore(store);
  const history = new History();
  wb.attachHistory(history);
  return { wb, store, history };
};

const only = (store: SpreadsheetStore): SessionIllustration => {
  const items = store.getState().illustrations.illustrations.filter((i) => i.kind === 'shape');
  expect(items).toHaveLength(1);
  return items[0] as SessionIllustration;
};

const adjOf = (wb: WorkbookHandle, index = 0) => wb.listShapes(0)[index]?.adjustValues;

const rounded = (store: SpreadsheetStore, history: History | null = null, cornerAdj?: number) =>
  createSessionShape(
    store,
    range,
    { id: 'r1', shape: 'rounded-rectangle', x: 120, y: 80, w: 200, h: 100, cornerAdj },
    history,
  );

/** A saved shape whose corner guide the file states. */
const loaded = (prst: string, adj: number) => (wb: WorkbookHandle) => {
  wb.insertShape(0, {
    presetGeometry: prst,
    anchor: {
      kind: AnchorKind.TwoCell,
      editAs: AnchorEditAs.TwoCell,
      fromRow: 4,
      fromCol: 1,
      toRow: 8,
      toCol: 4,
    },
    adjustValues: { adj },
    fill: { color: '#336699' },
  });
};

describe('corner of a rounded rectangle', () => {
  it('writes the corner guide as it is and reads it back after a save', async () => {
    const { wb, store } = await session();
    rounded(store, null, 12345);

    expect(adjOf(wb)).toEqual({ adj: 12345 });
    const reopened = await session(undefined, wb.save());
    expect(only(reopened.store).cornerAdj).toBe(12345);
  });

  it('states no corner for a shape without one', async () => {
    const { wb, store } = await session();
    rounded(store);

    expect(adjOf(wb)).toEqual({});
    expect(only(store).cornerAdj).toBeUndefined();
  });

  it('changes and clears the corner, and undoes both', async () => {
    const { wb, store, history } = await session();
    rounded(store, history, 10000);
    updateSessionIllustration(store, 'r1', { cornerAdj: 30000 }, history, wb);
    expect(adjOf(wb)).toEqual({ adj: 30000 });

    updateSessionIllustration(store, 'r1', { cornerAdj: undefined }, history, wb);
    expect(adjOf(wb)).toEqual({});

    history.undo();
    expect(adjOf(wb)).toEqual({ adj: 30000 });
    history.undo();
    expect(adjOf(wb)).toEqual({ adj: 10000 });
  });

  it('leaves the corner guide alone when the shape is resized', async () => {
    const { wb, store, history } = await session();
    rounded(store, history, 20000);

    updateSessionIllustration(store, 'r1', { w: 400, h: 200 }, history, wb);

    expect(adjOf(wb)).toEqual({ adj: 20000 });
    expect(only(store).cornerAdj).toBe(20000);
  });

  it('writes the corner a shape gets when it becomes a rounded rectangle', async () => {
    const { wb, store, history } = await session();
    createSessionShape(
      store,
      range,
      { id: 'r1', shape: 'rectangle', x: 120, y: 80, w: 200, h: 100 },
      history,
    );
    applySessionShapeKind(store, only(store), 'rounded-rectangle', history, wb);
    expect(adjOf(wb)).toEqual({});

    updateSessionIllustration(store, 'r1', { cornerAdj: 15000 }, history, wb);
    expect(adjOf(wb)).toEqual({ adj: 15000 });
  });

  it('puts the corner back when the shape is deleted and the delete is undone', async () => {
    const { wb, store, history } = await session();
    rounded(store, history, 25000);

    clearSessionIllustration(store, 'r1', history, wb);
    expect(wb.listShapes(0)).toHaveLength(0);
    history.undo();

    expect(adjOf(wb)).toEqual({ adj: 25000 });
  });

  it('gives a copy the corner of the source', async () => {
    const { wb, store, history } = await session();
    rounded(store, history, 25000);

    duplicateSessionIllustration(store, 'r1', history, wb);

    expect(adjOf(wb, 0)).toEqual({ adj: 25000 });
    expect(adjOf(wb, 1)).toEqual({ adj: 25000 });
  });

  it('keeps no corner on a shape that cannot hold one', async () => {
    const { store, history, wb } = await session();
    createSessionShape(
      store,
      range,
      { id: 'o1', shape: 'oval', x: 120, y: 80, w: 200, h: 100 },
      history,
    );

    updateSessionIllustration(store, 'o1', { cornerAdj: 20000 }, history, wb);

    expect(only(store).cornerAdj).toBeUndefined();
    expect(adjOf(wb)).toEqual({});
  });
});

describe('a corner the file states', () => {
  it('is read as it is', async () => {
    const { store } = await session(loaded('roundRect', 30000));
    expect(only(store).cornerAdj).toBe(30000);
  });

  it('is kept as it is through an edit that does not touch the corner', async () => {
    const { wb, store, history } = await session(loaded('roundRect', 12345));

    applySessionShapeFill(store, only(store), '#00ff00', history, wb);

    expect(adjOf(wb)).toEqual({ adj: 12345 });
  });

  it('is kept through a copy', async () => {
    const { wb, store, history } = await session(loaded('roundRect', 12345));

    duplicateSessionIllustration(store, only(store).id, history, wb);

    expect(adjOf(wb, 1)).toEqual({ adj: 12345 });
  });

  it('is changed by a new radius', async () => {
    const { wb, store, history } = await session(loaded('roundRect', 12345));
    const item = only(store);

    updateSessionIllustration(store, item.id, { cornerAdj: 0 }, history, wb);

    expect(adjOf(wb)).toEqual({ adj: 0 });
  });

  it('leaves the guides of another preset alone', async () => {
    const { wb, store, history } = await session(loaded('triangle', 25000));

    const item = only(store);
    expect(item.cornerAdj).toBeUndefined();
    applySessionShapeFill(store, item, '#00ff00', history, wb);
    duplicateSessionIllustration(store, item.id, history, wb);

    expect(adjOf(wb, 0)).toEqual({ adj: 25000 });
    expect(adjOf(wb, 1)).toEqual({ adj: 25000 });
  });
});
