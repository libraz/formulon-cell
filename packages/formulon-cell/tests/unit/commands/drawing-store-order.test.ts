import { AnchorEditAs, AnchorKind } from '@libraz/formulon';
import { afterEach, describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import {
  arrangeSessionChart,
  clearSessionChart,
  createSessionChart,
} from '../../../src/commands/session-chart.js';
import {
  arrangeSessionIllustration,
  arrangeStackPosition,
  clearSessionIllustration,
  createSessionImage,
  createSessionShape,
  type SessionIllustrationArrangeAction,
  sheetDrawingStack,
} from '../../../src/commands/session-illustration.js';
import { deleteRows, insertRows } from '../../../src/commands/structure.js';
import { hydrateChartsFromEngine } from '../../../src/engine/charts-sync.js';
import { hydrateImagesFromEngine } from '../../../src/engine/images-sync.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { reloadDrawingsAfterSheetEdit } from '../../../src/mount/engine-binding.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

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

const cell = (row: number, col: number) => ({ sheet: 0, r0: row, c0: col, r1: row, c1: col });

const bind = (wb: WorkbookHandle) => {
  const store = createSpreadsheetStore();
  hydrateImagesFromEngine(wb, store, 0);
  hydrateChartsFromEngine(wb, store, 0);
  wb.attachStore(store);
  const history = new History();
  wb.attachHistory(history);
  return { store, history };
};

const open$ = async (preferStub = false) => {
  const wb = await WorkbookHandle.createDefault(preferStub ? { preferStub: true } : undefined);
  open.push(wb);
  return { wb, ...bind(wb) };
};

const ids = (store: SpreadsheetStore): string[] =>
  sheetDrawingStack(store.getState(), 0).map((entry) => entry.id);

const addChart = (store: SpreadsheetStore, wb: WorkbookHandle): string => {
  wb.setNumber({ sheet: 0, row: 40, col: 0 }, 1);
  wb.setNumber({ sheet: 0, row: 41, col: 0 }, 2);
  const chart = createSessionChart(store, cell(40, 0), {
    wb,
    kind: 'column',
    anchor: { row: 44, col: 4 },
  });
  if (!chart) throw new Error('chart was not created');
  return chart.id;
};

const addPicture = (store: SpreadsheetStore, wb: WorkbookHandle, saved: boolean): string => {
  const item = createSessionImage(store, cell(30, 1), {
    id: 'P',
    src: 'data:image/png;base64,x',
    w: 24,
    h: 24,
    ...(saved ? { bytes: PNG, wb } : {}),
  });
  if (!item) throw new Error('picture was not created');
  expect(item.nativeObjectId !== undefined).toBe(saved);
  return item.id;
};

const addArrow = (store: SpreadsheetStore): string => {
  const item = createSessionShape(store, cell(0, 0), { id: 'A', shape: 'arrow' });
  if (!item) throw new Error('arrow was not created');
  return item.id;
};

/** Ids of the sheet's saved drawings in the engine's order, back to front. */
const engineIds = (store: SpreadsheetStore, wb: WorkbookHandle): string[] => {
  const state = store.getState();
  const byObject = new Map<number, string>();
  for (const chart of state.charts.charts) byObject.set(chart.nativeObjectId, chart.id);
  for (const item of state.illustrations.illustrations) {
    if (item.nativeObjectId !== undefined) byObject.set(item.nativeObjectId, item.id);
  }
  return wb.listDrawingObjects(0).flatMap((object) => byObject.get(object.objectId) ?? []);
};

/** Send the picture behind the other drawing, bring it back, then undo and redo each step. */
const runSequence = (
  store: SpreadsheetStore,
  history: History,
  wb: WorkbookHandle,
  other: string,
  picture: string,
  check: () => void = () => undefined,
): void => {
  const expectStack = (expected: string[]) => {
    expect(ids(store)).toEqual(expected);
    check();
  };
  expectStack([other, picture]);
  expect(arrangeSessionIllustration(store, picture, 'send-back', history, undefined, wb)).not.toBe(
    null,
  );
  expectStack([picture, other]);
  expect(
    arrangeSessionIllustration(store, picture, 'bring-front', history, undefined, wb),
  ).not.toBe(null);
  expectStack([other, picture]);
  expect(history.undo()).toBe(true);
  expectStack([picture, other]);
  expect(history.undo()).toBe(true);
  expectStack([other, picture]);
  expect(history.redo()).toBe(true);
  expectStack([picture, other]);
  expect(history.redo()).toBe(true);
  expectStack([other, picture]);
};

describe('one stored drawing order', () => {
  it('orders a saved picture against a chart, through undo, redo and a reload', async () => {
    const { wb, store, history } = await open$();
    const chart = addChart(store, wb);
    const picture = addPicture(store, wb, true);
    runSequence(store, history, wb, chart, picture, () =>
      expect(engineIds(store, wb)).toEqual(ids(store)),
    );
    expect(
      arrangeSessionIllustration(store, picture, 'send-back', history, undefined, wb),
    ).not.toBe(null);
    expect(ids(store)).toEqual([picture, chart]);

    const reopened = await WorkbookHandle.loadBytes(wb.save());
    open.push(reopened);
    const { store: reloaded } = bind(reopened);
    expect(sheetDrawingStack(reloaded.getState(), 0).map((entry) => entry.chart)).toEqual([
      false,
      true,
    ]);
  });

  it('orders a picture without bytes against a chart, through undo and redo', async () => {
    const { wb, store, history } = await open$();
    const chart = addChart(store, wb);
    const picture = addPicture(store, wb, false);
    runSequence(store, history, wb, chart, picture);
    // The chart moves past the session-only picture too, and back on undo.
    expect(arrangeSessionChart(store, chart, 'bring-front', history, wb)).not.toBeNull();
    expect(ids(store)).toEqual([picture, chart]);
    expect(history.undo()).toBe(true);
    expect(ids(store)).toEqual([chart, picture]);
  });

  it('orders session-only drawings on the stub engine, which holds no chart', async () => {
    const { wb, store, history } = await open$(true);
    expect(wb.capabilities.drawings).toBe(false);
    const shape = createSessionShape(store, cell(0, 0), { id: 'S', shape: 'rectangle' })?.id;
    if (!shape) throw new Error('shape was not created');
    const picture = addPicture(store, wb, false);
    runSequence(store, history, wb, shape, picture);
  });

  it('keeps a session-only shape behind a chart across a row insert', async () => {
    const { wb, store, history } = await open$();
    const chart = addChart(store, wb);
    const arrow = addArrow(store);
    expect(ids(store)).toEqual([chart, arrow]);
    expect(arrangeSessionIllustration(store, arrow, 'send-back', history, undefined, wb)).not.toBe(
      null,
    );
    expect(ids(store)).toEqual([arrow, chart]);
    expect(insertRows(store, wb, history, 0, 2)).toBe(true);
    expect(ids(store)).toEqual([arrow, chart]);
    expect(history.undo()).toBe(true);
    expect(ids(store)).toEqual([arrow, chart]);
  });

  it('keeps a session-only shape over a saved picture across a row insert', async () => {
    const { wb, store, history } = await open$();
    const picture = addPicture(store, wb, true);
    const arrow = addArrow(store);
    expect(ids(store)).toEqual([picture, arrow]);
    expect(insertRows(store, wb, history, 0, 2)).toBe(true);
    expect(ids(store)).toEqual([picture, arrow]);
  });

  it('enables each arrange command exactly when the command moves the drawing', async () => {
    const { wb, store, history } = await open$();
    addPicture(store, wb, true);
    const chart = addChart(store, wb);
    const arrow = addArrow(store);
    const actions: [SessionIllustrationArrangeAction, 'atFront' | 'atBack'][] = [
      ['bring-forward', 'atFront'],
      ['bring-front', 'atFront'],
      ['send-backward', 'atBack'],
      ['send-back', 'atBack'],
    ];
    expect(arrangeStackPosition(sheetDrawingStack(store.getState(), 0), arrow).atBack).toBe(false);
    for (const id of ids(store)) {
      for (const [action, end] of actions) {
        const enabled = !arrangeStackPosition(sheetDrawingStack(store.getState(), 0), id)[end];
        const moved =
          id === chart
            ? arrangeSessionChart(store, id, action, history, wb) !== null
            : arrangeSessionIllustration(store, id, action, history, undefined, wb) !== null;
        expect(moved, `${id} ${action}`).toBe(enabled);
        if (moved) expect(history.undo()).toBe(true);
      }
    }
  });
});

/** Move sheet 0 to the end and show it, as the mounted sheet tabs do. */
const moveFirstSheetLast = (store: SpreadsheetStore, wb: WorkbookHandle): number => {
  const last = wb.sheetCount - 1;
  const unsubscribe = wb.subscribe((event) => {
    if (event.kind === 'sheet-move') reloadDrawingsAfterSheetEdit(wb, store, event);
  });
  expect(wb.moveSheet(0, last)).toBe(true);
  unsubscribe();
  mutators.setSheetIndex(store, last);
  hydrateImagesFromEngine(wb, store, last);
  hydrateChartsFromEngine(wb, store, last);
  return last;
};

describe('saved drawing ids', () => {
  it('keeps a session-only shape over a saved picture across a sheet move', async () => {
    const { wb, store } = await open$();
    wb.addSheet('Second');
    const picture = addPicture(store, wb, true);
    const arrow = addArrow(store);
    const sheet = moveFirstSheetLast(store, wb);
    expect(sheetDrawingStack(store.getState(), sheet).map((entry) => entry.id)).toEqual([
      picture,
      arrow,
    ]);
  });

  it('keeps a session-only shape over a saved picture across the undo of a row delete', async () => {
    const { wb, store, history } = await open$();
    const picture = addPicture(store, wb, true);
    const objectId = store.getState().illustrations.illustrations[0]?.nativeObjectId ?? -1;
    // A picture that moves and resizes with its cells goes with the rows under it.
    expect(
      wb.setDrawingAnchor(0, objectId, {
        kind: AnchorKind.TwoCell,
        editAs: AnchorEditAs.TwoCell,
        fromRow: 30,
        fromCol: 1,
        toRow: 31,
        toCol: 2,
      }),
    ).toBe(true);
    hydrateImagesFromEngine(wb, store, 0);
    expect(ids(store)).toEqual([picture]);
    const arrow = addArrow(store);
    expect(deleteRows(store, wb, history, 30, 3)).toBe(true);
    expect(ids(store)).toEqual([arrow]);
    expect(history.undo()).toBe(true);
    expect(ids(store)).toEqual([picture, arrow]);
  });

  it("keeps a saved chart's id across a sheet move", async () => {
    const { wb, store, history } = await open$();
    wb.addSheet('Second');
    const chart = addChart(store, wb);
    const sheet = moveFirstSheetLast(store, wb);
    const stack = () => sheetDrawingStack(store.getState(), sheet).map((entry) => entry.id);
    expect(stack()).toEqual([chart]);
    // A delete undone on the new sheet brings the chart back under the same id.
    expect(clearSessionChart(store, chart, history, wb)).toBe(true);
    expect(history.undo()).toBe(true);
    expect(stack()).toEqual([chart]);
  });

  it('does not grow the drawing order over insert and delete cycles', async () => {
    const { wb, store, history } = await open$();
    for (let cycle = 0; cycle < 12; cycle += 1) {
      const chart = addChart(store, wb);
      expect(clearSessionChart(store, chart, history, wb)).toBe(true);
      const picture = createSessionImage(store, cell(30, 1), {
        id: `P${cycle}`,
        src: 'data:image/png;base64,x',
        bytes: PNG,
        wb,
      });
      if (!picture) throw new Error('picture was not created');
      expect(clearSessionIllustration(store, picture.id, history, wb)).toBe(true);
      const shape = createSessionShape(store, cell(0, 0), { id: `S${cycle}`, shape: 'arrow' });
      if (!shape) throw new Error('shape was not created');
      expect(clearSessionIllustration(store, shape.id, history, wb)).toBe(true);
    }
    expect(store.getState().drawings.order).toEqual([]);
  });
});

describe('sheet edits reloading drawings', () => {
  it('keeps chart and picture ids when the reload listens before the store', async () => {
    const wb = await WorkbookHandle.createDefault();
    open.push(wb);
    let store: SpreadsheetStore | null = null;
    // Subscribed ahead of every store binding, as a host may be.
    wb.subscribe((event) => {
      if (store && event.kind === 'sheet-move') reloadDrawingsAfterSheetEdit(wb, store, event);
    });
    wb.addSheet('Second');
    const bound = bind(wb);
    store = bound.store;
    const chart = addChart(store, wb);
    const picture = addPicture(store, wb, true);
    const arrow = addArrow(store);
    expect(wb.moveSheet(0, 1)).toBe(true);
    mutators.setSheetIndex(store, 1);
    hydrateImagesFromEngine(wb, store, 1);
    hydrateChartsFromEngine(wb, store, 1);
    expect(sheetDrawingStack(store.getState(), 1).map((entry) => entry.id)).toEqual([
      chart,
      picture,
      arrow,
    ]);
  });
});
