import { AnchorEditAs, AnchorKind, ShapePaint } from '@libraz/formulon';
import { afterEach, describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import {
  canDeleteIllustration,
  clearSessionIllustration,
  createSessionShape,
  duplicateSessionIllustration,
  sessionIllustrationById,
  updateSessionIllustration,
} from '../../../src/commands/session-illustration.js';
import {
  applySessionShapeFill,
  applySessionShapeOutline,
  applySessionShapeOutlineDash,
  applySessionShapeOutlineWeight,
} from '../../../src/commands/shape-format.js';
import { deleteRows, insertRows } from '../../../src/commands/structure.js';
import { hydrateImagesFromEngine } from '../../../src/engine/images-sync.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import {
  createSpreadsheetStore,
  mutators,
  type SessionIllustration,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

const open: WorkbookHandle[] = [];
afterEach(() => {
  for (const wb of open.splice(0)) wb.dispose();
});

const range = { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 };

/** A workbook with its store bound; `seed` writes the engine before hydration. */
async function session(seed?: (wb: WorkbookHandle) => void, bytes?: Uint8Array) {
  const wb = bytes ? await WorkbookHandle.loadBytes(bytes) : await WorkbookHandle.createDefault();
  open.push(wb);
  seed?.(wb);
  const store = createSpreadsheetStore();
  hydrateImagesFromEngine(wb, store, 0);
  wb.attachStore(store);
  const history = new History();
  wb.attachHistory(history);
  return { wb, store, history };
}

const shapesOf = (store: SpreadsheetStore): readonly SessionIllustration[] =>
  store.getState().illustrations.illustrations.filter((item) => item.kind === 'shape');

const only = (store: SpreadsheetStore): SessionIllustration => {
  const items = shapesOf(store);
  expect(items).toHaveLength(1);
  return items[0] as SessionIllustration;
};

/** What the file holds, read back from a fresh engine. */
const saved = async (wb: WorkbookHandle) => session(undefined, wb.save());

const oval = (store: SpreadsheetStore, history: History | null = null) =>
  createSessionShape(
    store,
    range,
    { id: 's1', shape: 'oval', x: 120, y: 80, w: 200, h: 100, color: '#ff0000', lineWidth: 2 },
    history,
  );

describe('shapes saved with the workbook', () => {
  it('writes a new shape and hydrates it from the saved file', async () => {
    const { wb, store } = await session();
    oval(store);

    const [native] = wb.listShapes(0);
    expect(wb.listShapes(0)).toHaveLength(1);
    expect(native).toMatchObject({
      prst: 'ellipse',
      fillKind: ShapePaint.Solid,
      fillRgb: 0xff0000,
      lineKind: ShapePaint.Solid,
      lineWidthEmu: 25400,
    });
    expect(only(store)).toMatchObject({
      nativeObjectId: native?.objectId,
      nativeGeometryEditable: true,
    });
    expect(only(store).nativeAnchor?.cx).toBeGreaterThan(0);

    const reopened = await saved(wb);
    expect(only(reopened.store)).toMatchObject({
      shape: 'oval',
      color: '#ff0000',
      lineWidth: 2,
      nativeObjectId: native?.objectId,
    });
  });

  it('saves every shape kind except the arrow, which has no line end to store', async () => {
    const { wb, store } = await session();
    const kinds = [
      'rectangle',
      'rounded-rectangle',
      'oval',
      'triangle',
      'diamond',
      'line',
    ] as const;
    for (const [index, shape] of kinds.entries()) {
      createSessionShape(store, range, {
        id: shape,
        shape,
        x: 40 + index * 30,
        y: 60,
        w: 90,
        h: 50,
      });
    }
    createSessionShape(store, range, { id: 'arrow', shape: 'arrow', x: 40, y: 200, w: 90, h: 50 });

    expect(wb.listShapes(0).map((shape) => shape.prst)).toEqual([
      'rect',
      'roundRect',
      'ellipse',
      'triangle',
      'diamond',
      'line',
    ]);
    const arrow = sessionIllustrationById(store.getState(), 'arrow');
    expect(arrow?.nativeObjectId).toBeUndefined();

    const reopened = await saved(wb);
    // A loaded line has no known direction, so it comes back as a locked box.
    expect(shapesOf(reopened.store).map((item) => item.shape)).toEqual([
      ...kinds.slice(0, 5),
      'rectangle',
    ]);
    expect(shapesOf(reopened.store).at(-1)?.nativeLocked).toBe(true);
  });

  it('restyles a shape: fill, no fill, outline colour and weight', async () => {
    const { wb, store, history } = await session();
    oval(store);
    const style = (patch: (item: SessionIllustration) => void): void => {
      patch(only(store));
    };

    style((item) => applySessionShapeFill(store, item, '#00ff00', history, wb));
    expect(wb.listShapes(0)[0]).toMatchObject({ fillKind: ShapePaint.Solid, fillRgb: 0x00ff00 });
    style((item) => applySessionShapeOutline(store, item, '#0000ff', history, wb));
    style((item) => applySessionShapeOutlineWeight(store, item, 4.5, history, wb));
    expect(wb.listShapes(0)[0]).toMatchObject({
      lineKind: ShapePaint.Solid,
      lineRgb: 0x0000ff,
      lineWidthEmu: 57150,
    });
    style((item) => applySessionShapeFill(store, item, null, history, wb));
    expect(wb.listShapes(0)[0]?.fillKind).toBe(ShapePaint.None);
    style((item) => applySessionShapeOutline(store, item, null, history, wb));
    expect(wb.listShapes(0)[0]?.lineKind).toBe(ShapePaint.None);

    const reopened = await saved(wb);
    expect(only(reopened.store)).toMatchObject({ opacity: 0, lineWidth: 0 });
  });

  it('resets radius, dash and partial opacity, which the file cannot hold', async () => {
    const { wb, store, history } = await session();
    oval(store);

    updateSessionIllustration(
      store,
      's1',
      { radius: 12, lineDash: 'dash', opacity: 0.4 },
      history,
      wb,
    );

    expect(only(store)).toMatchObject({ radius: undefined, lineDash: undefined, opacity: 1 });
    expect(wb.listShapes(0)[0]?.fillKind).toBe(ShapePaint.Solid);
    updateSessionIllustration(store, 's1', { opacity: 0 }, history, wb);
    expect(only(store).opacity).toBe(0);
    expect(wb.listShapes(0)[0]?.fillKind).toBe(ShapePaint.None);
  });

  it('refuses a dash style, which the file cannot hold', async () => {
    const { wb, store, history } = await session();
    oval(store);
    const before = only(store);

    expect(applySessionShapeOutlineDash(store, before, 'dash', history, wb)).toBeNull();
    expect(only(store).lineDash).toBeUndefined();
  });

  it('moves and resizes a shape through its anchor', async () => {
    const { wb, store, history } = await session();
    oval(store);
    const before = wb.listShapes(0)[0];

    updateSessionIllustration(store, 's1', { x: 300, y: 200, w: 120, h: 60 }, history, wb);

    const after = wb.listShapes(0)[0];
    expect(after?.objectId).toBe(before?.objectId);
    expect(after?.cx).toBeCloseTo(120 * 9525, -2);
    expect(after?.cy).toBeCloseTo(60 * 9525, -2);
    expect([after?.fromRow, after?.fromCol]).not.toEqual([before?.fromRow, before?.fromCol]);
    expect(only(store).nativeAnchor).toMatchObject({
      fromRow: after?.fromRow,
      fromCol: after?.fromCol,
      cx: after?.cx,
    });

    const reopened = await saved(wb);
    expect(only(reopened.store).nativeAnchor).toEqual(only(store).nativeAnchor);
  });

  it('writes the text of a shape and reads it back', async () => {
    const { wb, store, history } = await session();
    oval(store);

    updateSessionIllustration(store, 's1', { text: 'Total\nQ3' }, history, wb);

    expect(wb.listShapes(0)[0]?.text).toBe('Total\nQ3');
    const reopened = await saved(wb);
    expect(only(reopened.store).text).toBe('Total\nQ3');
  });

  it('deletes a shape from the file', async () => {
    const { wb, store, history } = await session();
    oval(store);

    expect(clearSessionIllustration(store, 's1', history, wb)).toBe(true);

    expect(wb.listShapes(0)).toEqual([]);
    expect(shapesOf((await saved(wb)).store)).toEqual([]);
  });

  it('copies a shape into a second engine shape', async () => {
    const { wb, store, history } = await session();
    oval(store);

    const copy = duplicateSessionIllustration(store, 's1', history, wb);

    expect(wb.listShapes(0)).toHaveLength(2);
    const [first, second] = wb.listShapes(0);
    expect(second?.objectId).not.toBe(first?.objectId);
    expect(second?.prst).toBe('ellipse');
    expect(sessionIllustrationById(store.getState(), copy?.id ?? '')?.nativeObjectId).toBe(
      second?.objectId,
    );
    expect(second?.fromCol).toBeGreaterThanOrEqual(first?.fromCol ?? 0);
  });

  it('undoes and redoes create, restyle, move and delete in the file', async () => {
    const { wb, store, history } = await session();
    oval(store, history);
    expect(wb.listShapes(0)).toHaveLength(1);

    applySessionShapeFill(store, only(store), '#00ff00', history, wb);
    updateSessionIllustration(store, 's1', { x: 400, y: 300 }, history, wb);
    const moved = wb.listShapes(0)[0];
    clearSessionIllustration(store, 's1', history, wb);
    expect(wb.listShapes(0)).toEqual([]);

    history.undo();
    expect(wb.listShapes(0)).toHaveLength(1);
    expect(wb.listShapes(0)[0]).toMatchObject({
      fromRow: moved?.fromRow,
      fromCol: moved?.fromCol,
      fillRgb: 0x00ff00,
    });
    expect(only(store).nativeObjectId).toBe(wb.listShapes(0)[0]?.objectId);

    history.undo();
    history.undo();
    expect(wb.listShapes(0)[0]).toMatchObject({ fillRgb: 0xff0000 });
    expect(wb.listShapes(0)[0]?.fromCol).not.toBe(moved?.fromCol);

    history.undo();
    expect(wb.listShapes(0)).toEqual([]);
    expect(shapesOf(store)).toEqual([]);

    history.redo();
    history.redo();
    history.redo();
    history.redo();
    expect(wb.listShapes(0)).toEqual([]);
    history.undo();
    expect(wb.listShapes(0)).toHaveLength(1);
    expect(shapesOf((await saved(wb)).store)).toHaveLength(1);
  });

  it('shifts a shape once when rows are inserted, and back on undo', async () => {
    const { wb, store, history } = await session();
    oval(store);
    const row = wb.listShapes(0)[0]?.fromRow ?? 0;

    expect(insertRows(store, wb, history, 0, 3)).toBe(true);

    expect(wb.listShapes(0)[0]?.fromRow).toBe(row + 3);
    expect(wb.listShapes(0)).toHaveLength(1);
    expect(only(store).nativeAnchor?.fromRow).toBe(row + 3);

    history.undo();
    expect(wb.listShapes(0)[0]?.fromRow).toBe(row);
    expect(only(store).nativeAnchor?.fromRow).toBe(row);
    history.redo();
    expect(wb.listShapes(0)[0]?.fromRow).toBe(row + 3);
  });

  it('keeps a shape in the file when an earlier sheet is removed', async () => {
    const { wb, store } = await session();
    const second = wb.addSheet();
    createSessionShape(
      store,
      { sheet: second, r0: 0, c0: 0, r1: 0, c1: 0 },
      { id: 's2', shape: 'rectangle', x: 100, y: 60, w: 120, h: 80 },
    );
    expect(wb.listShapes(second)).toHaveLength(1);

    expect(wb.removeSheet(0)).toBe(true);

    expect(wb.listShapes(0)).toHaveLength(1);
    expect(sessionIllustrationById(store.getState(), 's2')?.sheet).toBe(0);
  });

  it('puts a collapsed shape back where it was when a row delete is undone', async () => {
    const { wb, store, history } = await session();
    oval(store);
    const before = wb.listShapes(0)[0];

    expect(deleteRows(store, wb, history, (before?.fromRow ?? 0) - 1, 4)).toBe(true);
    history.undo();

    expect(wb.listShapes(0)).toHaveLength(1);
    expect(wb.listShapes(0)[0]).toMatchObject({
      fromRow: before?.fromRow,
      fromCol: before?.fromCol,
      cx: before?.cx,
      cy: before?.cy,
    });
    expect(only(store).nativeAnchor?.fromRow).toBe(before?.fromRow);
  });

  it('puts back a two-cell shape that a row delete took, and its style', async () => {
    const { wb, store, history } = await session((handle) => {
      handle.insertShape(0, {
        prst: 'rect',
        anchorKind: AnchorKind.TwoCell,
        editAs: AnchorEditAs.TwoCell,
        fromRow: 10,
        fromCol: 2,
        toRow: 12,
        toCol: 4,
        fillKind: ShapePaint.Solid,
        fillRgb: 0x336699,
        text: 'kept',
      });
    });
    expect(shapesOf(store)).toHaveLength(1);

    expect(deleteRows(store, wb, history, 9, 5)).toBe(true);
    expect(wb.listShapes(0)).toEqual([]);
    expect(shapesOf(store)).toEqual([]);

    history.undo();
    expect(wb.listShapes(0)).toHaveLength(1);
    expect(wb.listShapes(0)[0]).toMatchObject({
      fromRow: 10,
      toRow: 12,
      fillRgb: 0x336699,
      text: 'kept',
    });
    expect(shapesOf(store)).toHaveLength(1);
  });
});

describe('deleting and restoring saved shapes', () => {
  const callout = (wb: WorkbookHandle): void => {
    wb.insertShape(0, {
      name: 'Callout',
      prst: 'rect',
      anchorKind: AnchorKind.TwoCell,
      editAs: AnchorEditAs.TwoCell,
      fromRow: 10,
      fromCol: 2,
      toRow: 12,
      toCol: 4,
      textBox: true,
      text: 'note',
      fillKind: ShapePaint.Solid,
      fillRgb: 0x336699,
    });
  };

  it('restores a deleted shape from the engine description', async () => {
    const { wb, store, history } = await session(callout);
    const before = wb.listShapes(0)[0];

    expect(clearSessionIllustration(store, only(store).id, history, wb)).toBe(true);
    expect(wb.listShapes(0)).toEqual([]);
    history.undo();

    expect(wb.listShapes(0)).toHaveLength(1);
    expect(wb.listShapes(0)[0]).toMatchObject({
      name: 'Callout',
      textBox: true,
      text: 'note',
      anchorKind: AnchorKind.TwoCell,
      editAs: AnchorEditAs.TwoCell,
      fromRow: before?.fromRow,
      toRow: before?.toRow,
      toCol: before?.toCol,
    });
  });

  it('puts a shape back at its anchor whatever the scroll position', async () => {
    const { wb, store, history } = await session();
    oval(store);
    const before = wb.listShapes(0)[0];

    clearSessionIllustration(store, 's1', history, wb);
    store.setState((state) => ({
      ...state,
      viewport: { ...state.viewport, rowStart: 7, colStart: 3 },
    }));
    history.undo();

    expect(wb.listShapes(0)[0]).toMatchObject({
      fromRow: before?.fromRow,
      fromCol: before?.fromCol,
      fromRowOff: before?.fromRowOff,
      cx: before?.cx,
    });
  });

  it('keeps the overlay id of a shape across rehydration, and its selection', async () => {
    const { wb, store, history } = await session();
    oval(store);
    mutators.setSelectedObject(store, { id: 's1', name: '' });

    expect(insertRows(store, wb, history, 0, 2)).toBe(true);

    expect(only(store).id).toBe('s1');
    expect(store.getState().ui.selectedObject?.id).toBe('s1');
  });
});

describe('loaded shapes the overlay cannot draw', () => {
  const star = (wb: WorkbookHandle): void => {
    wb.insertShape(0, {
      prst: 'star5',
      anchorKind: AnchorKind.OneCell,
      editAs: AnchorEditAs.OneCell,
      fromRow: 10,
      fromCol: 2,
      cx: 952500,
      cy: 952500,
      fillKind: ShapePaint.Solid,
      fillRgb: 0xffcc00,
    });
  };

  it('shows an unknown preset as a locked bounding box', async () => {
    const { store } = await session(star);

    expect(only(store)).toMatchObject({ nativeLocked: true, lineDash: 'dash', opacity: 0 });
  });

  it('treats a loaded line as a locked box, since its direction is unknown', async () => {
    const { store } = await session((wb) => {
      wb.insertShape(0, {
        prst: 'straightConnector1',
        anchorKind: AnchorKind.TwoCell,
        editAs: AnchorEditAs.TwoCell,
        fromRow: 3,
        fromCol: 1,
        toRow: 6,
        toCol: 4,
        lineKind: ShapePaint.Solid,
        lineRgb: 0x00ff00,
        lineWidthEmu: 19050,
      });
    });

    expect(only(store)).toMatchObject({ nativeLocked: true, shape: 'rectangle' });
  });

  it('cannot be deleted, so there is no undo to get wrong', async () => {
    const { wb, store, history } = await session(star);

    expect(clearSessionIllustration(store, only(store).id, history, wb)).toBe(false);
    expect(wb.listShapes(0)).toHaveLength(1);
    expect(canDeleteIllustration(only(store))).toBe(false);
  });

  it('moves it without touching its preset or fill, and refuses restyling', async () => {
    const { wb, store, history } = await session(star);
    const item = only(store);

    expect(applySessionShapeFill(store, item, '#000000', history, wb)).toBeNull();
    updateSessionIllustration(store, item.id, { x: 500, y: 400 }, history, wb);

    expect(wb.listShapes(0)[0]).toMatchObject({
      prst: 'star5',
      fillKind: ShapePaint.Solid,
      fillRgb: 0xffcc00,
    });
    expect(wb.listShapes(0)[0]?.fromRow).not.toBe(10);
  });

  it('cannot be copied, and leaves hidden shapes alone', async () => {
    const { wb, store, history } = await session((handle) => {
      star(handle);
      handle.insertShape(0, {
        anchorKind: AnchorKind.OneCell,
        editAs: AnchorEditAs.OneCell,
        fromRow: 1,
        fromCol: 1,
        cx: 100000,
        cy: 100000,
        hidden: true,
      });
    });

    expect(shapesOf(store)).toHaveLength(1);
    expect(duplicateSessionIllustration(store, only(store).id, history, wb)).toBeNull();
    expect(clearSessionIllustration(store, only(store).id, history, wb)).toBe(false);
    expect(wb.listShapes(0)).toHaveLength(2);
    expect(wb.listShapes(0)[1]?.hidden).toBe(true);
  });
});
