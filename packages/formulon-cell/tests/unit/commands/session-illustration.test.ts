import { describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { setProtectedSheet } from '../../../src/commands/protection.js';
import {
  arrangeSessionIllustration,
  clearSessionIllustration,
  createRibbonImageFromSelection,
  createRibbonShapeFromSelection,
  createSessionImage,
  createSessionShape,
  duplicateSessionIllustration,
  listSessionIllustrations,
  sessionIllustrationById,
  updateSessionIllustration,
} from '../../../src/commands/session-illustration.js';
import {
  drawingAnchorOf,
  EMU_PER_PX,
  sessionPosition,
} from '../../../src/engine/drawing-geometry.js';
import { hydrateImagesFromEngine } from '../../../src/engine/images-sync.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { cellRectUnclamped, layoutForView } from '../../../src/render/geometry.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

const bytesFromBase64 = (value: string): Uint8Array =>
  Uint8Array.from(atob(value), (char) => char.charCodeAt(0));

const PNG = bytesFromBase64(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=',
);

const range = (r0: number, c0: number, r1: number, c1: number) =>
  ({ sheet: 0, r0, c0, r1, c1 }) as const;

describe('session illustration commands', () => {
  it('creates reusable session shape and image overlays with stable defaults', () => {
    const store = createSpreadsheetStore();
    const r = range(2, 1, 2, 1);

    const shape = createSessionShape(store, r, { shape: 'oval' });
    const image = createSessionImage(store, r, { src: 'data:image/png;base64,a', alt: 'a.png' });

    expect(shape).toEqual({
      id: 'shape-0-2-1-oval',
      kind: 'shape',
      shape: 'oval',
      sheet: 0,
      x: undefined,
      y: undefined,
      w: undefined,
      h: undefined,
      color: undefined,
      radius: undefined,
      lineWidth: undefined,
      opacity: undefined,
    });
    expect(image).toEqual({
      id: 'image-0-2-1',
      kind: 'image',
      src: 'data:image/png;base64,a',
      alt: 'a.png',
      sheet: 0,
      x: undefined,
      y: undefined,
      w: undefined,
      h: undefined,
    });
    expect(listSessionIllustrations(store.getState())).toEqual([shape, image]);
  });

  it('creates ribbon images and shapes with Excel-like overlay placement defaults', () => {
    const store = createSpreadsheetStore();
    const r = range(0, 0, 0, 0);

    const shape = createRibbonShapeFromSelection(store, r, 'arrow');
    const image = createRibbonImageFromSelection(
      store,
      r,
      'data:image/png;base64,b',
      null,
      'b.png',
    );

    expect(shape).toMatchObject({
      id: 'ribbon-shape-0-0-0-arrow-0',
      kind: 'shape',
      shape: 'arrow',
      sheet: 0,
      x: 300,
      y: 88,
      w: 180,
      h: 80,
      color: '#0f6cbd',
    });
    expect(image).toMatchObject({
      id: 'ribbon-image-0-0-0-1',
      kind: 'image',
      src: 'data:image/png;base64,b',
      alt: 'b.png',
      sheet: 0,
      x: 324,
      y: 112,
      w: 240,
      h: 160,
    });
    // Each insert selects what it inserted.
    expect(store.getState().ui.selectedObject?.id).toBe(image?.id);
  });

  it('lists, finds, updates, and clears session illustrations for host object panes', () => {
    const store = createSpreadsheetStore();
    const shape = createSessionShape(store, range(0, 0, 0, 0), { id: 'shape-a', shape: 'line' });
    createSessionImage(store, range(1, 0, 1, 0), { id: 'image-a', src: 'https://example.test/a' });

    expect(listSessionIllustrations(store.getState()).map((item) => item.id)).toEqual([
      'shape-a',
      'image-a',
    ]);
    expect(sessionIllustrationById(store.getState(), 'shape-a')).toEqual(shape);
    expect(sessionIllustrationById(store.getState(), 'missing')).toBeNull();
    expect(updateSessionIllustration(store, 'missing', { x: 10 })).toBeNull();

    expect(
      updateSessionIllustration(store, 'shape-a', {
        x: 24,
        y: 40,
        w: 160,
        color: '#ff0000',
        radius: 16,
        lineWidth: 6,
        opacity: 0.35,
      }),
    ).toMatchObject({
      id: 'shape-a',
      x: 24,
      y: 40,
      w: 160,
      color: '#ff0000',
      radius: 16,
      lineWidth: 6,
      opacity: 0.35,
    });
    expect(clearSessionIllustration(store, 'missing')).toBe(false);
    expect(clearSessionIllustration(store, 'image-a')).toBe(true);
    expect(listSessionIllustrations(store.getState()).map((item) => item.id)).toEqual(['shape-a']);
  });

  it('rejects create, update, and clear on protected sheets', () => {
    const store = createSpreadsheetStore();
    const shape = createSessionShape(store, range(0, 0, 0, 0), { id: 'shape-a', shape: 'line' });
    expect(shape).not.toBeNull();
    setProtectedSheet(store, 0, true);

    expect(
      createSessionShape(store, range(1, 0, 1, 0), { id: 'shape-b', shape: 'oval' }),
    ).toBeNull();
    expect(createSessionImage(store, range(1, 0, 1, 0), { id: 'image-a', src: 'x' })).toBeNull();
    expect(updateSessionIllustration(store, 'shape-a', { x: 32 })).toBeNull();
    expect(duplicateSessionIllustration(store, 'shape-a')).toBeNull();
    expect(clearSessionIllustration(store, 'shape-a')).toBe(false);
    expect(listSessionIllustrations(store.getState())).toEqual([shape]);
  });

  it('duplicates session illustrations with offset, appearance, and history', () => {
    const store = createSpreadsheetStore();
    const history = new History();
    createSessionShape(
      store,
      range(0, 0, 0, 0),
      {
        id: 'shape-a',
        shape: 'rectangle',
        x: 10,
        y: 20,
        color: '#111111',
        radius: 4,
        lineWidth: 5,
        opacity: 0.4,
      },
      history,
    );

    const copy = duplicateSessionIllustration(store, 'shape-a', history);

    expect(copy).toMatchObject({
      id: 'shape-a-copy',
      x: 34,
      y: 44,
      color: '#111111',
      radius: 4,
      lineWidth: 5,
      opacity: 0.4,
    });
    expect(listSessionIllustrations(store.getState()).map((item) => item.id)).toEqual([
      'shape-a',
      'shape-a-copy',
    ]);
    expect(history.undo()).toBe(true);
    expect(listSessionIllustrations(store.getState()).map((item) => item.id)).toEqual(['shape-a']);
  });

  it('records create, update, and clear in history', () => {
    const store = createSpreadsheetStore();
    const history = new History();

    const shape = createSessionShape(
      store,
      range(0, 0, 0, 0),
      { id: 'shape-a', shape: 'line' },
      history,
    );
    expect(shape).not.toBeNull();
    expect(listSessionIllustrations(store.getState()).map((item) => item.id)).toEqual(['shape-a']);
    expect(history.undo()).toBe(true);
    expect(listSessionIllustrations(store.getState())).toEqual([]);
    expect(history.redo()).toBe(true);
    expect(listSessionIllustrations(store.getState()).map((item) => item.id)).toEqual(['shape-a']);

    expect(updateSessionIllustration(store, 'shape-a', { x: 10 }, history)).toMatchObject({
      x: 10,
    });
    expect(history.undo()).toBe(true);
    expect(sessionIllustrationById(store.getState(), 'shape-a')?.x).toBeUndefined();
    expect(history.redo()).toBe(true);
    expect(sessionIllustrationById(store.getState(), 'shape-a')?.x).toBe(10);

    expect(clearSessionIllustration(store, 'shape-a', history)).toBe(true);
    expect(listSessionIllustrations(store.getState())).toEqual([]);
    expect(history.undo()).toBe(true);
    expect(listSessionIllustrations(store.getState()).map((item) => item.id)).toEqual(['shape-a']);
  });

  it('keeps a persisted image native object in sync with history', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      expect(wb.isStub).toBe(false);
      expect(wb.capabilities.images).toBe(true);
      const store = createSpreadsheetStore();
      const history = new History();
      const image = createSessionImage(
        store,
        range(0, 1, 0, 1),
        {
          id: 'native-image',
          src: 'data:image/png;base64,placeholder',
          bytes: PNG,
          wb,
          x: 24,
          y: 20,
          w: 24,
          h: 24,
        },
        history,
      );
      expect(image?.nativeObjectId).toBeGreaterThan(0);
      expect(wb.listDrawingObjects(0).filter((object) => object.imageFormat !== 0)).toHaveLength(1);

      expect(history.undo()).toBe(true);
      expect(sessionIllustrationById(store.getState(), 'native-image')).toBeNull();
      expect(wb.listDrawingObjects(0).filter((object) => object.imageFormat !== 0)).toHaveLength(0);
      expect(history.redo()).toBe(true);
      expect(
        sessionIllustrationById(store.getState(), 'native-image')?.nativeObjectId,
      ).toBeGreaterThan(0);
      expect(wb.listDrawingObjects(0).filter((object) => object.imageFormat !== 0)).toHaveLength(1);

      const original = wb
        .listDrawingObjects(0)
        .find((object) => object.objectId === image?.nativeObjectId);
      if (!original) throw new Error('native image missing after redo');
      store.setState((state) => ({
        ...state,
        viewport: { ...state.viewport, zoom: 2, rowStart: 3, colStart: 2 },
      }));
      const target = cellRectUnclamped(
        layoutForView(store.getState()),
        store.getState().viewport,
        4,
        3,
      );
      const moved = updateSessionIllustration(
        store,
        'native-image',
        {
          x: target.x + 10 * 2,
          y: target.y + 5 * 2,
          w: (original.cx / EMU_PER_PX) * 2,
          h: (original.cy / EMU_PER_PX) * 2,
        },
        history,
        wb,
      );
      expect(moved?.nativeObjectId).toBe(image?.nativeObjectId);
      const movedNative = wb
        .listDrawingObjects(0)
        .find((object) => object.objectId === moved?.nativeObjectId);
      expect(movedNative).toMatchObject({ fromRow: 4, fromCol: 3 });
      expect(movedNative?.fromColOff).toBe(10 * EMU_PER_PX);
      expect(movedNative?.fromRowOff).toBe(5 * EMU_PER_PX);
      expect(movedNative?.cx).toBe(original.cx);
      expect(movedNative?.cy).toBe(original.cy);
      expect(history.undo()).toBe(true);
      const undoneNative = wb.listDrawingObjects(0).find((object) => object.imageFormat !== 0);
      expect(undoneNative?.objectId).toBe(original.objectId);
      expect(sessionIllustrationById(store.getState(), 'native-image')?.nativeObjectId).toBe(
        original.objectId,
      );
      expect(undoneNative).toMatchObject({ fromRow: original.fromRow, fromCol: original.fromCol });
      expect(undoneNative?.cx).toBe(original.cx);
      expect(undoneNative?.cy).toBe(original.cy);
      expect(history.redo()).toBe(true);
      const redoneNative = wb.listDrawingObjects(0).find((object) => object.imageFormat !== 0);
      expect(redoneNative).toMatchObject({ fromRow: 4, fromCol: 3 });
      expect(redoneNative?.objectId).toBe(original.objectId);
      expect(redoneNative?.cx).toBe(original.cx);
      expect(redoneNative?.cy).toBe(original.cy);

      expect(clearSessionIllustration(store, 'native-image', history, wb)).toBe(true);
      expect(wb.listDrawingObjects(0).filter((object) => object.imageFormat !== 0)).toHaveLength(0);
      expect(history.undo()).toBe(true);
      expect(wb.listDrawingObjects(0).filter((object) => object.imageFormat !== 0)).toHaveLength(1);
      expect(sessionIllustrationById(store.getState(), 'native-image')?.nativeObjectId).toBe(
        original.objectId,
      );
      expect(wb.listDrawingObjects(0).find((object) => object.imageFormat !== 0)?.objectId).toBe(
        original.objectId,
      );
      expect(history.redo()).toBe(true);
      expect(wb.listDrawingObjects(0).filter((object) => object.imageFormat !== 0)).toHaveLength(0);
    } finally {
      wb.dispose();
    }
  });

  it('moves a loaded two-cell anchored image in place and undoes it', async () => {
    const first = await WorkbookHandle.createDefault();
    let wb: WorkbookHandle | null = null;
    try {
      const inserted = first.insertImage(0, PNG, {
        row: 2,
        col: 1,
        anchorKind: 1,
        editAs: 1,
        widthEmu: 40 * EMU_PER_PX,
        heightEmu: 30 * EMU_PER_PX,
      });
      expect(inserted).not.toBeNull();
      wb = await WorkbookHandle.loadBytes(first.save());
      const loaded = wb.listDrawingObjects(0).find((object) => object.imageFormat !== 0);
      if (!loaded) throw new Error('loaded image missing');
      expect(loaded.anchorKind).toBe(1);

      const store = createSpreadsheetStore();
      const history = new History();
      hydrateImagesFromEngine(wb, store, 0);
      const id = `image-native-0-${loaded.objectId}`;
      expect(sessionIllustrationById(store.getState(), id)?.nativeGeometryEditable).toBe(true);

      const target = cellRectUnclamped(
        layoutForView(store.getState()),
        store.getState().viewport,
        6,
        4,
      );
      const moved = updateSessionIllustration(store, id, { x: target.x, y: target.y }, history, wb);
      expect(moved?.nativeObjectId).toBe(loaded.objectId);
      const after = wb.listDrawingObjects(0).find((object) => object.imageFormat !== 0);
      expect(after).toMatchObject({
        objectId: loaded.objectId,
        anchorKind: 1,
        editAs: 1,
        fromRow: 6,
        fromCol: 4,
        cx: loaded.cx,
        cy: loaded.cy,
      });
      // The projection holds the engine's anchor, to marker included.
      if (!after) throw new Error('moved image missing');
      expect(moved?.nativeAnchor).toEqual(drawingAnchorOf(after));

      expect(history.undo()).toBe(true);
      const undone = wb.listDrawingObjects(0).find((object) => object.imageFormat !== 0);
      expect(undone).toMatchObject({
        objectId: loaded.objectId,
        anchorKind: 1,
        fromRow: loaded.fromRow,
        fromCol: loaded.fromCol,
      });
      expect(history.redo()).toBe(true);
      expect(wb.listDrawingObjects(0).find((object) => object.imageFormat !== 0)).toMatchObject({
        objectId: loaded.objectId,
        fromRow: 6,
        fromCol: 4,
      });
    } finally {
      wb?.dispose();
      first.dispose();
    }
  });

  it('duplicates a native image offset from the original, with undo and redo', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      const store = createSpreadsheetStore();
      const history = new History();
      const image = createSessionImage(
        store,
        range(2, 3, 2, 3),
        { id: 'img-a', src: 'data:image/png;base64,x', bytes: PNG, wb, w: 40, h: 30 },
        null,
      );
      if (!image?.nativeAnchor) throw new Error('image was not created');
      const at = (id: string) => {
        const anchor = sessionIllustrationById(store.getState(), id)?.nativeAnchor;
        if (!anchor) throw new Error(`missing image ${id}`);
        return sessionPosition(store, anchor);
      };
      const original = at('img-a');

      const copy = duplicateSessionIllustration(store, 'img-a', history, wb);
      if (!copy?.nativeObjectId) throw new Error('image was not duplicated');

      const native = () => wb.listDrawingObjects(0).map((object) => object.objectId);
      expect(native()).toEqual([image.nativeObjectId, copy.nativeObjectId]);
      expect(at(copy.id).x).toBeCloseTo(original.x + 24, 0);
      expect(at(copy.id).y).toBeCloseTo(original.y + 24, 0);
      expect(at(copy.id).w).toBeCloseTo(original.w, 0);
      expect(at('img-a')).toEqual(original);

      expect(history.undo()).toBe(true);
      expect(native()).toEqual([image.nativeObjectId]);
      expect(history.redo()).toBe(true);
      expect(native()).toEqual([image.nativeObjectId, copy.nativeObjectId]);
      expect(at(copy.id).x).toBeCloseTo(original.x + 24, 0);
    } finally {
      wb.dispose();
    }
  });

  it('arranges a native image through the engine z-order with undo and redo', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      const store = createSpreadsheetStore();
      const history = new History();
      const make = (id: string, col: number) =>
        createSessionImage(
          store,
          range(0, col, 0, col),
          { id, src: 'data:image/png;base64,x', bytes: PNG, wb, w: 24, h: 24 },
          null,
        );
      const a = make('img-a', 0);
      const b = make('img-b', 1);
      const c = make('img-c', 2);
      const order = () => wb.listDrawingObjects(0).map((object) => object.objectId);
      const ids = [a?.nativeObjectId, b?.nativeObjectId, c?.nativeObjectId];
      expect(order()).toEqual(ids);

      expect(
        arrangeSessionIllustration(store, 'img-a', 'bring-front', history, undefined, wb),
      ).not.toBeNull();
      expect(order()).toEqual([ids[1], ids[2], ids[0]]);
      expect(listSessionIllustrations(store.getState()).map((item) => item.id)).toEqual([
        'img-b',
        'img-c',
        'img-a',
      ]);

      expect(history.undo()).toBe(true);
      expect(order()).toEqual(ids);
      expect(listSessionIllustrations(store.getState()).map((item) => item.id)).toEqual([
        'img-a',
        'img-b',
        'img-c',
      ]);
      expect(sessionIllustrationById(store.getState(), 'img-a')?.nativeObjectId).toBe(ids[0]);
      expect(history.redo()).toBe(true);
      expect(order()).toEqual([ids[1], ids[2], ids[0]]);

      expect(
        arrangeSessionIllustration(store, 'img-a', 'bring-front', history, undefined, wb),
      ).toBeNull();
      expect(arrangeSessionIllustration(store, 'img-b', 'bring-front', history)).toBeNull();
    } finally {
      wb.dispose();
    }
  });

  it('arranges session illustrations within their sheet and records history', () => {
    const store = createSpreadsheetStore();
    const history = new History();
    createSessionShape(store, range(0, 0, 0, 0), {
      id: 'back',
      shape: 'rectangle',
    });
    createSessionShape(store, range(0, 1, 0, 1), {
      id: 'middle',
      shape: 'oval',
    });
    createSessionShape(
      store,
      { sheet: 1, r0: 0, c0: 0, r1: 0, c1: 0 },
      {
        id: 'other-sheet',
        shape: 'line',
      },
    );
    createSessionImage(store, range(0, 2, 0, 2), { id: 'front', src: 'x' });

    expect(arrangeSessionIllustration(store, 'middle', 'bring-forward', history)).toMatchObject({
      id: 'middle',
    });
    expect(listSessionIllustrations(store.getState()).map((item) => item.id)).toEqual([
      'back',
      'front',
      'other-sheet',
      'middle',
    ]);

    expect(arrangeSessionIllustration(store, 'middle', 'send-back', history)).toMatchObject({
      id: 'middle',
    });
    expect(listSessionIllustrations(store.getState()).map((item) => item.id)).toEqual([
      'middle',
      'back',
      'other-sheet',
      'front',
    ]);

    expect(history.undo()).toBe(true);
    expect(listSessionIllustrations(store.getState()).map((item) => item.id)).toEqual([
      'back',
      'front',
      'other-sheet',
      'middle',
    ]);
  });

  it('ignores arrange requests that cannot change order or target protected sheets', () => {
    const store = createSpreadsheetStore();
    createSessionShape(store, range(0, 0, 0, 0), {
      id: 'shape-a',
      shape: 'rectangle',
    });
    createSessionShape(store, range(0, 1, 0, 1), {
      id: 'shape-b',
      shape: 'oval',
    });

    expect(arrangeSessionIllustration(store, 'missing', 'bring-front')).toBeNull();
    expect(arrangeSessionIllustration(store, 'shape-a', 'send-back')).toBeNull();
    setProtectedSheet(store, 0, true);
    expect(arrangeSessionIllustration(store, 'shape-b', 'send-backward')).toBeNull();
    expect(listSessionIllustrations(store.getState()).map((item) => item.id)).toEqual([
      'shape-a',
      'shape-b',
    ]);
  });
});
