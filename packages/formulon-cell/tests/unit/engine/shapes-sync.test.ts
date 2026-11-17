import type { ChartFill, ChartLine } from '@libraz/formulon';
import { AnchorEditAs, AnchorKind, ChartLineDash, ColorKind } from '@libraz/formulon';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { createDrawingClipboard } from '../../../src/commands/clipboard/drawing.js';
import { History } from '../../../src/commands/history.js';
import {
  canDeleteIllustration,
  clearSessionIllustration,
  copySessionIllustration,
  createSessionShape,
  duplicateSessionIllustration,
  pasteSessionIllustration,
  sessionIllustrationById,
  updateSessionIllustration,
} from '../../../src/commands/session-illustration.js';
import {
  applySessionShapeFill,
  applySessionShapeKind,
  applySessionShapeOutline,
  applySessionShapeOutlineDash,
  applySessionShapeOutlineWeight,
  applySessionShapeRotation,
  SHAPE_LINE_DASHES,
  shapeDashEditable,
} from '../../../src/commands/shape-format.js';
import { deleteRows, insertRows } from '../../../src/commands/structure.js';
import { hydrateImagesFromEngine } from '../../../src/engine/images-sync.js';
import { captureShapeRestore, shapeFromNative } from '../../../src/engine/shapes-sync.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import type { EngineShape } from '../../../src/engine/workbook-handle-drawing-objects.js';
import { en } from '../../../src/i18n/strings.js';
import { attachClipboard } from '../../../src/interact/clipboard.js';
import { attachSessionIllustrations } from '../../../src/interact/session-illustrations.js';
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

/** How a shape paints its fill or outline, as these tests compare it. */
const ShapePaint = { Unset: 0, Solid: 1, None: 2, Other: 3 } as const;
type ShapePaint = (typeof ShapePaint)[keyof typeof ShapePaint];

const EMU_PER_PT = 12700;

const paintKind = (paint: ChartFill | ChartLine | null): ShapePaint => {
  if (paint === null) return ShapePaint.Unset;
  if (paint.other) return ShapePaint.Other;
  if (paint.none) return ShapePaint.None;
  return paint.color === null ? ShapePaint.Unset : ShapePaint.Solid;
};

/** A shape read back, flattened: anchor fields at the top, paint as kind and RGB. */
const flat = (shape: EngineShape) => ({
  objectId: shape.objectId,
  name: shape.name,
  prst: shape.presetGeometry,
  text: shape.text,
  textBox: shape.textBox,
  hidden: shape.hidden,
  rotation: shape.rotationDeg,
  anchorKind: shape.anchor.kind,
  editAs: shape.anchor.editAs,
  fromRow: shape.anchor.fromRow,
  fromCol: shape.anchor.fromCol,
  fromRowOff: shape.anchor.fromRowOffsetEmu,
  fromColOff: shape.anchor.fromColOffsetEmu,
  toRow: shape.anchor.toRow,
  toCol: shape.anchor.toCol,
  toRowOff: shape.anchor.toRowOffsetEmu,
  toColOff: shape.anchor.toColOffsetEmu,
  cx: shape.anchor.widthEmu,
  cy: shape.anchor.heightEmu,
  fillKind: paintKind(shape.fill),
  fillRgb: (shape.fill?.color?.resolved?.argb ?? 0) & 0xffffff,
  lineKind: paintKind(shape.line),
  lineRgb: (shape.line?.color?.resolved?.argb ?? 0) & 0xffffff,
  lineWidthEmu: Math.round((shape.line?.widthPt ?? 0) * EMU_PER_PT),
});

const listShapes = (wb: WorkbookHandle, sheet: number) => wb.listShapes(sheet).map(flat);

interface FlatShapeInput {
  name?: string;
  prst?: string;
  text?: string;
  textBox?: boolean;
  hidden?: boolean;
  anchorKind?: AnchorKind;
  editAs?: AnchorEditAs;
  fromRow?: number;
  fromCol?: number;
  toRow?: number;
  toCol?: number;
  cx?: number;
  cy?: number;
  fillKind?: ShapePaint;
  fillRgb?: number;
  lineKind?: ShapePaint;
  lineRgb?: number;
  lineWidthEmu?: number;
}

const hex = (rgb: number): string => `#${rgb.toString(16).padStart(6, '0')}`;

/** Insert a shape described in the flattened form. */
const insertShape = (wb: WorkbookHandle, sheet: number, input: FlatShapeInput) =>
  wb.insertShape(sheet, {
    ...(input.name !== undefined ? { name: input.name } : {}),
    ...(input.prst !== undefined ? { presetGeometry: input.prst } : {}),
    ...(input.text !== undefined ? { text: input.text } : {}),
    ...(input.textBox !== undefined ? { textBox: input.textBox } : {}),
    ...(input.hidden !== undefined ? { hidden: input.hidden } : {}),
    anchor: {
      ...(input.anchorKind !== undefined ? { kind: input.anchorKind } : {}),
      ...(input.editAs !== undefined ? { editAs: input.editAs } : {}),
      ...(input.fromRow !== undefined ? { fromRow: input.fromRow } : {}),
      ...(input.fromCol !== undefined ? { fromCol: input.fromCol } : {}),
      ...(input.toRow !== undefined ? { toRow: input.toRow } : {}),
      ...(input.toCol !== undefined ? { toCol: input.toCol } : {}),
      ...(input.cx !== undefined ? { widthEmu: input.cx } : {}),
      ...(input.cy !== undefined ? { heightEmu: input.cy } : {}),
    },
    ...(input.fillKind === ShapePaint.Solid ? { fill: { color: hex(input.fillRgb ?? 0) } } : {}),
    ...(input.fillKind === ShapePaint.None ? { fill: { none: true } } : {}),
    ...(input.lineKind === ShapePaint.Solid
      ? {
          line: {
            color: hex(input.lineRgb ?? 0),
            ...(input.lineWidthEmu !== undefined
              ? { widthPt: input.lineWidthEmu / EMU_PER_PT }
              : {}),
          },
        }
      : {}),
    ...(input.lineKind === ShapePaint.None ? { line: { none: true } } : {}),
  });

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

    const [native] = listShapes(wb, 0);
    expect(listShapes(wb, 0)).toHaveLength(1);
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
    expect(only(store).nativeAnchor?.widthEmu).toBeGreaterThan(0);

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

    expect(listShapes(wb, 0).map((shape) => shape.prst)).toEqual([
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
    expect(listShapes(wb, 0)[0]).toMatchObject({ fillKind: ShapePaint.Solid, fillRgb: 0x00ff00 });
    style((item) => applySessionShapeOutline(store, item, '#0000ff', history, wb));
    style((item) => applySessionShapeOutlineWeight(store, item, 4.5, history, wb));
    expect(listShapes(wb, 0)[0]).toMatchObject({
      lineKind: ShapePaint.Solid,
      lineRgb: 0x0000ff,
      lineWidthEmu: 57150,
    });
    style((item) => applySessionShapeFill(store, item, null, history, wb));
    expect(listShapes(wb, 0)[0]?.fillKind).toBe(ShapePaint.None);
    style((item) => applySessionShapeOutline(store, item, null, history, wb));
    expect(listShapes(wb, 0)[0]?.lineKind).toBe(ShapePaint.None);

    const reopened = await saved(wb);
    expect(only(reopened.store)).toMatchObject({ opacity: 0, lineHidden: true, lineWidth: 4.5 });
  });

  it('keeps the outline weight in the file while the outline is off', async () => {
    const { wb, store, history } = await session();
    oval(store);
    applySessionShapeOutlineWeight(store, only(store), 3, history, wb);
    applySessionShapeOutline(store, only(store), null, history, wb);
    expect(listShapes(wb, 0)[0]).toMatchObject({
      lineKind: ShapePaint.None,
      lineWidthEmu: 38100,
    });

    const reopened = await saved(wb);
    expect(only(reopened.store)).toMatchObject({ lineHidden: true, lineWidth: 3 });
    applySessionShapeOutline(
      reopened.store,
      only(reopened.store),
      '#0000ff',
      reopened.history,
      reopened.wb,
    );
    expect(listShapes(reopened.wb, 0)[0]).toMatchObject({
      lineKind: ShapePaint.Solid,
      lineWidthEmu: 38100,
    });
  });

  it('resets radius and a dash style the file cannot hold', async () => {
    const { wb, store, history } = await session();
    oval(store);

    updateSessionIllustration(store, 's1', { radius: 12, lineDash: 'longDash' }, history, wb);

    expect(only(store)).toMatchObject({ radius: undefined, lineDash: undefined });
    expect(listShapes(wb, 0)[0]?.fillKind).toBe(ShapePaint.Solid);
    updateSessionIllustration(store, 's1', { opacity: 0 }, history, wb);
    expect(only(store).opacity).toBe(0);
    expect(listShapes(wb, 0)[0]?.fillKind).toBe(ShapePaint.None);
  });

  it('refuses a dash style the file cannot hold', async () => {
    const { wb, store, history } = await session();
    oval(store);
    const before = only(store);

    expect(applySessionShapeOutlineDash(store, before, 'roundDot', history, wb)).toBeNull();
    expect(only(store).lineDash).toBeUndefined();
  });

  it('moves and resizes a shape through its anchor', async () => {
    const { wb, store, history } = await session();
    oval(store);
    const before = listShapes(wb, 0)[0];

    updateSessionIllustration(store, 's1', { x: 300, y: 200, w: 120, h: 60 }, history, wb);

    const after = listShapes(wb, 0)[0];
    expect(after?.objectId).toBe(before?.objectId);
    expect(after?.cx).toBeCloseTo(120 * 9525, -2);
    expect(after?.cy).toBeCloseTo(60 * 9525, -2);
    expect([after?.fromRow, after?.fromCol]).not.toEqual([before?.fromRow, before?.fromCol]);
    expect(only(store).nativeAnchor).toMatchObject({
      fromRow: after?.fromRow,
      fromCol: after?.fromCol,
      widthEmu: after?.cx,
    });

    const reopened = await saved(wb);
    expect(only(reopened.store).nativeAnchor).toEqual(only(store).nativeAnchor);
  });

  it('writes the text of a shape and reads it back', async () => {
    const { wb, store, history } = await session();
    oval(store);

    updateSessionIllustration(store, 's1', { text: 'Total\nQ3' }, history, wb);

    expect(listShapes(wb, 0)[0]?.text).toBe('Total\nQ3');
    const reopened = await saved(wb);
    expect(only(reopened.store).text).toBe('Total\nQ3');
  });

  it('turns a shape a quarter turn in the file, and back on undo', async () => {
    const { wb, store, history } = await session();
    oval(store);

    applySessionShapeRotation(store, only(store), -90, history, wb);
    expect(listShapes(wb, 0)[0]?.rotation).toBe(270);
    expect(only((await saved(wb)).store).rotationDeg).toBe(270);
    applySessionShapeRotation(store, only(store), 90, history, wb);
    expect(listShapes(wb, 0)[0]?.rotation).toBe(0);

    history.undo();
    expect(listShapes(wb, 0)[0]?.rotation).toBe(270);
  });

  it('changes the kind of a shape in the file, keeping its paint, text and anchor', async () => {
    const { wb, store, history } = await session();
    oval(store);
    updateSessionIllustration(store, 's1', { text: 'Total' }, history, wb);
    const before = listShapes(wb, 0)[0];

    applySessionShapeKind(store, only(store), 'diamond', history, wb);
    expect(listShapes(wb, 0)[0]).toEqual({ ...before, prst: 'diamond' });
    expect(only((await saved(wb)).store)).toMatchObject({ shape: 'diamond', text: 'Total' });

    history.undo();
    expect(listShapes(wb, 0)[0]).toEqual(before);
  });

  it('deletes a shape from the file', async () => {
    const { wb, store, history } = await session();
    oval(store);

    expect(clearSessionIllustration(store, 's1', history, wb)).toBe(true);

    expect(listShapes(wb, 0)).toEqual([]);
    expect(shapesOf((await saved(wb)).store)).toEqual([]);
  });

  it('copies a shape into a second engine shape', async () => {
    const { wb, store, history } = await session();
    oval(store);

    const copy = duplicateSessionIllustration(store, 's1', history, wb);

    expect(listShapes(wb, 0)).toHaveLength(2);
    const [first, second] = listShapes(wb, 0);
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
    expect(listShapes(wb, 0)).toHaveLength(1);

    applySessionShapeFill(store, only(store), '#00ff00', history, wb);
    updateSessionIllustration(store, 's1', { x: 400, y: 300 }, history, wb);
    const moved = listShapes(wb, 0)[0];
    clearSessionIllustration(store, 's1', history, wb);
    expect(listShapes(wb, 0)).toEqual([]);

    history.undo();
    expect(listShapes(wb, 0)).toHaveLength(1);
    expect(listShapes(wb, 0)[0]).toMatchObject({
      fromRow: moved?.fromRow,
      fromCol: moved?.fromCol,
      fillRgb: 0x00ff00,
    });
    expect(only(store).nativeObjectId).toBe(listShapes(wb, 0)[0]?.objectId);

    history.undo();
    history.undo();
    expect(listShapes(wb, 0)[0]).toMatchObject({ fillRgb: 0xff0000 });
    expect(listShapes(wb, 0)[0]?.fromCol).not.toBe(moved?.fromCol);

    history.undo();
    expect(listShapes(wb, 0)).toEqual([]);
    expect(shapesOf(store)).toEqual([]);

    history.redo();
    history.redo();
    history.redo();
    history.redo();
    expect(listShapes(wb, 0)).toEqual([]);
    history.undo();
    expect(listShapes(wb, 0)).toHaveLength(1);
    expect(shapesOf((await saved(wb)).store)).toHaveLength(1);
  });

  it('shifts a shape once when rows are inserted, and back on undo', async () => {
    const { wb, store, history } = await session();
    oval(store);
    const row = listShapes(wb, 0)[0]?.fromRow ?? 0;

    expect(insertRows(store, wb, history, 0, 3)).toBe(true);

    expect(listShapes(wb, 0)[0]?.fromRow).toBe(row + 3);
    expect(listShapes(wb, 0)).toHaveLength(1);
    expect(only(store).nativeAnchor?.fromRow).toBe(row + 3);

    history.undo();
    expect(listShapes(wb, 0)[0]?.fromRow).toBe(row);
    expect(only(store).nativeAnchor?.fromRow).toBe(row);
    history.redo();
    expect(listShapes(wb, 0)[0]?.fromRow).toBe(row + 3);
  });

  it('keeps a shape in the file when an earlier sheet is removed', async () => {
    const { wb, store } = await session();
    const second = wb.addSheet();
    createSessionShape(
      store,
      { sheet: second, r0: 0, c0: 0, r1: 0, c1: 0 },
      { id: 's2', shape: 'rectangle', x: 100, y: 60, w: 120, h: 80 },
    );
    expect(listShapes(wb, second)).toHaveLength(1);

    expect(wb.removeSheet(0)).toBe(true);

    expect(listShapes(wb, 0)).toHaveLength(1);
    expect(sessionIllustrationById(store.getState(), 's2')?.sheet).toBe(0);
  });

  it('puts a collapsed shape back where it was when a row delete is undone', async () => {
    const { wb, store, history } = await session();
    oval(store);
    const before = listShapes(wb, 0)[0];

    expect(deleteRows(store, wb, history, (before?.fromRow ?? 0) - 1, 4)).toBe(true);
    history.undo();

    expect(listShapes(wb, 0)).toHaveLength(1);
    expect(listShapes(wb, 0)[0]).toMatchObject({
      fromRow: before?.fromRow,
      fromCol: before?.fromCol,
      cx: before?.cx,
      cy: before?.cy,
    });
    expect(only(store).nativeAnchor?.fromRow).toBe(before?.fromRow);
  });

  it('puts back a two-cell shape that a row delete took, and its style', async () => {
    const { wb, store, history } = await session((handle) => {
      insertShape(handle, 0, {
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
    expect(listShapes(wb, 0)).toEqual([]);
    expect(shapesOf(store)).toEqual([]);

    history.undo();
    expect(listShapes(wb, 0)).toHaveLength(1);
    expect(listShapes(wb, 0)[0]).toMatchObject({
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
    insertShape(wb, 0, {
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
    const before = listShapes(wb, 0)[0];

    expect(clearSessionIllustration(store, only(store).id, history, wb)).toBe(true);
    expect(listShapes(wb, 0)).toEqual([]);
    history.undo();

    expect(listShapes(wb, 0)).toHaveLength(1);
    expect(listShapes(wb, 0)[0]).toMatchObject({
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
    const before = listShapes(wb, 0)[0];

    clearSessionIllustration(store, 's1', history, wb);
    store.setState((state) => ({
      ...state,
      viewport: { ...state.viewport, rowStart: 7, colStart: 3 },
    }));
    history.undo();

    expect(listShapes(wb, 0)[0]).toMatchObject({
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
    insertShape(wb, 0, {
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

  it('shows a custom geometry, which has no preset name, as a locked box it cannot restore', async () => {
    const { wb, store } = await session(star);
    const shape = wb.listShapes(0)[0];
    if (!shape) throw new Error('no shape');
    const custom = { ...shape, presetGeometry: null };
    expect(shapeFromNative(store, 0, custom, wb)).toMatchObject({ nativeLocked: true });
    const handle = { listShapes: () => [custom] } as unknown as WorkbookHandle;
    expect(captureShapeRestore(handle, 0, shape.objectId)).toBeNull();
  });

  it('treats a loaded line as a locked box, since its direction is unknown', async () => {
    const { store } = await session((wb) => {
      insertShape(wb, 0, {
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
    expect(listShapes(wb, 0)).toHaveLength(1);
    expect(canDeleteIllustration(only(store))).toBe(false);
  });

  it('moves it without touching its preset or fill, and refuses restyling', async () => {
    const { wb, store, history } = await session(star);
    const item = only(store);

    expect(applySessionShapeFill(store, item, '#000000', history, wb)).toBeNull();
    updateSessionIllustration(store, item.id, { x: 500, y: 400 }, history, wb);

    expect(listShapes(wb, 0)[0]).toMatchObject({
      prst: 'star5',
      fillKind: ShapePaint.Solid,
      fillRgb: 0xffcc00,
    });
    expect(listShapes(wb, 0)[0]?.fromRow).not.toBe(10);
  });

  it('cannot be copied, and leaves hidden shapes alone', async () => {
    const { wb, store, history } = await session((handle) => {
      star(handle);
      insertShape(handle, 0, {
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
    expect(listShapes(wb, 0)).toHaveLength(2);
    expect(listShapes(wb, 0)[1]?.hidden).toBe(true);
  });

  it('says why a keyboard duplicate or copy is refused', async () => {
    const { wb, store, history } = await session(star);
    const host = document.createElement('div');
    document.body.appendChild(host);
    vi.stubGlobal('navigator', { clipboard: { writeText: async () => {} } });
    const notice = () => {
      const dialog = document.querySelector('[role="alertdialog"]');
      const text = dialog?.querySelector('.fc-tb__dlg__message')?.textContent;
      dialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
      return text;
    };
    const illustrations = attachSessionIllustrations({
      host,
      store,
      history,
      getWb: () => wb,
      getStrings: () => en,
    });
    const clipboard = attachClipboard({
      host,
      store,
      wb,
      history,
      onAfterCommit: () => {},
      getStrings: () => en,
    });
    try {
      const panel = host.querySelector<HTMLElement>('.fc-illustration');
      if (!panel) throw new Error('no shape panel');
      panel.dispatchEvent(new PointerEvent('pointerdown', { bubbles: true, button: 0 }));
      window.dispatchEvent(new PointerEvent('pointerup', { bubbles: true }));
      panel.dispatchEvent(
        new KeyboardEvent('keydown', { bubbles: true, cancelable: true, key: 'd', ctrlKey: true }),
      );
      expect(notice()).toBe(en.workbookObjects.shapeLocked);
      await clipboard.runShortcut('copy');
      expect(notice()).toBe(en.workbookObjects.shapeLocked);
      expect(clipboard.getSnapshot()).toBeNull();
      expect(shapesOf(store)).toHaveLength(1);
    } finally {
      clipboard.detach();
      illustrations.detach();
      vi.unstubAllGlobals();
      host.remove();
    }
  });
});

describe('theme colours of saved shapes', () => {
  const accentFill = { kind: ColorKind.Theme, theme: 5, tint: 0.4, argb: 0, indexed: 0 } as const;
  const accentLine = { kind: ColorKind.Theme, theme: 4, tint: -0.25, argb: 0, indexed: 0 } as const;

  const accented = (wb: WorkbookHandle): void => {
    wb.insertShape(0, {
      presetGeometry: 'rect',
      anchor: {
        kind: AnchorKind.TwoCell,
        editAs: AnchorEditAs.TwoCell,
        fromRow: 10,
        fromCol: 2,
        toRow: 12,
        toCol: 4,
      },
      fill: { color: accentFill },
      line: { color: accentLine, widthPt: 2 },
    });
  };

  // Excel's rendering of these in the default theme: accent2 at 60 % / +40 %, accent1 at 75 %.
  const ACCENT_FILL_CSS = '#f2aa84';
  const ACCENT_LINE_CSS = '#104862';

  /** The theme reference of a paint's colour, as the engine holds it. */
  const themeOf = (paint: ChartFill | ChartLine | null | undefined) => ({
    kind: paint?.color?.base.kind,
    theme: paint?.color?.base.theme,
    transforms: paint?.color?.transforms,
  });

  /** The theme references of the shape as first inserted. */
  let inserted: { fill: ReturnType<typeof themeOf>; line: ReturnType<typeof themeOf> };

  const recordAccents = (wb: WorkbookHandle): void => {
    const shape = wb.listShapes(0)[0];
    inserted = { fill: themeOf(shape?.fill), line: themeOf(shape?.line) };
    expect(inserted.fill).toMatchObject({ kind: ColorKind.Theme, theme: 5 });
    expect(inserted.fill.transforms?.length).toBeGreaterThan(0);
    expect(inserted.line).toMatchObject({ kind: ColorKind.Theme, theme: 4 });
    expect(inserted.line.transforms?.length).toBeGreaterThan(0);
  };

  const expectAccents = (shape: EngineShape | undefined): void => {
    expect(themeOf(shape?.fill)).toEqual(inserted.fill);
    expect(themeOf(shape?.line)).toEqual(inserted.line);
  };

  it('shows the colour the theme renders', async () => {
    const { store } = await session(accented);

    expect(only(store)).toMatchObject({
      color: ACCENT_FILL_CSS,
      lineColor: ACCENT_LINE_CSS,
    });
  });

  it('keeps the theme colours through duplicate, paste, sheet move and a save', async () => {
    const { wb, store, history } = await session(accented);
    recordAccents(wb);
    const id = only(store).id;

    duplicateSessionIllustration(store, id, history, wb);
    const clip = copySessionIllustration(store, id, wb);
    if (!clip) throw new Error('no clip');
    pasteSessionIllustration(store, clip, { sheet: 0, row: 20, col: 1 }, history, wb);
    expect(wb.listShapes(0)).toHaveLength(3);
    for (const shape of wb.listShapes(0)) expectAccents(shape);

    const second = wb.addSheet();
    updateSessionIllustration(store, id, { sheet: second }, history, wb);
    expectAccents(wb.listShapes(second)[0]);

    const reopened = await saved(wb);
    for (const shape of reopened.wb.listShapes(0)) expectAccents(shape);
    expect(shapesOf(reopened.store)[0]?.color).toBe(ACCENT_FILL_CSS);
  });

  it('puts the theme colours back when a row delete is undone', async () => {
    const { wb, store, history } = await session(accented);
    recordAccents(wb);

    expect(deleteRows(store, wb, history, 9, 5)).toBe(true);
    history.undo();

    expectAccents(wb.listShapes(0)[0]);
  });

  it('writes a colour the user picks as RGB', async () => {
    const { wb, store, history } = await session(accented);

    applySessionShapeFill(store, only(store), '#00ff00', history, wb);
    applySessionShapeOutline(store, only(store), '#0000ff', history, wb);

    expect(listShapes(wb, 0)[0]).toMatchObject({
      fillKind: ShapePaint.Solid,
      fillRgb: 0x00ff00,
      lineRgb: 0x0000ff,
    });
    expect(wb.listShapes(0)[0]?.fill?.color?.base.kind).toBe(ColorKind.Rgb);
    expect(wb.listShapes(0)[0]?.line?.color?.base.kind).toBe(ColorKind.Rgb);
  });
});

describe('dash styles of saved shapes', () => {
  it('saves the dashes the file can hold and reads them back', async () => {
    const { wb, store, history } = await session();
    oval(store);

    expect(applySessionShapeOutlineDash(store, only(store), 'dash', history, wb)).not.toBeNull();
    expect(wb.listShapes(0)[0]?.line?.dash).toBe(ChartLineDash.Dash);
    expect(only((await saved(wb)).store).lineDash).toBe('dash');

    applySessionShapeOutlineDash(store, only(store), 'dashDot', history, wb);
    expect(wb.listShapes(0)[0]?.line?.dash).toBe(ChartLineDash.DashDot);
    expect(only((await saved(wb)).store).lineDash).toBe('dashDot');

    applySessionShapeOutlineDash(store, only(store), 'solid', history, wb);
    expect([null, ChartLineDash.Solid]).toContain(wb.listShapes(0)[0]?.line?.dash);
    expect(only((await saved(wb)).store).lineDash).toBeUndefined();
  });

  it('keeps a stored dash through a colour change and a duplicate', async () => {
    const { wb, store, history } = await session();
    oval(store);
    applySessionShapeOutlineDash(store, only(store), 'dash', history, wb);

    applySessionShapeOutline(store, only(store), '#0000ff', history, wb);
    duplicateSessionIllustration(store, 's1', history, wb);

    expect(wb.listShapes(0).map((shape) => shape.line?.dash)).toEqual([
      ChartLineDash.Dash,
      ChartLineDash.Dash,
    ]);
  });

  it('enables exactly the storable dashes for a saved shape', async () => {
    const { store } = await session();
    oval(store);

    expect(SHAPE_LINE_DASHES.filter((dash) => shapeDashEditable(only(store), dash))).toEqual([
      'solid',
      'dash',
      'dashDot',
    ]);
  });
});

describe('transparency of saved shapes', () => {
  it('saves a partly transparent fill and reads it back', async () => {
    const { wb, store, history } = await session();
    oval(store);

    updateSessionIllustration(store, 's1', { opacity: 0.4 }, history, wb);

    expect(only(store).opacity).toBe(0.4);
    expect(wb.listShapes(0)[0]?.fill?.transparencyPct).toBe(60);
    expect(only((await saved(wb)).store).opacity).toBe(0.4);

    updateSessionIllustration(store, 's1', { opacity: 1 }, history, wb);
    expect(wb.listShapes(0)[0]?.fill?.transparencyPct ?? 0).toBe(0);
  });
});

describe('name, alt text and text box of saved shapes', () => {
  const note = (wb: WorkbookHandle): void => {
    wb.insertShape(0, {
      name: 'Callout',
      description: 'Quarter total',
      presetGeometry: 'rect',
      textBox: true,
      text: 'note',
      anchor: {
        kind: AnchorKind.TwoCell,
        editAs: AnchorEditAs.TwoCell,
        fromRow: 10,
        fromCol: 2,
        toRow: 12,
        toCol: 4,
      },
    });
  };

  const identity = (shape: EngineShape | undefined) => ({
    name: shape?.name,
    description: shape?.description,
    textBox: shape?.textBox,
  });
  const kept = { name: 'Callout', description: 'Quarter total', textBox: true };

  it('gives a copy the source name, alt text and text box, and its own id', async () => {
    const { wb, store, history } = await session(note);
    const id = only(store).id;

    duplicateSessionIllustration(store, id, history, wb);
    const clip = copySessionIllustration(store, id, wb);
    if (!clip) throw new Error('no clip');
    pasteSessionIllustration(store, clip, { sheet: 0, row: 20, col: 1 }, history, wb);

    expect(wb.listShapes(0).map(identity)).toEqual([kept, kept, kept]);
    expect(new Set(wb.listShapes(0).map((shape) => shape.objectId)).size).toBe(3);
    const reopened = await saved(wb);
    expect(reopened.wb.listShapes(0).map(identity)).toEqual([kept, kept, kept]);
  });

  it('keeps all three through cut-paste and a sheet move', async () => {
    const { wb, store, history } = await session(note);
    const clipboard = createDrawingClipboard({ store, wb, history });
    mutators.setSelectedObject(store, { id: only(store).id, name: '' });

    expect(clipboard.take('cut')).not.toBeNull();
    expect(wb.listShapes(0)).toEqual([]);
    expect(clipboard.paste()).toBe(true);
    expect(wb.listShapes(0).map(identity)).toEqual([kept]);

    const second = wb.addSheet();
    updateSessionIllustration(store, only(store).id, { sheet: second }, history, wb);
    expect(identity(wb.listShapes(second)[0])).toEqual(kept);
  });
});

describe('copies of a saved shape', () => {
  const dotted = (wb: WorkbookHandle): void => {
    wb.insertShape(0, {
      presetGeometry: 'rect',
      anchor: {
        kind: AnchorKind.TwoCell,
        editAs: AnchorEditAs.TwoCell,
        fromRow: 10,
        fromCol: 2,
        toRow: 12,
        toCol: 4,
      },
      fill: { color: '#336699' },
      line: { color: '#ff0000', widthPt: 2, dash: ChartLineDash.Dot, transparencyPct: 40 },
    });
  };

  it('keeps every outline value the file holds, and edits written since the load', async () => {
    const { wb, store, history } = await session(dotted);
    const id = only(store).id;
    applySessionShapeFill(store, only(store), '#00ff00', history, wb);

    duplicateSessionIllustration(store, id, history, wb);
    const clip = copySessionIllustration(store, id, wb);
    if (!clip) throw new Error('no clip');
    pasteSessionIllustration(store, clip, { sheet: 0, row: 20, col: 1 }, history, wb);

    const shapes = wb.listShapes(0);
    expect(shapes).toHaveLength(3);
    for (const shape of shapes) {
      expect(shape.line).toMatchObject({
        dash: ChartLineDash.Dot,
        transparencyPct: 40,
        widthPt: 2,
      });
      expect((shape.fill?.color?.resolved?.argb ?? 0) & 0xffffff).toBe(0x00ff00);
    }
    expect(shapes[1]?.anchor.kind).toBe(AnchorKind.TwoCell);
    expect(shapes[1]?.anchor).not.toEqual(shapes[0]?.anchor);
    expect(shapes[2]?.anchor).toMatchObject({ kind: AnchorKind.TwoCell, fromRow: 20, fromCol: 1 });
  });
});
