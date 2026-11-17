import { AnchorEditAs, AnchorKind, ShapeTextVert } from '@libraz/formulon';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { History } from '../../../src/commands/history.js';
import {
  copySessionIllustration,
  duplicateSessionIllustration,
  pasteSessionIllustration,
  updateSessionIllustration,
} from '../../../src/commands/session-illustration.js';
import { applySessionShapeFill } from '../../../src/commands/shape-format.js';
import { hydrateImagesFromEngine } from '../../../src/engine/images-sync.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import {
  createSpreadsheetStore,
  type SessionIllustration,
  type SessionTextVert,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

const open: WorkbookHandle[] = [];
afterEach(() => {
  for (const wb of open.splice(0)) wb.dispose();
});

const VERTS: [SessionTextVert, ShapeTextVert][] = [
  ['horz', ShapeTextVert.Horz],
  ['vert', ShapeTextVert.Vert],
  ['vert270', ShapeTextVert.Vert270],
  ['eaVert', ShapeTextVert.EaVert],
  ['wordArtVert', ShapeTextVert.WordArtVert],
  ['wordArtVertRtl', ShapeTextVert.WordArtVertRtl],
  ['mongolianVert', ShapeTextVert.MongolianVert],
];

const seed = (textVert?: ShapeTextVert) => (wb: WorkbookHandle) => {
  wb.insertShape(0, {
    presetGeometry: 'rect',
    text: 'abc',
    ...(textVert === undefined ? {} : { textVert }),
    anchor: {
      kind: AnchorKind.TwoCell,
      editAs: AnchorEditAs.TwoCell,
      fromRow: 4,
      fromCol: 1,
      toRow: 8,
      toCol: 4,
    },
    fill: { color: '#336699' },
  });
};

const session = async (fill: (wb: WorkbookHandle) => void, bytes?: Uint8Array) => {
  const wb = bytes ? await WorkbookHandle.loadBytes(bytes) : await WorkbookHandle.createDefault();
  open.push(wb);
  fill(wb);
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

const vertOf = (wb: WorkbookHandle, index = 0) => wb.listShapes(0)[index]?.textVert;

describe('the text direction of a loaded shape', () => {
  it.each(VERTS)('reads %s as the file states it', async (token, value) => {
    const { store } = await session(seed(value));
    expect(only(store).textVert).toBe(token);
  });

  it('leaves an unstated direction unstated', async () => {
    const { store } = await session(seed());
    expect(only(store).textVert).toBeUndefined();
  });

  it.each(VERTS)('keeps %s through an unrelated edit and a save', async (_token, value) => {
    const { wb, store, history } = await session(seed(value));
    applySessionShapeFill(store, only(store), '#00ff00', history, wb);
    updateSessionIllustration(store, only(store).id, { x: 90, y: 90 }, history, wb);

    expect(vertOf(wb)).toBe(value);
    const reopened = await session(() => undefined, wb.save());
    expect(vertOf(reopened.wb)).toBe(value);
  });

  it('writes no direction for a shape whose value did not change', async () => {
    const { wb, store, history } = await session(seed(ShapeTextVert.Vert));
    const update = vi.spyOn(wb, 'updateShape');

    applySessionShapeFill(store, only(store), '#00ff00', history, wb);

    expect(update).toHaveBeenCalled();
    for (const [, , patch] of update.mock.calls) expect(patch).not.toHaveProperty('textVert');
  });

  it.each(VERTS)('keeps %s through duplicate, copy and paste', async (_token, value) => {
    const { wb, store, history } = await session(seed(value));
    const id = only(store).id;

    duplicateSessionIllustration(store, id, history, wb);
    const clip = copySessionIllustration(store, id, wb);
    if (!clip) throw new Error('no clip');
    pasteSessionIllustration(store, clip, { sheet: 0, row: 20, col: 1 }, history, wb);

    expect(wb.listShapes(0)).toHaveLength(3);
    expect([vertOf(wb, 0), vertOf(wb, 1), vertOf(wb, 2)]).toEqual([value, value, value]);
  });

  it('changes and clears the direction in the file', async () => {
    const { wb, store, history } = await session(seed());
    const id = only(store).id;

    updateSessionIllustration(store, id, { textVert: 'vert270' }, history, wb);
    expect(vertOf(wb)).toBe(ShapeTextVert.Vert270);
    updateSessionIllustration(store, id, { textVert: undefined }, history, wb);
    expect(vertOf(wb)).toBeNull();
  });
});
