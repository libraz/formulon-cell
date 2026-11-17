import { AnchorEditAs, AnchorKind } from '@libraz/formulon';
import { afterEach, describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import {
  sessionIllustrationById,
  updateSessionIllustration,
} from '../../../src/commands/session-illustration.js';
import {
  applySessionDrawingFlip,
  applySessionShapeRotation,
} from '../../../src/commands/shape-format.js';
import { EMU_PER_PX } from '../../../src/engine/drawing-geometry.js';
import { hydrateImagesFromEngine } from '../../../src/engine/images-sync.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';
import { zipEntries, zipStored } from '../engine/zip-entries.js';

const PNG = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=',
  ),
  (char) => char.charCodeAt(0),
);

const open: WorkbookHandle[] = [];
afterEach(() => {
  for (const wb of open.splice(0)) wb.dispose();
});

type AnchorCase = 'twoCell' | 'oneCell' | 'absolute';

/** A workbook holding one picture anchored as `anchor`, bound to a store. */
async function picture(anchor: AnchorCase, seed?: (wb: WorkbookHandle, id: number) => void) {
  const source = await WorkbookHandle.createDefault();
  const frame = { fromRow: 4, fromCol: 2, widthEmu: 40 * EMU_PER_PX, heightEmu: 30 * EMU_PER_PX };
  const placed =
    anchor === 'twoCell'
      ? { ...frame, kind: AnchorKind.TwoCell, editAs: AnchorEditAs.TwoCell }
      : { ...frame, kind: AnchorKind.OneCell };
  const objectId = source.insertImage(0, PNG, { anchor: placed, name: 'Logo' }) ?? 0;
  let bytes = source.save();
  source.dispose();
  if (anchor === 'absolute') {
    const entries = zipEntries(bytes);
    const path = 'xl/drawings/drawing1.xml';
    const xml = new TextDecoder()
      .decode(entries.get(path))
      .replace(
        /<xdr:oneCellAnchor><xdr:from>.*?<\/xdr:from>/,
        '<xdr:absoluteAnchor><xdr:pos x="1219200" y="762000"/>',
      )
      .replace('</xdr:oneCellAnchor>', '</xdr:absoluteAnchor>');
    entries.set(path, new TextEncoder().encode(xml));
    bytes = zipStored(entries);
  }
  const wb = await WorkbookHandle.loadBytes(bytes);
  open.push(wb);
  seed?.(wb, objectId);
  const store = createSpreadsheetStore();
  hydrateImagesFromEngine(wb, store, 0);
  wb.attachStore(store);
  const history = new History();
  wb.attachHistory(history);
  const id = `image-native-0-${objectId}`;
  const current = () => {
    const item = sessionIllustrationById(store.getState(), id);
    if (!item) throw new Error('picture is gone');
    return item;
  };
  return { wb, store, history, objectId, current };
}

describe.each(['twoCell', 'oneCell', 'absolute'] as const)(
  'a loaded picture anchored %s',
  (anchor) => {
    it('rotates and flips through the engine, keeping them through save and undo', async () => {
      const { wb, store, history, objectId, current } = await picture(anchor);
      const placed = wb.listDrawingObjects(0).find((o) => o.objectId === objectId)?.anchor;

      applySessionShapeRotation(store, current(), 90, history, wb);
      expect(wb.getImage(0, objectId)).toMatchObject({ rotationDeg: 90, flipH: false });
      applySessionDrawingFlip(store, current(), 'horizontal', history, wb);
      expect(wb.getImage(0, objectId)).toMatchObject({ rotationDeg: 270, flipH: true });
      expect(current()).toMatchObject({ rotationDeg: 270, flipH: true });
      expect(wb.listDrawingObjects(0).find((o) => o.objectId === objectId)?.anchor).toEqual(placed);

      const reopened = await WorkbookHandle.loadBytes(wb.save());
      open.push(reopened);
      expect(reopened.getImage(0, objectId)).toMatchObject({ rotationDeg: 270, flipH: true });
      const reloaded = createSpreadsheetStore();
      hydrateImagesFromEngine(reopened, reloaded, 0);
      expect(reloaded.getState().illustrations.illustrations[0]).toMatchObject({
        rotationDeg: 270,
        flipH: true,
      });

      history.undo();
      expect(wb.getImage(0, objectId)).toMatchObject({ rotationDeg: 90, flipH: false });
      expect(current().flipH).toBeUndefined();
      history.undo();
      expect(wb.getImage(0, objectId)).toMatchObject({ rotationDeg: 0 });
      history.redo();
      history.redo();
      expect(wb.getImage(0, objectId)).toMatchObject({ rotationDeg: 270, flipH: true });
    });
  },
);

describe('a loaded picture', () => {
  it('reads the rotation as the file stores it, negative or past a full turn', async () => {
    for (const rotationDeg of [-90, 400]) {
      const { store } = await picture('oneCell', (wb, id) => {
        wb.updatePicture(0, id, { rotationDeg, flipV: true });
      });
      expect(store.getState().illustrations.illustrations[0]).toMatchObject({
        rotationDeg,
        flipV: true,
      });
    }
  });

  it('keeps an unrotated picture without rotation or flips', async () => {
    const { current } = await picture('oneCell');
    expect(current().rotationDeg).toBeUndefined();
    expect(current().flipH).toBeUndefined();
  });

  it('restores both the frame and the turn when a move and a rotation undo together', async () => {
    const { wb, store, history, objectId, current } = await picture('twoCell');
    const before = wb.listDrawingObjects(0).find((o) => o.objectId === objectId)?.anchor;

    updateSessionIllustration(
      store,
      current().id,
      { x: 300, y: 200, rotationDeg: 45 },
      history,
      wb,
    );
    expect(wb.getImage(0, objectId)?.rotationDeg).toBe(45);
    expect(wb.listDrawingObjects(0).find((o) => o.objectId === objectId)?.anchor).not.toEqual(
      before,
    );

    history.undo();
    expect(wb.getImage(0, objectId)?.rotationDeg).toBe(0);
    expect(wb.listDrawingObjects(0).find((o) => o.objectId === objectId)?.anchor).toEqual(before);
    history.redo();
    expect(wb.getImage(0, objectId)?.rotationDeg).toBe(45);
  });

  it('still refuses to move or resize a picture fixed in place', async () => {
    const { wb, store, history, objectId, current } = await picture('absolute');
    const before = wb.listDrawingObjects(0).find((o) => o.objectId === objectId)?.anchor;

    expect(updateSessionIllustration(store, current().id, { x: 300 }, history, wb)).toBeNull();
    expect(wb.listDrawingObjects(0).find((o) => o.objectId === objectId)?.anchor).toEqual(before);
  });
});
