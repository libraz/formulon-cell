import { AnchorEditAs, AnchorKind } from '@libraz/formulon';
import { afterEach, describe, expect, it } from 'vitest';
import { createDrawingClipboard } from '../../../src/commands/clipboard/drawing.js';
import { History } from '../../../src/commands/history.js';
import {
  clearSessionIllustration,
  copySessionIllustration,
  duplicateSessionIllustration,
  pasteSessionIllustration,
  updateSessionIllustration,
} from '../../../src/commands/session-illustration.js';
import { deleteRows } from '../../../src/commands/structure.js';
import { hydrateImagesFromEngine } from '../../../src/engine/images-sync.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import {
  createSpreadsheetStore,
  mutators,
  type SessionIllustration,
} from '../../../src/store/store.js';
import { zipEntries, zipStored } from './zip-entries.js';

const open: WorkbookHandle[] = [];
afterEach(() => {
  for (const wb of open.splice(0)) wb.dispose();
});

// 1x1 transparent PNG.
const PNG = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=',
  ),
  (char) => char.charCodeAt(0),
);

/** The file's `rot` attribute for each raw rotation the tests keep: -90 and 400 degrees. */
const STORED_ROT = new Map([
  [-90, '-5400000'],
  [400, '24000000'],
]);

const bind = async (bytes: Uint8Array) => {
  const wb = await WorkbookHandle.loadBytes(bytes);
  open.push(wb);
  const store = createSpreadsheetStore();
  hydrateImagesFromEngine(wb, store, 0);
  wb.attachStore(store);
  const history = new History();
  wb.attachHistory(history);
  return { wb, store, history };
};

/** A workbook whose only shape is stored with the rotation `rotationDeg`, as a file can hold it. */
const shapeFile = async (rotationDeg: number) => {
  const source = await WorkbookHandle.createDefault();
  source.insertShape(0, {
    presetGeometry: 'rect',
    rotationDeg: 30,
    text: 'kept',
    anchor: {
      kind: AnchorKind.TwoCell,
      editAs: AnchorEditAs.TwoCell,
      fromRow: 10,
      fromCol: 2,
      toRow: 12,
      toCol: 4,
    },
  });
  const entries = zipEntries(source.save());
  source.dispose();
  const path = [...entries.keys()].find((name) => /^xl\/drawings\/drawing\d+\.xml$/.test(name));
  if (!path) throw new Error('no drawing part');
  const xml = new TextDecoder()
    .decode(entries.get(path))
    .replace('rot="1800000"', `rot="${STORED_ROT.get(rotationDeg)}"`);
  entries.set(path, new TextEncoder().encode(xml));
  return bind(zipStored(entries));
};

const rotations = (wb: WorkbookHandle, sheet = 0): number[] =>
  wb.listShapes(sheet).map((shape) => shape.rotationDeg);

const shapeItems = (store: ReturnType<typeof createSpreadsheetStore>): SessionIllustration[] =>
  store.getState().illustrations.illustrations.filter((item) => item.kind === 'shape');

describe.each([-90, 400])('a shape stored with a rotation of %i degrees', (rotationDeg) => {
  it('loads the rotation as stored', async () => {
    const { wb, store } = await shapeFile(rotationDeg);
    expect(rotations(wb)).toEqual([rotationDeg]);
    expect(shapeItems(store)[0]?.rotationDeg).toBe(rotationDeg);
  });

  it('keeps it through duplicate, paste and a reopen', async () => {
    const { wb, store, history } = await shapeFile(rotationDeg);
    const id = shapeItems(store)[0]?.id ?? '';

    duplicateSessionIllustration(store, id, history, wb);
    const clip = copySessionIllustration(store, id, wb);
    if (!clip) throw new Error('no clip');
    pasteSessionIllustration(store, clip, { sheet: 0, row: 20, col: 1 }, history, wb);

    expect(rotations(wb)).toEqual([rotationDeg, rotationDeg, rotationDeg]);
    expect(shapeItems(store).map((item) => item.rotationDeg)).toEqual([
      rotationDeg,
      rotationDeg,
      rotationDeg,
    ]);
    const reopened = await bind(wb.save());
    expect(rotations(reopened.wb)).toEqual([rotationDeg, rotationDeg, rotationDeg]);
  });

  it('keeps it through cut-paste and a sheet move', async () => {
    const { wb, store, history } = await shapeFile(rotationDeg);
    const clipboard = createDrawingClipboard({ store, wb, history });
    mutators.setSelectedObject(store, { id: shapeItems(store)[0]?.id ?? '', name: '' });

    expect(clipboard.take('cut')).not.toBeNull();
    expect(clipboard.paste()).toBe(true);
    expect(rotations(wb)).toEqual([rotationDeg]);

    const second = wb.addSheet();
    updateSessionIllustration(
      store,
      shapeItems(store)[0]?.id ?? '',
      { sheet: second },
      history,
      wb,
    );
    expect(rotations(wb, second)).toEqual([rotationDeg]);
    expect(shapeItems(store)[0]?.rotationDeg).toBe(rotationDeg);
  });

  it('keeps it when a delete is undone', async () => {
    const { wb, store, history } = await shapeFile(rotationDeg);
    const id = shapeItems(store)[0]?.id ?? '';

    expect(clearSessionIllustration(store, id, history, wb)).toBe(true);
    expect(rotations(wb)).toEqual([]);
    history.undo();
    expect(rotations(wb)).toEqual([rotationDeg]);
    expect(shapeItems(store)[0]?.rotationDeg).toBe(rotationDeg);

    expect(deleteRows(store, wb, history, 9, 5)).toBe(true);
    history.undo();
    expect(rotations(wb)).toEqual([rotationDeg]);
  });
});

describe('a picture stored with a rotation of -90 degrees', () => {
  it('keeps it through duplicate and paste', async () => {
    const source = await WorkbookHandle.createDefault();
    const objectId =
      source.insertImage(0, PNG, {
        anchor: { fromRow: 4, fromCol: 2, widthEmu: 952500, heightEmu: 952500 },
      }) ?? 0;
    source.updatePicture(0, objectId, { rotationDeg: -90 });
    const bytes = source.save();
    source.dispose();
    const { wb, store, history } = await bind(bytes);
    const id = `image-native-0-${objectId}`;

    duplicateSessionIllustration(store, id, history, wb);
    const clip = copySessionIllustration(store, id, wb);
    if (!clip) throw new Error('no clip');
    pasteSessionIllustration(store, clip, { sheet: 0, row: 20, col: 1 }, history, wb);

    const pictures = store.getState().illustrations.illustrations.filter((i) => i.kind === 'image');
    expect(pictures).toHaveLength(3);
    for (const picture of pictures) {
      expect(picture.rotationDeg).toBe(-90);
      expect(wb.getImage(0, picture.nativeObjectId ?? -1)?.rotationDeg).toBe(-90);
    }
  });
});
