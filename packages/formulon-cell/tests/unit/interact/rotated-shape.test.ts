import { crc32 } from 'node:zlib';
import { AnchorEditAs, AnchorKind } from '@libraz/formulon';
import { afterEach, describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { updateSessionIllustration } from '../../../src/commands/session-illustration.js';
import { hydrateImagesFromEngine } from '../../../src/engine/images-sync.js';
import { shapeFromNative } from '../../../src/engine/shapes-sync.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { attachSessionIllustrations } from '../../../src/interact/session-illustrations.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';
import { zipEntries, zipText } from '../engine/zip-entries.js';

const cleanup: (() => void)[] = [];
afterEach(() => {
  for (const dispose of cleanup.splice(0).reverse()) dispose();
});

/** A zip of stored entries. */
function zipStored(entries: Map<string, Uint8Array>): Uint8Array {
  const parts: Buffer[] = [];
  const central: Buffer[] = [];
  let offset = 0;
  for (const [name, data] of entries) {
    const nameBytes = Buffer.from(name);
    const head = Buffer.alloc(30);
    head.writeUInt32LE(0x04034b50, 0);
    head.writeUInt16LE(20, 4);
    head.writeUInt32LE(crc32(data), 14);
    head.writeUInt32LE(data.length, 18);
    head.writeUInt32LE(data.length, 22);
    head.writeUInt16LE(nameBytes.length, 26);
    const entry = Buffer.alloc(46);
    entry.writeUInt32LE(0x02014b50, 0);
    entry.writeUInt16LE(20, 4);
    entry.writeUInt16LE(20, 6);
    entry.writeUInt32LE(crc32(data), 16);
    entry.writeUInt32LE(data.length, 20);
    entry.writeUInt32LE(data.length, 24);
    entry.writeUInt16LE(nameBytes.length, 28);
    entry.writeUInt32LE(offset, 42);
    central.push(Buffer.concat([entry, nameBytes]));
    parts.push(head, nameBytes, Buffer.from(data));
    offset += 30 + nameBytes.length + data.length;
  }
  const directory = Buffer.concat(central);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  end.writeUInt16LE(entries.size, 8);
  end.writeUInt16LE(entries.size, 10);
  end.writeUInt32LE(directory.length, 12);
  end.writeUInt32LE(offset, 16);
  return Buffer.concat([...parts, directory, end]);
}

/** A workbook holding one rect shape turned by `rotationDeg`, drawn into a host. */
async function drawn(rotationDeg: number) {
  const wb = await WorkbookHandle.createDefault();
  cleanup.push(() => wb.dispose());
  wb.insertShape(0, {
    name: 'R',
    presetGeometry: 'rect',
    rotationDeg,
    text: 'hi',
    anchor: {
      kind: AnchorKind.TwoCell,
      editAs: AnchorEditAs.TwoCell,
      fromRow: 2,
      fromCol: 1,
      toRow: 6,
      toCol: 4,
    },
  });
  const store = createSpreadsheetStore();
  hydrateImagesFromEngine(wb, store, 0);
  wb.attachStore(store);
  const history = new History();
  wb.attachHistory(history);
  const host = document.createElement('div');
  document.body.appendChild(host);
  cleanup.push(() => host.remove());
  const handle = attachSessionIllustrations({ host, store, history, getWb: () => wb });
  cleanup.push(() => handle.detach());
  const panel = host.querySelector<HTMLElement>('.fc-illustration');
  if (!panel) throw new Error('no panel');
  const item = store.getState().illustrations.illustrations[0];
  if (!item) throw new Error('no item');
  return { wb, store, history, host, panel, item };
}

describe('a shape loaded with a rotation', () => {
  it('turns its panel about the box centre', async () => {
    const { panel, item } = await drawn(30);

    expect(item.nativeLocked).not.toBe(true);
    expect(panel.style.transform).toBe('rotate(30deg)');
    expect(panel.style.transformOrigin).toBe('50% 50%');
  });

  it('carries the shape, its text, its handles and the rotation grip inside the turned panel', async () => {
    const { panel } = await drawn(30);

    for (const part of ['svg', 'svg text', '.fc-drawing-handle', '.fc-drawing-rotate']) {
      expect(panel.querySelector(part), part).not.toBeNull();
    }
    // The panel is the only hit target, so the browser tests clicks and hover
    // against the turned outline and not the unrotated box.
    expect(panel.style.pointerEvents).toBe('auto');
    expect(panel.parentElement?.style.transform).toBe('');
  });

  it('draws a negative or full-turn rotation as its angle in [0, 360)', async () => {
    expect((await drawn(-90)).panel.style.transform).toBe('rotate(270deg)');
    expect((await drawn(400)).panel.style.transform).toBe('rotate(40deg)');
    expect((await drawn(360)).panel.style.transform).toBe('');
  });

  it('keeps the rotation the engine reports, un-normalised', async () => {
    const { wb, store } = await drawn(30);
    const shape = wb.listShapes(0)[0];
    if (!shape) throw new Error('no shape');

    expect(shapeFromNative(store, 0, { ...shape, rotationDeg: -90 })?.rotationDeg).toBe(-90);
  });

  it('leaves the stored rotation alone when the shape moves', async () => {
    const { wb, store, history, item } = await drawn(-90);
    const before = wb.listShapes(0)[0]?.rotationDeg;

    expect(item.rotationDeg).toBe(before);
    updateSessionIllustration(store, item.id, { x: 200, y: 120 }, history, wb);
    expect(wb.listShapes(0)[0]?.rotationDeg).toBe(before);

    const reopened = await WorkbookHandle.loadBytes(wb.save());
    cleanup.push(() => reopened.dispose());
    expect(reopened.listShapes(0)[0]?.rotationDeg).toBe(before);
  });

  it('leaves an unrotated shape without a transform', async () => {
    const { panel, item } = await drawn(0);

    expect(panel.style.transform).toBe('');
    expect(item.rotationDeg).toBeUndefined();
  });

  it('loads a stored negative rotation as it is and saves it back unchanged', async () => {
    const { wb } = await drawn(30);
    const entries = zipEntries(wb.save());
    const path = [...entries.keys()].find((name) => /^xl\/drawings\/drawing\d+\.xml$/.test(name));
    if (!path) throw new Error('no drawing part');
    const xml = zipText(entries, path);
    expect(xml).toMatch(/<a:xfrm[^>]*rot="1800000"/);
    entries.set(path, new TextEncoder().encode(xml.replace(/rot="1800000"/, 'rot="-5400000"')));

    const loaded = await WorkbookHandle.loadBytes(zipStored(entries));
    cleanup.push(() => loaded.dispose());
    expect(loaded.listShapes(0)[0]?.rotationDeg).toBe(-90);

    const store = createSpreadsheetStore();
    hydrateImagesFromEngine(loaded, store, 0);
    expect(store.getState().illustrations.illustrations[0]?.rotationDeg).toBe(-90);

    const again = zipEntries(loaded.save());
    expect(zipText(again, path)).toMatch(/<a:xfrm[^>]*rot="-5400000"/);
  });
});
