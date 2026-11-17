import { AnchorEditAs, AnchorKind } from '@libraz/formulon';
import { afterEach, describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { updateSessionIllustration } from '../../../src/commands/session-illustration.js';
import { hydrateImagesFromEngine } from '../../../src/engine/images-sync.js';
import { shapeFromNative } from '../../../src/engine/shapes-sync.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { attachSessionIllustrations } from '../../../src/interact/session-illustrations.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';
import { zipEntries, zipStored, zipText } from '../engine/zip-entries.js';

const cleanup: (() => void)[] = [];
afterEach(() => {
  for (const dispose of cleanup.splice(0).reverse()) dispose();
});

/** A workbook holding one rect shape turned by `rotationDeg`, drawn into a host. */
async function drawn(rotationDeg: number, flips: { flipH?: boolean; flipV?: boolean } = {}) {
  const wb = await WorkbookHandle.createDefault();
  cleanup.push(() => wb.dispose());
  wb.insertShape(0, {
    name: 'R',
    presetGeometry: 'rect',
    rotationDeg,
    ...flips,
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

    expect(shapeFromNative(store, 0, { ...shape, rotationDeg: -90 }, wb)?.rotationDeg).toBe(-90);
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

  it('draws a flipped shape mirrored inside its turned panel, with the text unmirrored', async () => {
    const { panel, item } = await drawn(330, { flipH: true });

    expect(item.flipH).toBe(true);
    expect(item.flipV).toBeUndefined();
    expect(panel.style.transform).toBe('rotate(330deg)');
    const group = panel.querySelector('svg > g');
    expect(group?.getAttribute('transform')).toMatch(/^translate\([\d.]+ 0\) scale\(-1 1\)$/);
    expect(group?.querySelector('rect')).not.toBeNull();
    expect(group?.querySelector('text')).toBeNull();
    expect(panel.querySelector('svg text')).not.toBeNull();
  });

  it('mirrors a vertically flipped shape about its horizontal axis', async () => {
    const { panel } = await drawn(0, { flipV: true });
    expect(panel.querySelector('svg > g')?.getAttribute('transform')).toMatch(
      /^translate\(0 [\d.]+\) scale\(1 -1\)$/,
    );
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

describe('a picture loaded with a rotation and a flip', () => {
  it('is drawn turned and mirrored', async () => {
    const wb = await WorkbookHandle.createDefault();
    cleanup.push(() => wb.dispose());
    const png = Uint8Array.from(
      atob(
        'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=',
      ),
      (char) => char.charCodeAt(0),
    );
    const id = wb.insertImage(0, png, {
      anchor: { fromRow: 2, fromCol: 1, widthEmu: 952500, heightEmu: 952500 },
    });
    wb.updatePicture(0, id ?? 0, { rotationDeg: 330, flipH: true });
    const store = createSpreadsheetStore();
    hydrateImagesFromEngine(wb, store, 0);
    const host = document.createElement('div');
    document.body.appendChild(host);
    cleanup.push(() => host.remove());
    const handle = attachSessionIllustrations({
      host,
      store,
      history: new History(),
      getWb: () => wb,
    });
    cleanup.push(() => handle.detach());

    const panel = host.querySelector<HTMLElement>('.fc-illustration');
    expect(panel?.style.transform).toBe('rotate(330deg)');
    expect(panel?.querySelector('img')?.style.transform).toBe('scale(-1, 1)');
  });
});
