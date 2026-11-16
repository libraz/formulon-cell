import { AnchorKind } from '@libraz/formulon';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { History } from '../../../src/commands/history.js';
import {
  clearSessionIllustration,
  updateSessionIllustration,
} from '../../../src/commands/session-illustration.js';
import {
  applySessionShapeKind,
  applySessionShapeRotation,
} from '../../../src/commands/shape-format.js';
import { hydrateImagesFromEngine } from '../../../src/engine/images-sync.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { en } from '../../../src/i18n/strings.js';
import { attachClipboard } from '../../../src/interact/clipboard.js';
import { attachSessionIllustrations } from '../../../src/interact/session-illustrations.js';
import { attachWorkbookObjectsPanel } from '../../../src/interact/workbook-objects.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';
import { zipEntries, zipStored } from '../engine/zip-entries.js';

// 1x1 transparent PNG.
const PNG = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=',
  ),
  (char) => char.charCodeAt(0),
);

const cleanup: (() => void)[] = [];
afterEach(() => {
  for (const dispose of cleanup.splice(0).reverse()) dispose();
  vi.unstubAllGlobals();
});

/** A workbook holding one picture or shape placed at a fixed position on the sheet (an absolute
 *  anchor), as a file can hold one. */
async function fixedDrawing(kind: 'picture' | 'shape') {
  const source = await WorkbookHandle.createDefault();
  const anchor = {
    kind: AnchorKind.OneCell,
    fromRow: 4,
    fromCol: 2,
    widthEmu: 952500,
    heightEmu: 952500,
  };
  if (kind === 'picture') source.insertImage(0, PNG, { anchor, name: 'Logo' });
  else source.insertShape(0, { name: 'Box', presetGeometry: 'rect', anchor });
  const entries = zipEntries(source.save());
  source.dispose();
  const path = 'xl/drawings/drawing1.xml';
  const xml = new TextDecoder()
    .decode(entries.get(path))
    .replace(
      /<xdr:oneCellAnchor><xdr:from>.*?<\/xdr:from>/,
      '<xdr:absoluteAnchor><xdr:pos x="1219200" y="762000"/>',
    )
    .replace('</xdr:oneCellAnchor>', '</xdr:absoluteAnchor>');
  entries.set(path, new TextEncoder().encode(xml));
  const wb = await WorkbookHandle.loadBytes(zipStored(entries));
  cleanup.push(() => wb.dispose());
  const store = createSpreadsheetStore();
  hydrateImagesFromEngine(wb, store, 0);
  wb.attachStore(store);
  const history = new History();
  wb.attachHistory(history);
  const [item] = store.getState().illustrations.illustrations;
  if (!item) throw new Error(`no ${kind}`);
  expect(item).toMatchObject({
    kind: kind === 'picture' ? 'image' : 'shape',
    nativeGeometryEditable: false,
  });
  return { wb, store, history, item };
}

/** The drawing's anchor as the engine reads it. */
const anchors = (wb: WorkbookHandle) =>
  wb.listDrawingObjects(0).map(({ anchor }) => ({
    kind: anchor.kind,
    x: anchor.fromColOffsetEmu,
    y: anchor.fromRowOffsetEmu,
    widthEmu: anchor.widthEmu,
    heightEmu: anchor.heightEmu,
  }));
const FIXED = {
  kind: AnchorKind.Absolute,
  x: 1219200,
  y: 762000,
  widthEmu: 952500,
  heightEmu: 952500,
};

/** The open refusal notice's message, closed with its OK button. */
const notice = (): string | null | undefined => {
  const dialog = document.querySelector('[role="alertdialog"]');
  const text = dialog?.querySelector('.fc-tb__dlg__message')?.textContent;
  dialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
  expect(document.querySelector('[role="alertdialog"]')).toBeNull();
  return text;
};

describe.each(['picture', 'shape'] as const)(
  'a %s placed at a fixed position on the sheet',
  (kind) => {
    it('says why a keyboard duplicate, copy or cut is refused, and leaves the cells alone', async () => {
      const { wb, store, history, item } = await fixedDrawing(kind);

      const host = document.createElement('div');
      document.body.appendChild(host);
      cleanup.push(() => host.remove());
      let systemText = 'Earlier text';
      vi.stubGlobal('navigator', {
        clipboard: {
          readText: async () => systemText,
          writeText: async (text: string) => {
            systemText = text;
          },
        },
      });
      const illustrations = attachSessionIllustrations({
        host,
        store,
        history,
        getWb: () => wb,
        getStrings: () => en,
      });
      cleanup.push(() => illustrations.detach());
      const clipboard = attachClipboard({
        host,
        store,
        wb,
        history,
        onAfterCommit: () => {},
        getStrings: () => en,
      });
      cleanup.push(() => clipboard.detach());

      await clipboard.runShortcut('copy');
      const snapshot = clipboard.getSnapshot();
      const copyRange = store.getState().ui.copyRange;
      const copiedText = systemText;
      expect(snapshot).not.toBeNull();
      // A copy that fell through to the cells would take this cell instead.
      mutators.setActive(store, { sheet: 0, row: 4, col: 3 });

      const panel = host.querySelector<HTMLElement>(`[data-illustration-id="${item.id}"]`);
      if (!panel) throw new Error('no panel');
      panel.dispatchEvent(new PointerEvent('pointerdown', { bubbles: true, button: 0 }));
      window.dispatchEvent(new PointerEvent('pointerup', { bubbles: true }));
      const key = new KeyboardEvent('keydown', {
        bubbles: true,
        cancelable: true,
        key: 'd',
        ctrlKey: true,
      });
      panel.dispatchEvent(key);
      expect(key.defaultPrevented).toBe(true);
      expect(notice()).toBe(en.workbookObjects.drawingFixedPosition);
      expect(store.getState().illustrations.illustrations).toHaveLength(1);

      for (const kind of ['copy', 'cut'] as const) {
        await clipboard.runShortcut(kind);
        expect(notice()).toBe(en.workbookObjects.drawingFixedPosition);
        expect(clipboard.getSnapshot()).toEqual(snapshot);
        expect(store.getState().ui.copyRange).toEqual(copyRange);
        expect(systemText).toBe(copiedText);
      }
      expect(store.getState().illustrations.illustrations).toHaveLength(1);
      expect(wb.listDrawingObjects(0)).toHaveLength(1);
    });

    it('greys Duplicate with the reason in the objects dialog, and leaves Delete on', async () => {
      const { wb, store } = await fixedDrawing(kind);
      const host = document.createElement('div');
      const panel = attachWorkbookObjectsPanel({
        getLocale: () => 'en-US',
        store,
        host,
        wb,
        strings: en,
        listSessionIllustrations: () => store.getState().illustrations.illustrations,
        onDuplicateSessionIllustration: () => null,
        onClearSessionIllustration: () => {},
      });
      cleanup.push(() => panel.detach());
      panel.open();
      const buttons = Array.from(
        host.querySelectorAll<HTMLButtonElement>('.fc-objects__illustration-actions button'),
      ).filter((button) =>
        [en.workbookObjects.objectDuplicate, en.workbookObjects.objectDelete].includes(
          button.textContent ?? '',
        ),
      );
      expect(buttons.map((button) => button.textContent)).toEqual([
        en.workbookObjects.objectDuplicate,
        en.workbookObjects.objectDelete,
      ]);
      const [duplicate, remove] = buttons;
      expect(duplicate?.disabled || duplicate?.getAttribute('aria-disabled') === 'true').toBe(true);
      expect(duplicate?.title).toContain(en.workbookObjects.drawingFixedPosition);
      expect(remove?.disabled || remove?.getAttribute('aria-disabled') === 'true').toBe(false);
    });

    it('refuses a move to another sheet', async () => {
      const { wb, store, history, item } = await fixedDrawing(kind);
      const second = wb.addSheet();

      expect(updateSessionIllustration(store, item.id, { sheet: second }, history, wb)).toBeNull();
      expect(store.getState().illustrations.illustrations[0]?.sheet).toBe(0);
      expect(anchors(wb)).toEqual([FIXED]);
    });

    it('deletes, and an undo puts it back where it was', async () => {
      const { wb, store, history, item } = await fixedDrawing(kind);
      expect(anchors(wb)).toEqual([FIXED]);

      expect(clearSessionIllustration(store, item.id, history, wb)).toBe(true);
      expect(anchors(wb)).toEqual([]);
      expect(store.getState().illustrations.illustrations).toHaveLength(0);

      history.undo();
      expect(anchors(wb)).toEqual([FIXED]);
      expect(store.getState().illustrations.illustrations).toHaveLength(1);
    });
  },
);

describe('a shape placed at a fixed position on the sheet', () => {
  it('changes kind and turns in place, since neither moves its anchor', async () => {
    const { wb, store, history, item } = await fixedDrawing('shape');

    expect(applySessionShapeKind(store, item, 'oval', history, wb)).not.toBeNull();
    const kinded = store.getState().illustrations.illustrations[0];
    if (!kinded) throw new Error('shape is gone');
    expect(applySessionShapeRotation(store, kinded, 90, history, wb)).not.toBeNull();

    const [shape] = wb.listShapes(0);
    expect(shape).toMatchObject({ presetGeometry: 'ellipse', rotationDeg: 90 });
    expect(anchors(wb)).toEqual([FIXED]);
  });
});
