/**
 * Shapes Excel saved with theme colours, opened in the cell. The expectations
 * are the colours Excel reports for each shape, not the adapter's own output.
 */
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import type { ChartFill, ChartLine } from '@libraz/formulon';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { History } from '../../../src/commands/history.js';
import {
  clearSessionIllustration,
  copySessionIllustration,
  duplicateSessionIllustration,
  updateSessionIllustration,
} from '../../../src/commands/session-illustration.js';
import {
  applySessionShapeFill,
  applySessionShapeOutline,
  applySessionShapeOutlineWeight,
} from '../../../src/commands/shape-format.js';
import { addSheet, moveSheet } from '../../../src/commands/sheet-mutate.js';
import { hydrateImagesFromEngine } from '../../../src/engine/images-sync.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { en } from '../../../src/i18n/strings.js';
import { attachClipboard } from '../../../src/interact/clipboard.js';
import { attachSessionIllustrations } from '../../../src/interact/session-illustrations.js';
import { attachWorkbookObjectsPanel } from '../../../src/interact/workbook-objects.js';
import { createShapeFormatDropdownDefaults } from '../../../src/mount/dynamic-dropdown-defaults/shape-format.js';
import type { SpreadsheetInstance } from '../../../src/mount.js';
import {
  createSpreadsheetStore,
  mutators,
  type SessionIllustration,
  type SpreadsheetStore,
} from '../../../src/store/store.js';
import { zipEntries } from './zip-entries.js';

const FIXTURE = join(__dirname, '../../fixtures/shapes/excel-shape-theme-colors.xlsx');

const open: WorkbookHandle[] = [];
afterEach(() => {
  for (const wb of open.splice(0)) wb.dispose();
});

async function session(bytes: Uint8Array) {
  const wb = await WorkbookHandle.loadBytes(bytes);
  open.push(wb);
  const store = createSpreadsheetStore();
  hydrateImagesFromEngine(wb, store, 0);
  wb.attachStore(store);
  const history = new History();
  wb.attachHistory(history);
  return { wb, store, history };
}

const named = (store: SpreadsheetStore, name: string): SessionIllustration => {
  const item = store.getState().illustrations.illustrations.find((entry) => entry.name === name);
  if (!item) throw new Error(`no shape ${name}`);
  return item;
};

/** A paint as the engine holds it: colour spec (tint at its stored precision) and transparency. */
const paint = (value: ChartFill | ChartLine | null | undefined) => ({
  kind: value?.color?.kind,
  theme: value?.color?.theme,
  tint: Math.round((value?.color?.tint ?? 0) * 10_000) / 10_000,
  rgb: (value?.color?.argb ?? 0) & 0xffffff,
  transparencyPct: value?.transparencyPct ?? null,
});

const paints = (wb: WorkbookHandle) =>
  wb.listShapes(0).map((shape) => ({
    name: shape.name,
    fill: paint(shape.fill),
    line: paint(shape.line),
  }));

const spec = (item: SessionIllustration['nativeFill']) => ({
  theme: item?.spec.theme,
  tint: Math.round((item?.spec.tint ?? 0) * 10_000) / 10_000,
});

describe('shapes Excel saved with theme colours', () => {
  const bytes = new Uint8Array(readFileSync(FIXTURE));

  it('shows a lumMod/lumOff accent shape as Excel does', async () => {
    const { store } = await session(bytes);
    const box = named(store, 'Accent Box');

    expect(box).toMatchObject({
      color: '#f2aa84',
      opacity: 0.7,
      lineColor: '#104862',
      lineWidth: 2,
      lineDash: 'dash',
      alt: 'alt text here',
    });
    expect(spec(box.nativeFill)).toEqual({ theme: 5, tint: 0.4 });
    expect(spec(box.nativeLine)).toEqual({ theme: 4, tint: -0.25 });
  });

  it('loads a shape whose outline it cannot state, unlocked', async () => {
    const { store } = await session(bytes);
    const note = named(store, 'Note Box');

    // The outline is lt1 with a shade the engine does not render yet, so its colour is not asserted.
    expect(note).toMatchObject({ textBox: true, color: '#ffffff', lineWidth: 0.75 });
    expect(note.nativeLocked).toBeUndefined();
  });

  it('writes the same colour specs and transparency back', async () => {
    const { wb } = await session(bytes);
    const loaded = paints(wb);

    const reopened = await session(wb.save());

    expect(paints(reopened.wb)).toEqual(loaded);
  });

  it('keeps the fill spec when only the outline weight changes', async () => {
    const { wb, store, history } = await session(bytes);
    const fill = paints(wb).find((shape) => shape.name === 'Accent Box')?.fill;

    applySessionShapeOutlineWeight(store, named(store, 'Accent Box'), 3, history, wb);
    const reopened = await session(wb.save());

    const after = reopened.wb.listShapes(0).find((shape) => shape.name === 'Accent Box');
    expect(paint(after?.fill)).toEqual(fill);
    expect(after?.line?.widthPt).toBe(3);
  });
});

describe('a shape whose outline the cell cannot state', () => {
  const bytes = new Uint8Array(readFileSync(FIXTURE));

  /** Note Box's `a:ln` in the saved drawing part, and the colour element inside it. */
  const noteLine = (wb: WorkbookHandle) => {
    const xml = new TextDecoder().decode(zipEntries(wb.save()).get('xl/drawings/drawing1.xml'));
    const shape = xml.slice(xml.indexOf('name="Note Box"'));
    const ln = shape.match(/<a:ln\b[^>]*>.*?<\/a:ln>/)?.[0] ?? '';
    return { ln, fill: ln.match(/<a:solidFill>.*?<\/a:solidFill>/)?.[0] ?? '' };
  };
  const SHADED =
    '<a:solidFill><a:schemeClr val="lt1"><a:shade val="50000"/></a:schemeClr></a:solidFill>';

  /** Note Box as the engine reads it after a save and reload. */
  const reloadedLine = async (wb: WorkbookHandle) =>
    (await session(wb.save())).wb.listShapes(0).find((shape) => shape.name === 'Note Box')?.line;

  it('greys the outline controls and keeps the fill controls', async () => {
    const { wb, store, history } = await session(bytes);
    mutators.setSelectedObject(store, { id: named(store, 'Note Box').id, name: '' });
    const instance = {
      host: document.createElement('div'),
      store,
      history,
      workbook: wb,
      i18n: { locale: 'en', strings: en },
      isSheetProtected: () => false,
    } as unknown as SpreadsheetInstance;
    const menu = (attr: string): HTMLElement => {
      const element = document.createElement('div');
      element.innerHTML = `<button ${attr}="#ff0000"></button><button ${attr}="none"></button>
        <button data-shape-outline-weight="1"></button><button data-shape-outline-dash="dash"></button>`;
      return element;
    };
    const disabled = (button: Element | null): boolean =>
      (button as HTMLButtonElement | null)?.disabled === true ||
      button?.getAttribute('aria-disabled') === 'true';
    const defaults = createShapeFormatDropdownDefaults(instance);

    const fill = menu('data-shape-fill');
    defaults.updateShapeFillMenu(fill);
    const outline = menu('data-shape-outline');
    defaults.updateShapeOutlineMenu(outline);

    expect(disabled(fill.querySelector('[data-shape-fill="#ff0000"]'))).toBe(false);
    for (const button of outline.querySelectorAll('button')) expect(disabled(button)).toBe(true);
    expect(
      applySessionShapeOutline(store, named(store, 'Note Box'), '#00ff00', history, wb),
    ).toBeNull();
  });

  it('keeps the outline paint through a fill restyle', async () => {
    const { wb, store, history } = await session(bytes);

    applySessionShapeFill(store, named(store, 'Note Box'), '#ff0000', history, wb);

    expect(noteLine(wb).fill).toBe(SHADED);
    expect(await reloadedLine(wb)).toMatchObject({ other: true, widthPt: 0.75 });
  });

  it('keeps the outline colour through a weight or dash change', async () => {
    const { wb, store, history } = await session(bytes);
    const id = named(store, 'Note Box').id;

    updateSessionIllustration(store, id, { lineWidth: 2 }, history, wb);
    expect(noteLine(wb).fill).toBe(SHADED);
    expect(await reloadedLine(wb)).toMatchObject({ other: true, widthPt: 2 });

    updateSessionIllustration(store, id, { lineDash: 'dash' }, history, wb);
    expect(noteLine(wb).fill).toBe(SHADED);
    expect(noteLine(wb).ln).toContain('<a:prstDash val="dash"/>');
  });

  it('keeps the outline colour through the objects dialog', async () => {
    const { wb, store, history } = await session(bytes);
    const host = document.createElement('div');
    const handle = attachWorkbookObjectsPanel({
      getLocale: () => 'en-US',
      store,
      host,
      wb,
      strings: en,
      listSessionIllustrations: () => store.getState().illustrations.illustrations,
      onUpdateSessionIllustration: (id, patch) =>
        updateSessionIllustration(store, id, patch, history, wb),
    });
    handle.open();
    const form = Array.from(
      host.querySelectorAll<HTMLFormElement>('.fc-objects__illustration-edit'),
    )[1];
    const lineWidth = form?.querySelectorAll<HTMLInputElement>('input[type="number"]')[1];
    expect(lineWidth?.disabled).toBe(true);

    form?.dispatchEvent(new Event('submit', { bubbles: true, cancelable: true }));
    handle.detach();

    expect(noteLine(wb).fill).toBe(SHADED);
    expect(await reloadedLine(wb)).toMatchObject({ other: true, widthPt: 0.75 });
  });

  it('keeps the outline paint when its sheet is reordered, and through a delete undone', async () => {
    const { wb, store, history } = await session(bytes);
    const objectId = named(store, 'Note Box').nativeObjectId;
    addSheet(store, wb, history);

    expect(moveSheet(store, wb, 0, 1, history)).toBe(true);
    expect(named(store, 'Note Box')).toMatchObject({ sheet: 1, nativeObjectId: objectId });
    expect(wb.listShapes(1).find((shape) => shape.name === 'Note Box')?.line?.other).toBe(true);

    expect(moveSheet(store, wb, 1, 0, history)).toBe(true);
    clearSessionIllustration(store, named(store, 'Note Box').id, history, wb);
    history.undo();
    expect(noteLine(wb).fill).toBe(SHADED);
  });

  it('refuses a copy or a move to another sheet, and says why', async () => {
    const { wb, store, history } = await session(bytes);
    const note = named(store, 'Note Box');

    expect(duplicateSessionIllustration(store, note.id, history, wb)).toBeNull();
    expect(copySessionIllustration(store, note.id, wb)).toBeNull();
    const second = wb.addSheet();
    updateSessionIllustration(store, note.id, { sheet: second }, history, wb);
    expect(named(store, 'Note Box').sheet).toBe(0);
    expect(noteLine(wb).fill).toBe(SHADED);

    const host = document.createElement('div');
    const handle = attachWorkbookObjectsPanel({
      getLocale: () => 'en-US',
      store,
      host,
      wb,
      strings: en,
      listSessionIllustrations: () => store.getState().illustrations.illustrations,
      onDuplicateSessionIllustration: () => null,
    });
    handle.open();
    const duplicate = Array.from(
      host.querySelectorAll<HTMLButtonElement>('.fc-objects__illustration-actions button'),
    ).filter((button) => button.textContent === en.workbookObjects.objectDuplicate)[1];
    expect(duplicate?.disabled || duplicate?.getAttribute('aria-disabled') === 'true').toBe(true);
    expect(host.textContent + (duplicate?.title ?? '')).toContain(
      en.workbookObjects.shapePaintKept,
    );
    handle.detach();
  });

  it('says why a keyboard duplicate, copy or cut is refused, and leaves the cells alone', async () => {
    const { wb, store, history } = await session(bytes);
    const note = named(store, 'Note Box');
    const host = document.createElement('div');
    document.body.appendChild(host);
    let systemText = 'Earlier text';
    vi.stubGlobal('navigator', {
      clipboard: {
        readText: async () => systemText,
        writeText: async (text: string) => {
          systemText = text;
        },
      },
    });
    const notice = () => {
      const dialog = document.querySelector('[role="alertdialog"]');
      const text = dialog?.querySelector('.fc-tb__dlg__message')?.textContent;
      dialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
      expect(document.querySelector('[role="alertdialog"]')).toBeNull();
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
      await clipboard.runShortcut('copy');
      const snapshot = clipboard.getSnapshot();
      const copyRange = store.getState().ui.copyRange;
      expect(snapshot).not.toBeNull();
      const copiedText = systemText;
      const count = store.getState().illustrations.illustrations.length;
      // A copy that fell through to the cells would take this cell instead.
      mutators.setActive(store, { sheet: 0, row: 4, col: 3 });

      const panel = host.querySelector<HTMLElement>(`[data-illustration-id="${note.id}"]`);
      if (!panel) throw new Error('no Note Box panel');
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
      expect(notice()).toBe(en.workbookObjects.shapePaintKept);
      expect(store.getState().illustrations.illustrations).toHaveLength(count);

      for (const kind of ['copy', 'cut'] as const) {
        await clipboard.runShortcut(kind);
        expect(notice()).toBe(en.workbookObjects.shapePaintKept);
        expect(clipboard.getSnapshot()).toEqual(snapshot);
        expect(store.getState().ui.copyRange).toEqual(copyRange);
        expect(systemText).toBe(copiedText);
      }
      expect(named(store, 'Note Box').id).toBe(note.id);
    } finally {
      clipboard.detach();
      illustrations.detach();
      vi.unstubAllGlobals();
      host.remove();
    }
  });
});
