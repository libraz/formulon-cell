import { ChartLineEndType } from '@libraz/formulon';
import { afterEach, describe, expect, it } from 'vitest';
import { History } from '../../../../src/commands/history.js';
import {
  createRibbonImageFromSelection,
  createSessionShape,
} from '../../../../src/commands/session-illustration.js';
import { hydrateImagesFromEngine } from '../../../../src/engine/images-sync.js';
import { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { en, ja } from '../../../../src/i18n/strings.js';
import { attachSessionIllustrations } from '../../../../src/interact/session-illustrations.js';
import { createIllustrationDropdownDefaults } from '../../../../src/mount/dynamic-dropdown-defaults/illustrations.js';
import { createShapeFormatDropdownDefaults } from '../../../../src/mount/dynamic-dropdown-defaults/shape-format.js';
import type { SpreadsheetInstance } from '../../../../src/mount.js';
import { createSpreadsheetStore, mutators } from '../../../../src/store/store.js';
import { projectMacRibbonState } from '../../../../src/toolbar/ribbon/mac/model.js';
import { zipEntries } from '../../engine/zip-entries.js';

const open: WorkbookHandle[] = [];
afterEach(() => {
  for (const wb of open.splice(0)) wb.dispose();
  document.body.replaceChildren();
});

async function instance(locale: 'en' | 'ja' = 'en'): Promise<SpreadsheetInstance> {
  const workbook = await WorkbookHandle.createDefault();
  open.push(workbook);
  const store = createSpreadsheetStore();
  workbook.attachStore(store);
  return {
    host: document.createElement('div'),
    store,
    history: new History(),
    workbook,
    i18n: { locale, strings: locale === 'ja' ? ja : en },
    isSheetProtected: () => false,
  } as unknown as SpreadsheetInstance;
}

// 1x1 transparent PNG.
const PNG = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=',
  ),
  (char) => char.charCodeAt(0),
);

const dashMenu = (): HTMLElement => {
  const menu = document.createElement('div');
  menu.innerHTML = `
    <button data-format-submenu="weight"></button>
    <button data-format-submenu="dash"></button>
    <button data-shape-outline-weight="1"></button>
    <button data-shape-outline-dash="dash"></button>
    <button data-shape-outline-dash="roundDot"></button>`;
  return menu;
};

const disabled = (menu: HTMLElement, selector: string): boolean =>
  menu.querySelector<HTMLButtonElement>(selector)?.disabled === true ||
  menu.querySelector<HTMLElement>(selector)?.getAttribute('aria-disabled') === 'true';

describe('shape controls against an engine that saves shapes', () => {
  it('keeps every dash style and the weight enabled for a saved shape', async () => {
    const sheet = await instance();
    createSessionShape(
      sheet.store,
      { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
      {
        id: 's1',
        shape: 'rectangle',
        x: 100,
        y: 60,
        w: 120,
        h: 80,
      },
    );
    mutators.setSelectedObject(sheet.store, { id: 's1', name: '' });
    const menu = dashMenu();

    createShapeFormatDropdownDefaults(sheet).updateShapeOutlineMenu(menu);

    expect(disabled(menu, '[data-shape-outline-dash="roundDot"]')).toBe(false);
    expect(disabled(menu, '[data-shape-outline-dash="dash"]')).toBe(false);
    expect(disabled(menu, '[data-format-submenu="dash"]')).toBe(false);
    expect(disabled(menu, '[data-shape-outline-weight]')).toBe(false);
    expect(disabled(menu, '[data-format-submenu="weight"]')).toBe(false);
  });

  it('keeps the arrow entry enabled in the standard and the Mac ribbon', async () => {
    const sheet = await instance();
    const host = document.createElement('div');
    host.innerHTML =
      '<button data-ribbon-command="mac.insert.shapeArrow"></button><button data-ribbon-command="mac.insert.shapeOval"></button>';
    projectMacRibbonState(host, sheet);
    expect(disabled(host, '[data-ribbon-command="mac.insert.shapeArrow"]')).toBe(false);
    expect(disabled(host, '[data-ribbon-command="mac.insert.shapeOval"]')).toBe(false);
  });

  it('inserts an arrow as a connector with a triangle at its end, and reloads it as one', async () => {
    const sheet = await instance();

    createIllustrationDropdownDefaults(sheet).insertShapeFromRibbon('arrow');

    const [shape] = sheet.workbook.listShapes(0);
    expect(shape).toMatchObject({
      presetGeometry: 'straightConnector1',
      line: { tailEnd: { type: ChartLineEndType.Triangle } },
    });
    const reopened = await WorkbookHandle.loadBytes(sheet.workbook.save());
    open.push(reopened);
    const loaded = createSpreadsheetStore();
    hydrateImagesFromEngine(reopened, loaded, 0);
    expect(loaded.getState().illustrations.illustrations[0]).toMatchObject({ shape: 'arrow' });
    expect(loaded.getState().illustrations.illustrations[0]?.nativeLocked).not.toBe(true);
  });

  it('saves a text box with the white fill and grey outline of a new one', async () => {
    const sheet = await instance();

    createIllustrationDropdownDefaults(sheet).insertTextBoxFromRibbon();

    const [shape] = sheet.workbook.listShapes(0);
    expect(shape).toMatchObject({ presetGeometry: 'rect', textBox: true });
    const xml = new TextDecoder().decode(
      zipEntries(sheet.workbook.save()).get('xl/drawings/drawing1.xml'),
    );
    expect(xml).toContain('txBox="1"');
    // The fill follows the theme's light colour (bg1 maps to lt1); the outline is the grey Excel renders.
    expect(xml).toMatch(/<\/a:prstGeom><a:solidFill><a:schemeClr val="bg1"\/><\/a:solidFill>/);
    expect(xml).toMatch(/<a:ln w="9525"[^>]*><a:solidFill><a:srgbClr val="BCBCBC"\/>/);
  });
});

describe('names of new drawings', () => {
  it.each([
    ['en', ['Rectangle 1', 'Rectangle 2', 'TextBox 1', 'Picture 1']],
    ['ja', ['正方形/長方形 1', '正方形/長方形 2', 'テキスト ボックス 1', '図 1']],
  ] as const)('saves the %s name of each new drawing', async (locale, names) => {
    const sheet = await instance(locale);
    const ribbon = createIllustrationDropdownDefaults(sheet);
    ribbon.insertShapeFromRibbon('rectangle');
    ribbon.insertShapeFromRibbon('rectangle');
    ribbon.insertTextBoxFromRibbon();
    createRibbonImageFromSelection(
      sheet.store,
      { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
      'data:image/png;base64,',
      sheet.history,
      'logo.png',
      { bytes: PNG, wb: sheet.workbook },
      undefined,
      sheet.i18n.strings.sessionIllustrations,
    );

    const reopened = await WorkbookHandle.loadBytes(sheet.workbook.save());
    open.push(reopened);
    expect(reopened.listDrawingObjects(0).map((object) => object.name)).toEqual(names);
    const picture = reopened.listDrawingObjects(0)[3];
    expect(picture?.description).toBe('logo.png');
    const loaded = createSpreadsheetStore();
    hydrateImagesFromEngine(reopened, loaded, 0);
    expect(loaded.getState().illustrations.illustrations.map((item) => item.name)).toEqual(names);
    expect(sheet.store.getState().illustrations.illustrations.map((item) => item.name)).toEqual(
      names,
    );
    const overlay = attachSessionIllustrations({
      host: sheet.host,
      store: sheet.store,
      labels: sheet.i18n.strings.sessionIllustrations,
    });
    const labels = Array.from(
      sheet.host.querySelectorAll<HTMLElement>('.fc-illustration'),
      (panel) => panel.getAttribute('aria-label'),
    );
    // A picture is announced by its alt text; the shapes by the names the file holds.
    expect(labels).toEqual([...names.slice(0, 3), 'logo.png']);
    overlay.detach();
  });

  it('numbers past the highest name of its kind, so a delete leaves no twin', async () => {
    const sheet = await instance();
    const ribbon = createIllustrationDropdownDefaults(sheet);
    ribbon.insertShapeFromRibbon('oval');
    ribbon.insertShapeFromRibbon('oval');
    const [first] = sheet.store.getState().illustrations.illustrations;
    if (first) mutators.removeIllustration(sheet.store, first.id);
    ribbon.insertShapeFromRibbon('oval');

    expect(sheet.workbook.listShapes(0).map((shape) => shape.name)).toEqual(['Oval 2', 'Oval 3']);
    // The overlay names them as saved, not by their position among the ovals.
    const overlay = attachSessionIllustrations({ host: sheet.host, store: sheet.store });
    const labels = Array.from(
      sheet.host.querySelectorAll<HTMLElement>('.fc-illustration'),
      (panel) => panel.getAttribute('aria-label'),
    );
    expect(labels).toEqual(['Oval 2', 'Oval 3']);
    overlay.detach();
  });
});
