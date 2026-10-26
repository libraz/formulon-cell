import { afterEach, describe, expect, it } from 'vitest';
import { History } from '../../../../src/commands/history.js';
import { createSessionShape } from '../../../../src/commands/session-illustration.js';
import { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { en } from '../../../../src/i18n/strings.js';
import type { SpreadsheetInstance } from '../../../../src/index.js';
import { createIllustrationDropdownDefaults } from '../../../../src/mount/dynamic-dropdown-defaults/illustrations.js';
import { createShapeFormatDropdownDefaults } from '../../../../src/mount/dynamic-dropdown-defaults/shape-format.js';
import { createSpreadsheetStore, mutators } from '../../../../src/store/store.js';
import { projectMacRibbonState } from '../../../../src/toolbar/ribbon/mac/model.js';

const open: WorkbookHandle[] = [];
afterEach(() => {
  for (const wb of open.splice(0)) wb.dispose();
  document.body.replaceChildren();
});

async function instance(): Promise<SpreadsheetInstance> {
  const workbook = await WorkbookHandle.createDefault();
  open.push(workbook);
  const store = createSpreadsheetStore();
  workbook.attachStore(store);
  return {
    host: document.createElement('div'),
    store,
    history: new History(),
    workbook,
    i18n: { locale: 'en', strings: en },
    isSheetProtected: () => false,
  } as unknown as SpreadsheetInstance;
}

const dashMenu = (): HTMLElement => {
  const menu = document.createElement('div');
  menu.innerHTML = `
    <button data-format-submenu="weight"></button>
    <button data-format-submenu="dash"></button>
    <button data-shape-outline-weight="1"></button>
    <button data-shape-outline-dash="dash"></button>`;
  return menu;
};

const disabled = (menu: HTMLElement, selector: string): boolean =>
  menu.querySelector<HTMLButtonElement>(selector)?.disabled === true ||
  menu.querySelector<HTMLElement>(selector)?.getAttribute('aria-disabled') === 'true';

describe('shape controls against an engine that saves shapes', () => {
  it('disables the dash style of a saved shape and keeps the weight', async () => {
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

    expect(disabled(menu, '[data-shape-outline-dash]')).toBe(true);
    expect(disabled(menu, '[data-format-submenu="dash"]')).toBe(true);
    expect(disabled(menu, '[data-shape-outline-weight]')).toBe(false);
    expect(disabled(menu, '[data-format-submenu="weight"]')).toBe(false);
  });

  it('disables the arrow entry in the standard and the Mac ribbon', async () => {
    const sheet = await instance();
    const menu = document.createElement('div');
    menu.innerHTML =
      '<button data-shape-insert="arrow"></button><button data-shape-insert="oval"></button>';
    createIllustrationDropdownDefaults(sheet).updateShapesInsertMenu?.(menu);
    expect(disabled(menu, '[data-shape-insert="arrow"]')).toBe(true);
    expect(disabled(menu, '[data-shape-insert="oval"]')).toBe(false);

    const host = document.createElement('div');
    host.innerHTML =
      '<button data-ribbon-command="mac.insert.shapeArrow"></button><button data-ribbon-command="mac.insert.shapeOval"></button>';
    projectMacRibbonState(host, sheet);
    expect(disabled(host, '[data-ribbon-command="mac.insert.shapeArrow"]')).toBe(true);
    expect(
      host.querySelector('[data-ribbon-command="mac.insert.shapeArrow"]')?.getAttribute('title'),
    ).toContain(en.ribbonMenu.shapeArrowNotSaved);
    expect(disabled(host, '[data-ribbon-command="mac.insert.shapeOval"]')).toBe(false);
  });

  it('explains instead of inserting an arrow, which cannot be saved', async () => {
    const sheet = await instance();

    createIllustrationDropdownDefaults(sheet).insertShapeFromRibbon('arrow');

    expect(sheet.workbook.listShapes(0)).toEqual([]);
    expect(sheet.store.getState().illustrations.illustrations).toEqual([]);
    await Promise.resolve();
    expect(document.body.textContent).toContain(
      en.workbookObjects.compatibilityDetails.shapes.slice(0, 40),
    );

    createIllustrationDropdownDefaults(sheet).insertShapeFromRibbon('oval');
    expect(sheet.workbook.listShapes(0)).toHaveLength(1);
  });
});
