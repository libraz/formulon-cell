import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import {
  shapeFillColor,
  shapeOutlineColor,
  shapeOutlineDash,
  shapeOutlineWeight,
} from '../../../../src/commands/shape-format.js';
import { createIllustrationDropdownDefaults } from '../../../../src/mount/dynamic-dropdown-defaults/illustrations.js';
import { Spreadsheet } from '../../../../src/mount.js';
import { mutators, type SessionIllustration } from '../../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../../test-utils/mount.js';
import { scaledTimeout } from '../../../timeouts.js';
import { selectStoreChart, selectStorePicture, stubHelpers } from './fixtures.js';

vi.setConfig({ testTimeout: scaledTimeout(20_000) });

describe('shape and picture tabs of the ribbon', () => {
  let sheet: MountedStubSheet;
  let host: HTMLElement;

  beforeEach(async () => {
    sheet = await mountStubSheet({ locale: 'en' });
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 });
    host = document.createElement('div');
    document.body.appendChild(host);
  });

  afterEach(() => {
    sheet.dispose();
    host.remove();
  });

  const tabIds = (): string[] =>
    Array.from(host.querySelectorAll<HTMLElement>('[data-ribbon-tab]')).map(
      (tab) => tab.dataset.ribbonTab ?? '',
    );
  const command = (id: string): HTMLButtonElement => {
    const found = host.querySelector<HTMLButtonElement>(`[data-ribbon-command="${id}"]`);
    if (!found) throw new Error(`missing ${id}`);
    return found;
  };
  const illustrations = (): readonly SessionIllustration[] =>
    sheet.instance.store.getState().illustrations.illustrations;
  const insertShape = (): SessionIllustration => {
    createIllustrationDropdownDefaults(sheet.instance).insertShapeFromRibbon('rectangle');
    const item = illustrations().at(-1);
    if (!item) throw new Error('no shape');
    return item;
  };
  const current = (id: string): SessionIllustration => {
    const found = illustrations().find((item) => item.id === id);
    if (!found) throw new Error('shape is gone');
    return found;
  };
  /** Click the menu entry `selector` the way the ribbon's click delegation sees it. */
  const pick = (tb: ReturnType<typeof Spreadsheet.mountToolbar>, selector: string): void => {
    const target = host.querySelector<HTMLElement>(selector);
    expect(target, selector).toBeTruthy();
    const click = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(click, 'target', { value: target });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(click)).toBe(true);
  };

  it('opens Shape Format on an inserted shape and goes back once it is left', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('insert');
    insertShape();
    expect(tabIds().at(-1)).toBe('shapeFormat');
    expect(tabIds()).not.toContain('chartDesign');
    expect(tb.getActiveTab()).toBe('shapeFormat');

    sheet.host.dispatchEvent(new PointerEvent('pointerdown', { bubbles: true }));
    expect(tabIds()).not.toContain('shapeFormat');
    expect(tb.getActiveTab()).toBe('insert');
    tb.dispose();
  });

  it('fills and outlines the selected shape from its menus', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    const shape = insertShape();

    command('shapeFill').click();
    const red = host.querySelector<HTMLButtonElement>('[data-shape-fill="#ff0000"]');
    expect(red?.getAttribute('aria-pressed')).toBe('false');
    pick(tb, '[data-shape-fill="#ff0000"]');
    expect(shapeFillColor(current(shape.id))).toBe('#ff0000');

    command('shapeFill').click();
    expect(host.querySelector('[data-shape-fill="#ff0000"]')?.getAttribute('aria-pressed')).toBe(
      'true',
    );
    pick(tb, '[data-shape-fill="none"]');
    expect(shapeFillColor(current(shape.id))).toBeNull();

    command('shapeOutline').click();
    pick(tb, '[data-shape-outline="none"]');
    expect(shapeOutlineColor(current(shape.id))).toBeNull();
    command('shapeOutline').click();
    expect(host.querySelector('[data-shape-outline="none"]')?.getAttribute('aria-checked')).toBe(
      'true',
    );

    // The shape stays selected and each change is its own undo step.
    expect(sheet.instance.store.getState().ui.selectedObject?.id).toBe(shape.id);
    sheet.instance.history.undo();
    expect(shapeOutlineColor(current(shape.id))).not.toBeNull();
    expect(shapeFillColor(current(shape.id))).toBeNull();
    tb.dispose();
  });

  it('weighs and dashes the outline from the submenus of Shape Outline', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    const shape = insertShape();

    command('shapeOutline').click();
    const panel = host.querySelector<HTMLElement>('[data-format-panel="weight"]');
    expect(panel?.hidden).toBe(true);
    pick(tb, '[data-format-submenu="weight"]');
    expect(panel?.hidden).toBe(false);
    expect(
      host.querySelector('[data-shape-outline-weight="1.5"]')?.getAttribute('aria-checked'),
    ).toBe('true');
    pick(tb, '[data-shape-outline-weight="3"]');
    expect(shapeOutlineWeight(current(shape.id))).toBe(3);

    command('shapeOutline').click();
    pick(tb, '[data-shape-outline-dash="dash"]');
    expect(shapeOutlineDash(current(shape.id))).toBe('dash');
    command('shapeOutline').click();
    expect(
      host.querySelector('[data-shape-outline-weight="3"]')?.getAttribute('aria-checked'),
    ).toBe('true');
    expect(
      host.querySelector('[data-shape-outline-dash="dash"]')?.getAttribute('aria-checked'),
    ).toBe('true');
    tb.dispose();
  });

  it('turns the selected shape a quarter turn from the Rotate menu', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    const shape = insertShape();

    command('shapeRotate').click();
    pick(tb, '[data-shape-rotate="right"]');
    expect(current(shape.id).rotationDeg).toBe(90);
    command('shapeRotate').click();
    pick(tb, '[data-shape-rotate="left"]');
    command('shapeRotate').click();
    pick(tb, '[data-shape-rotate="left"]');
    expect(current(shape.id).rotationDeg).toBe(270);

    expect(sheet.instance.store.getState().ui.selectedObject?.id).toBe(shape.id);
    sheet.instance.history.undo();
    expect(current(shape.id).rotationDeg).toBe(0);
    tb.dispose();
  });

  it('flips the selected shape from the Rotate menu, mirroring its rotation', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    const shape = insertShape();

    command('shapeRotate').click();
    const entries = Array.from(
      host.querySelectorAll<HTMLElement>('[data-shape-rotate]'),
      (entry) => entry.dataset.shapeRotate,
    );
    expect(entries).toEqual(['right', 'left', 'flip-vertical', 'flip-horizontal']);
    pick(tb, '[data-shape-rotate="right"]');
    command('shapeRotate').click();
    pick(tb, '[data-shape-rotate="flip-horizontal"]');
    expect(current(shape.id)).toMatchObject({ flipH: true, rotationDeg: 270 });
    command('shapeRotate').click();
    pick(tb, '[data-shape-rotate="flip-vertical"]');
    expect(current(shape.id)).toMatchObject({ flipH: true, flipV: true, rotationDeg: 90 });

    sheet.instance.history.undo();
    expect(current(shape.id)).toMatchObject({ flipH: true, rotationDeg: 270 });
    expect(current(shape.id).flipV).toBeUndefined();
    sheet.instance.history.undo();
    expect(current(shape.id)).toMatchObject({ rotationDeg: 90 });
    expect(current(shape.id).flipH).toBeUndefined();
    tb.dispose();
  });

  it('changes the selected shape to another from the shapes gallery', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    const shape = insertShape();

    command('shapeChangeShape').click();
    const kinds = Array.from(
      host.querySelectorAll<HTMLElement>('[data-shape-change]'),
      (tile) => tile.dataset.shapeChange,
    );
    expect(kinds).toEqual(['rectangle', 'rounded-rectangle', 'oval', 'triangle', 'diamond']);
    expect(host.querySelector('[data-shape-change="oval"]')?.getAttribute('aria-label')).toBe(
      host.querySelector('[data-shape-insert="oval"]')?.getAttribute('aria-label') ?? 'Oval',
    );
    pick(tb, '[data-shape-change="oval"]');
    expect(current(shape.id)).toEqual({ ...shape, shape: 'oval' });

    expect(sheet.instance.store.getState().ui.selectedObject?.id).toBe(shape.id);
    sheet.instance.history.undo();
    expect(current(shape.id).shape).toBe('rectangle');
    tb.dispose();
  });

  it('greys Change Shape for a line, which still turns', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    createIllustrationDropdownDefaults(sheet.instance).insertShapeFromRibbon('line');

    command('shapeChangeShape').click();
    const tiles = host.querySelectorAll<HTMLButtonElement>('[data-shape-change]');
    expect(tiles.length).toBeGreaterThan(0);
    for (const tile of tiles) expect(tile.disabled, tile.dataset.shapeChange).toBe(true);
    command('shapeRotate').click();
    expect(host.querySelector<HTMLButtonElement>('[data-shape-rotate="right"]')?.disabled).toBe(
      false,
    );
    tb.dispose();
  });

  it('brings the selected shape forward and sends it backward', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    const back = insertShape();
    const front = insertShape();
    expect(illustrations().map((item) => item.id)).toEqual([back.id, front.id]);

    command('objectSendBackward').click();
    expect(illustrations().map((item) => item.id)).toEqual([front.id, back.id]);
    command('objectBringForward').click();
    expect(illustrations().map((item) => item.id)).toEqual([back.id, front.id]);
    expect(sheet.instance.store.getState().ui.selectedObject?.id).toBe(front.id);
    tb.dispose();
  });

  it('turns and flips the selected picture from the Picture Format Rotate button', () => {
    selectStorePicture(sheet.instance.store);
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('pictureFormat');

    command('shapeRotate').click();
    pick(tb, '[data-shape-rotate="right"]');
    expect(current('picture-test').rotationDeg).toBe(90);
    command('shapeRotate').click();
    pick(tb, '[data-shape-rotate="flip-horizontal"]');
    expect(current('picture-test')).toMatchObject({ flipH: true, rotationDeg: 270 });

    sheet.instance.history.undo();
    expect(current('picture-test').flipH).toBeUndefined();
    expect(current('picture-test').rotationDeg).toBe(90);
    tb.dispose();
  });

  it('follows the selection from one kind of object to another', () => {
    selectStorePicture(sheet.instance.store);
    selectStoreChart(sheet.instance.store);
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    expect(tabIds().slice(-2)).toEqual(['chartDesign', 'chartFormat']);
    tb.setActiveTab('chartFormat');

    mutators.setSelectedObject(sheet.instance.store, { id: 'picture-test', name: 'Picture 1' });
    expect(tabIds().at(-1)).toBe('pictureFormat');
    expect(tabIds()).not.toContain('chartFormat');
    // The chart's tab went away with it, and a picture already there opens nothing.
    expect(tb.getActiveTab()).toBe('home');
    expect(host.querySelector('.fc-chart--selected')).toBeNull();
    tb.dispose();
  });
});
