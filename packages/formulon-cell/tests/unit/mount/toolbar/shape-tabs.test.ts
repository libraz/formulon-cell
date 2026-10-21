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
import { selectStoreChart, selectStorePicture, stubHelpers } from './fixtures.js';

vi.setConfig({ testTimeout: 20_000 });

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
