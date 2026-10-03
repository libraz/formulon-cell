import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { Spreadsheet } from '../../../../src/mount.js';
import { type MountedStubSheet, mountStubSheet } from '../../../test-utils/mount.js';
import { stubHelpers } from './fixtures.js';

vi.setConfig({ testTimeout: 20_000 });

describe('Spreadsheet.mountToolbar', () => {
  let sheet: MountedStubSheet;
  let host: HTMLElement;

  beforeEach(async () => {
    sheet = await mountStubSheet({ locale: 'en' });
    host = document.createElement('div');
    document.body.appendChild(host);
  });

  afterEach(() => {
    sheet.dispose();
    host.remove();
  });

  it('opens the Mac More Functions hierarchy with pointer, hover, and two-stage Escape', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      platform: 'mac',
      activeTab: 'formulas',
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    const opener = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="mac.formulas.more"]',
    );
    const menu = host.querySelector<HTMLElement>('#menu-mac-formulas-more');
    const statistical = menu?.querySelector<HTMLElement>(
      '[data-function-category-submenu="statistical"]',
    );
    const engineering = menu?.querySelector<HTMLElement>(
      '[data-function-category-submenu="engineering"]',
    );
    const statisticalPanel = menu?.querySelector<HTMLElement>(
      '[data-function-category-panel="statistical"]',
    );
    const engineeringPanel = menu?.querySelector<HTMLElement>(
      '[data-function-category-panel="engineering"]',
    );
    if (!opener || !menu || !statistical || !engineering || !statisticalPanel || !engineeringPanel)
      throw new Error('Missing Mac More Functions hierarchy.');

    opener.click();
    expect(menu.hidden).toBe(false);

    const hover = new MouseEvent('mouseover', { bubbles: true });
    Object.defineProperty(hover, 'target', { value: engineering });
    const beforeHoverFocus = document.activeElement;
    expect(tb.dropdownsApi?.dynamicRibbonDropdownHover(hover)).toBe(true);
    expect(engineeringPanel.hidden).toBe(false);
    expect(statisticalPanel.hidden).toBe(true);
    expect(document.activeElement).toBe(beforeHoverFocus);

    const click = new MouseEvent('click', { bubbles: true, cancelable: true });
    Object.defineProperty(click, 'target', { value: statistical });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(click)).toBe(true);
    expect(statisticalPanel.hidden).toBe(false);
    expect(engineeringPanel.hidden).toBe(true);
    expect(statistical.getAttribute('aria-expanded')).toBe('true');

    statistical.focus();
    const open = new KeyboardEvent('keydown', {
      bubbles: true,
      cancelable: true,
      key: 'ArrowRight',
    });
    Object.defineProperty(open, 'target', { value: statistical });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownKeydown(open)).toBe(true);
    expect(statisticalPanel.querySelector<HTMLButtonElement>('button')).toBe(
      document.activeElement,
    );

    const closeChild = new KeyboardEvent('keydown', {
      bubbles: true,
      cancelable: true,
      key: 'ArrowLeft',
    });
    Object.defineProperty(closeChild, 'target', { value: document.activeElement });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownKeydown(closeChild)).toBe(true);
    expect(statisticalPanel.hidden).toBe(true);
    expect(document.activeElement).toBe(statistical);

    const closeRoot = new KeyboardEvent('keydown', {
      bubbles: true,
      cancelable: true,
      key: 'Escape',
    });
    Object.defineProperty(closeRoot, 'target', { value: statistical });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownKeydown(closeRoot)).toBe(true);
    expect(menu.hidden).toBe(true);
    expect(document.activeElement).toBe(opener);

    tb.dispose();
  });

  it('dismisses an all-disabled More child before its parent for pointer and keyboard opens', async () => {
    const workbook = await WorkbookHandle.createDefault();
    Object.defineProperty(workbook, 'functionNames', {
      configurable: true,
      value: () => ['CUBESET', 'CUBEVALUE'],
    });
    await sheet.instance.setWorkbook(workbook);
    sheet.instance.commands.setPolicy({
      editable: () => true,
      operations: { formulaEdit: true },
      defaultOperation: 'deny',
      restrict: ({ commandId }) => !commandId?.startsWith('mac.function.CUBE'),
    });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      platform: 'mac',
      activeTab: 'formulas',
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    const opener = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="mac.formulas.more"]',
    );
    const menu = host.querySelector<HTMLElement>('#menu-mac-formulas-more');
    const cubeTrigger = menu?.querySelector<HTMLElement>('[data-function-category-submenu="cube"]');
    const cubePanel = menu?.querySelector<HTMLElement>('[data-function-category-panel="cube"]');
    if (!opener || !menu || !cubeTrigger || !cubePanel)
      throw new Error('Missing disabled More Functions hierarchy.');

    const cubeLeaves = (): HTMLButtonElement[] =>
      Array.from(
        cubePanel.querySelectorAll<HTMLButtonElement>('[data-ribbon-command^="mac.function."]'),
      );
    expect(cubeLeaves().length).toBeGreaterThan(0);
    expect(cubeLeaves().every((button) => button.disabled)).toBe(true);

    opener.click();
    const pointerOpen = new MouseEvent('click', { bubbles: true, cancelable: true });
    Object.defineProperty(pointerOpen, 'target', { value: cubeTrigger });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(pointerOpen)).toBe(true);
    expect(cubePanel.hidden).toBe(false);

    // Pointer opening can leave focus on the parent trigger. Escape must
    // therefore dismiss the visible child before the More root.
    const pointerEscape = new KeyboardEvent('keydown', {
      bubbles: true,
      cancelable: true,
      key: 'Escape',
    });
    Object.defineProperty(pointerEscape, 'target', { value: cubeTrigger });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownKeydown(pointerEscape)).toBe(true);
    expect(cubePanel.hidden).toBe(true);
    expect(menu.hidden).toBe(false);
    expect(document.activeElement).toBe(cubeTrigger);

    const rootEscape = new KeyboardEvent('keydown', {
      bubbles: true,
      cancelable: true,
      key: 'Escape',
    });
    Object.defineProperty(rootEscape, 'target', { value: cubeTrigger });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownKeydown(rootEscape)).toBe(true);
    expect(menu.hidden).toBe(true);
    expect(document.activeElement).toBe(opener);

    opener.click();
    cubeTrigger.focus();
    const keyboardOpen = new KeyboardEvent('keydown', {
      bubbles: true,
      cancelable: true,
      key: 'ArrowRight',
    });
    Object.defineProperty(keyboardOpen, 'target', { value: cubeTrigger });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownKeydown(keyboardOpen)).toBe(true);
    expect(cubePanel.hidden).toBe(false);

    const keyboardClose = new KeyboardEvent('keydown', {
      bubbles: true,
      cancelable: true,
      key: 'ArrowLeft',
    });
    Object.defineProperty(keyboardClose, 'target', { value: cubePanel });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownKeydown(keyboardClose)).toBe(true);
    expect(cubePanel.hidden).toBe(true);
    expect(menu.hidden).toBe(false);
    expect(document.activeElement).toBe(cubeTrigger);

    const keyboardRootEscape = new KeyboardEvent('keydown', {
      bubbles: true,
      cancelable: true,
      key: 'Escape',
    });
    Object.defineProperty(keyboardRootEscape, 'target', { value: cubeTrigger });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownKeydown(keyboardRootEscape)).toBe(true);
    expect(menu.hidden).toBe(true);
    expect(document.activeElement).toBe(opener);

    tb.dispose();
  });

  it('scrolls and horizontally clamps a tall Mac More submenu in a narrow viewport', () => {
    const previousWidth = window.innerWidth;
    const previousHeight = window.innerHeight;
    Object.defineProperty(window, 'innerWidth', { configurable: true, value: 390 });
    Object.defineProperty(window, 'innerHeight', { configurable: true, value: 360 });
    try {
      const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
        platform: 'mac',
        activeTab: 'formulas',
        dynamicDropdowns: true,
        helpers: stubHelpers(),
      });
      const opener = host.querySelector<HTMLButtonElement>(
        '[data-ribbon-command="mac.formulas.more"]',
      );
      const menu = host.querySelector<HTMLElement>('#menu-mac-formulas-more');
      const trigger = menu?.querySelector<HTMLElement>(
        '[data-function-category-submenu="statistical"]',
      );
      const panel = menu?.querySelector<HTMLElement>(
        '[data-function-category-panel="statistical"]',
      );
      if (!opener || !menu || !trigger || !panel) throw new Error('Missing More submenu geometry.');

      opener.click();
      panel.style.maxHeight = '300px';
      panel.style.overflowY = 'visible';
      panel.style.overscrollBehavior = 'none';
      menu.getBoundingClientRect = vi.fn(
        () => ({ left: 8, right: 224, top: 20, bottom: 194, width: 216, height: 174 }) as DOMRect,
      );
      trigger.getBoundingClientRect = vi.fn(
        () => ({ left: 8, right: 224, top: 20, bottom: 48, width: 216, height: 28 }) as DOMRect,
      );
      const measurePanel = vi.fn(() => {
        expect(panel.style.maxHeight).toBe('');
        expect(panel.style.overflowY).toBe('');
        expect(panel.style.overscrollBehavior).toBe('');
        return { left: 0, right: 0, top: 0, bottom: 0, width: 0, height: 0 } as DOMRect;
      });
      panel.getBoundingClientRect = measurePanel;
      Object.defineProperty(panel, 'offsetWidth', { configurable: true, value: 216 });
      Object.defineProperty(panel, 'offsetHeight', { configurable: true, value: 0 });
      Object.defineProperty(panel, 'scrollHeight', { configurable: true, value: 3083 });

      const event = new MouseEvent('click', { bubbles: true, cancelable: true });
      Object.defineProperty(event, 'target', { value: trigger });
      expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);
      expect(panel.hidden).toBe(false);
      expect(panel.style.overflowY).toBe('auto');
      expect(panel.style.overscrollBehavior).toBe('contain');
      expect(panel.style.maxHeight).toBe('155px');
      expect(panel.style.left).toBe('0px');
      expect(panel.style.top).toBe('177px');
      expect(8 + Number.parseInt(panel.style.left, 10)).toBeGreaterThanOrEqual(8);
      expect(8 + Number.parseInt(panel.style.left, 10) + 216).toBeLessThanOrEqual(382);
      expect(20 + Number.parseInt(panel.style.top, 10)).toBe(197);
      expect(20 + Number.parseInt(panel.style.top, 10)).toBeGreaterThanOrEqual(194 + 3);
      expect(20 + Number.parseInt(panel.style.top, 10) + 155).toBeLessThanOrEqual(352);
      expect(measurePanel).toHaveBeenCalledTimes(1);

      tb.dispose();
    } finally {
      Object.defineProperty(window, 'innerWidth', { configurable: true, value: previousWidth });
      Object.defineProperty(window, 'innerHeight', { configurable: true, value: previousHeight });
    }
  });

  it.each([
    {
      name: 'above when above space wins',
      viewportHeight: 260,
      menu: { top: 100, bottom: 220, height: 120 },
      expectedTop: -92,
      expectedHeight: 89,
      expectedPageTop: 8,
      expectedPageBottom: 97,
    },
    {
      name: 'below with a sub-row viewport cap',
      viewportHeight: 240,
      menu: { top: 20, bottom: 194, height: 174 },
      expectedTop: 177,
      expectedHeight: 35,
      expectedPageTop: 197,
      expectedPageBottom: 232,
    },
  ])(
    'stacks a narrow More submenu $name without overlapping its parent',
    ({
      viewportHeight,
      menu: menuGeometry,
      expectedTop,
      expectedHeight,
      expectedPageTop,
      expectedPageBottom,
    }) => {
      const previousWidth = window.innerWidth;
      const previousHeight = window.innerHeight;
      Object.defineProperty(window, 'innerWidth', { configurable: true, value: 390 });
      Object.defineProperty(window, 'innerHeight', { configurable: true, value: viewportHeight });
      try {
        const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
          platform: 'mac',
          activeTab: 'formulas',
          dynamicDropdowns: true,
          helpers: stubHelpers(),
        });
        const opener = host.querySelector<HTMLButtonElement>(
          '[data-ribbon-command="mac.formulas.more"]',
        );
        const menu = host.querySelector<HTMLElement>('#menu-mac-formulas-more');
        const trigger = menu?.querySelector<HTMLElement>(
          '[data-function-category-submenu="statistical"]',
        );
        const panel = menu?.querySelector<HTMLElement>(
          '[data-function-category-panel="statistical"]',
        );
        if (!opener || !menu || !trigger || !panel)
          throw new Error('Missing More submenu stacking geometry.');

        opener.click();
        menu.getBoundingClientRect = vi.fn(
          () =>
            ({
              left: 8,
              right: 224,
              top: menuGeometry.top,
              bottom: menuGeometry.bottom,
              width: 216,
              height: menuGeometry.height,
            }) as DOMRect,
        );
        trigger.getBoundingClientRect = vi.fn(
          () =>
            ({
              left: 8,
              right: 224,
              top: menuGeometry.top,
              bottom: menuGeometry.top + 28,
              width: 216,
              height: 28,
            }) as DOMRect,
        );
        panel.getBoundingClientRect = vi.fn(
          () => ({ left: 0, right: 0, top: 0, bottom: 0, width: 0, height: 0 }) as DOMRect,
        );
        Object.defineProperty(panel, 'offsetWidth', { configurable: true, value: 216 });
        Object.defineProperty(panel, 'offsetHeight', { configurable: true, value: 0 });
        Object.defineProperty(panel, 'scrollHeight', { configurable: true, value: 3083 });

        const event = new MouseEvent('click', { bubbles: true, cancelable: true });
        Object.defineProperty(event, 'target', { value: trigger });
        expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);
        expect(panel.style.left).toBe('0px');
        expect(panel.style.top).toBe(`${expectedTop}px`);
        expect(panel.style.maxHeight).toBe(`${expectedHeight}px`);
        expect(panel.style.overflowY).toBe('auto');
        expect(panel.style.overscrollBehavior).toBe('contain');
        expect(8 + Number.parseInt(panel.style.left, 10)).toBeGreaterThanOrEqual(8);
        expect(8 + Number.parseInt(panel.style.left, 10) + 216).toBeLessThanOrEqual(382);
        expect(menuGeometry.top + Number.parseInt(panel.style.top, 10)).toBe(expectedPageTop);
        expect(expectedPageBottom).toBeLessThanOrEqual(viewportHeight - 8);
        if (expectedPageTop < menuGeometry.top) {
          expect(expectedPageBottom).toBeLessThanOrEqual(menuGeometry.top - 3);
        } else {
          expect(expectedPageTop).toBeGreaterThanOrEqual(menuGeometry.bottom + 3);
        }

        tb.dispose();
      } finally {
        Object.defineProperty(window, 'innerWidth', { configurable: true, value: previousWidth });
        Object.defineProperty(window, 'innerHeight', { configurable: true, value: previousHeight });
      }
    },
  );

  it('keeps More navigation available under policy while denying a restricted function leaf', () => {
    sheet.instance.commands.setPolicy({
      editable: () => true,
      operations: { formulaEdit: true },
      defaultOperation: 'deny',
      restrict: ({ commandId }) => commandId !== 'mac.function.COUNTIF',
    });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      platform: 'mac',
      activeTab: 'formulas',
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    const opener = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="mac.formulas.more"]',
    );
    const menu = host.querySelector<HTMLElement>('#menu-mac-formulas-more');
    const statistical = menu?.querySelector<HTMLButtonElement>(
      '[data-function-category-submenu="statistical"]',
    );
    if (!opener || !menu || !statistical) throw new Error('Missing More Functions controls.');

    opener.click();
    expect(menu.hidden).toBe(false);
    statistical.click();
    const countIf = menu.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="mac.function.COUNTIF"]',
    );
    expect(countIf?.getAttribute('aria-disabled')).toBe('true');

    countIf?.click();
    expect(host.querySelector('.fc-fxdialog')).toBeNull();
    expect(menu.hidden).toBe(false);

    tb.dispose();
  });

  it('projects the initial Mac function catalog once across all category menus', async () => {
    const workbook = await WorkbookHandle.createDefault();
    expect(workbook.isStub).toBe(false);
    // Keep the workbook native for mount wiring, but use a tiny deterministic
    // catalog fixture; this counts projector method calls, not raw WASM work.
    const functionNames = vi.fn(() => ['COUNTIF', 'SUM'] as const);
    Object.defineProperty(workbook, 'functionNames', {
      configurable: true,
      value: functionNames,
    });
    await sheet.instance.setWorkbook(workbook);

    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      platform: 'mac',
      activeTab: 'formulas',
      helpers: stubHelpers(),
    });
    try {
      expect(functionNames).toHaveBeenCalledTimes(1);

      const categoryLeaves = (menuId: string): string[] =>
        Array.from(
          host.querySelectorAll<HTMLButtonElement>(`#${menuId} [data-ribbon-command]`),
        ).map((button) => button.dataset.ribbonCommand ?? '');
      expect(categoryLeaves('menu-mac-formulas-financial')).toHaveLength(0);
      expect(categoryLeaves('menu-mac-formulas-logical')).toHaveLength(0);
      expect(categoryLeaves('menu-mac-formulas-text')).toHaveLength(0);
      expect(categoryLeaves('menu-mac-formulas-date-time')).toHaveLength(0);
      expect(categoryLeaves('menu-mac-formulas-lookup')).toHaveLength(0);
      expect(categoryLeaves('menu-mac-formulas-math')).toEqual(['mac.function.SUM']);
      expect(categoryLeaves('menu-mac-formulas-more-statistical')).toEqual([
        'mac.function.COUNTIF',
      ]);
      expect(categoryLeaves('menu-mac-formulas-more-engineering')).toHaveLength(0);
      expect(categoryLeaves('menu-mac-formulas-more-cube')).toHaveLength(0);
      expect(categoryLeaves('menu-mac-formulas-more-information')).toHaveLength(0);
      expect(categoryLeaves('menu-mac-formulas-more-compatibility')).toHaveLength(0);
      expect(categoryLeaves('menu-mac-formulas-more-web')).toHaveLength(0);
    } finally {
      tb.dispose();
    }
  });

  it('reprojects More Functions after an actual mounted workbook swap', async () => {
    const first = await WorkbookHandle.createDefault();
    const second = await WorkbookHandle.createDefault();
    const empty = await WorkbookHandle.createDefault();
    const staticFallback = await WorkbookHandle.createDefault();
    expect(first.isStub).toBe(false);
    expect(second.isStub).toBe(false);
    expect(empty.isStub).toBe(false);
    expect(staticFallback.isStub).toBe(false);

    // Default native workbooks expose the same catalog. Keep the handles real
    // and vary only their catalog response so the mounted setWorkbook path is
    // exercised without replacing the projector directly.
    const setCatalog = (workbook: WorkbookHandle, names: readonly string[] | null): void => {
      Object.defineProperty(workbook, 'functionNames', {
        configurable: true,
        value: () => names,
      });
    };
    setCatalog(first, ['COUNTIF']);
    setCatalog(second, ['AVERAGE', 'CUBEVALUE']);
    setCatalog(empty, []);
    setCatalog(staticFallback, null);

    await sheet.instance.setWorkbook(first);
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      platform: 'mac',
      activeTab: 'formulas',
      helpers: stubHelpers(),
    });
    try {
      const panelLeaves = (category: string): string[] =>
        Array.from(
          host.querySelectorAll<HTMLButtonElement>(
            `#menu-mac-formulas-more [data-function-category-panel="${category}"] [data-ribbon-command]`,
          ),
        ).map((button) => button.dataset.ribbonCommand ?? '');

      expect(panelLeaves('statistical')).toContain('mac.function.COUNTIF');
      expect(panelLeaves('statistical')).not.toContain('mac.function.AVERAGE');

      await sheet.instance.setWorkbook(second);
      expect(panelLeaves('statistical')).toContain('mac.function.AVERAGE');
      expect(panelLeaves('statistical')).not.toContain('mac.function.COUNTIF');
      expect(panelLeaves('cube')).toContain('mac.function.CUBEVALUE');

      await sheet.instance.setWorkbook(empty);
      expect(
        host.querySelectorAll('#menu-mac-formulas-more [data-function-category-panel]'),
      ).toHaveLength(6);
      expect(
        host.querySelectorAll(
          '#menu-mac-formulas-more [data-function-category-panel] [data-ribbon-command]',
        ),
      ).toHaveLength(0);

      await sheet.instance.setWorkbook(staticFallback);
      expect(panelLeaves('statistical')).toContain('mac.function.COUNTIF');
      expect(panelLeaves('cube')).toHaveLength(0);
    } finally {
      tb.dispose();
    }
  });
});
