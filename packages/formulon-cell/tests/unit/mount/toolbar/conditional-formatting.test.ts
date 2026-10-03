import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { Spreadsheet } from '../../../../src/mount.js';
import { mutators } from '../../../../src/store/store.js';
import { SUPPORTED_CONDITIONAL_MENU_ACTIONS } from '../../../../src/toolbar/ribbon/conditional-menu-action.js';
import { type MountedStubSheet, mountStubSheet } from '../../../test-utils/mount.js';
import { dynamicDropdownNoopOverrides, stubHelpers } from './fixtures.js';

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

  it('adds prompted conditional formatting rules through the Home dropdown', async () => {
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 0 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    const conditionalButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="conditional"]',
    );
    expect(conditionalButton).toBeTruthy();
    conditionalButton?.click();
    const greaterThanButton = host.querySelector<HTMLButtonElement>('[data-cf-action="cell-gt"]');
    expect(greaterThanButton).toBeTruthy();
    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: greaterThanButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);

    await Promise.resolve();
    const dialog = document.body.querySelector<HTMLElement>('.fc-tb__dlg');
    expect(dialog?.textContent).toContain('Format cells that are GREATER THAN:');
    expect(dialog?.textContent).toContain('Light Red Fill with Dark Red Text');
    expect(dialog?.textContent).toContain('Yellow Fill with Dark Yellow Text');
    const formatSelect = dialog?.querySelector<HTMLSelectElement>('select.fc-tb__dlg__select');
    expect(formatSelect?.value).toBe('light-red-dark-red');
    const formatPreview = dialog?.querySelector<HTMLElement>('[data-conditional-format-preview]');
    expect(formatPreview?.style.background).toBe('#ffc7ce');
    if (!formatSelect) throw new Error('Expected conditional formatting style select.');
    formatSelect.value = 'yellow-dark-yellow';
    formatSelect.dispatchEvent(new Event('change', { bubbles: true }));
    expect(formatPreview?.style.background).toBe('#ffeb9c');
    const input = dialog?.querySelector<HTMLInputElement>('input[type="number"]');
    expect(input).toBeTruthy();
    if (!input) throw new Error('Expected conditional formatting number dialog.');
    input.value = '5';
    dialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await Promise.resolve();
    await Promise.resolve();

    expect(sheet.instance.store.getState().conditional.rules).toHaveLength(1);
    expect(sheet.instance.store.getState().conditional.rules[0]).toMatchObject({
      kind: 'cell-value',
      range: { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 0 },
      op: '>',
      a: 5,
      apply: { fill: '#ffeb9c', color: '#9c6500' },
    });

    tb.dispose();
  });

  it('opens Conditional Formatting submenus on hover through shared dropdown wiring', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    const conditionalButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="conditional"]',
    );
    expect(conditionalButton).toBeTruthy();
    conditionalButton?.click();

    const menu = host.querySelector<HTMLElement>('#menu-conditional');
    const highlight = menu?.querySelector<HTMLElement>('[data-cf-submenu="highlight"]');
    const topBottom = menu?.querySelector<HTMLElement>('[data-cf-submenu="topBottom"]');
    expect(menu?.hidden).toBe(false);
    expect(highlight?.textContent).toContain('Highlight Cells Rules');
    expect(topBottom?.textContent).toContain('Top/Bottom Rules');

    const hoverHighlight = new MouseEvent('mouseover', { bubbles: true });
    Object.defineProperty(hoverHighlight, 'target', { value: highlight });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownHover(hoverHighlight)).toBe(true);
    expect(menu?.querySelector<HTMLElement>('[data-cf-panel="highlight"]')?.hidden).toBe(false);
    expect(highlight?.getAttribute('aria-expanded')).toBe('true');

    const hoverTopBottom = new MouseEvent('mouseover', { bubbles: true });
    Object.defineProperty(hoverTopBottom, 'target', { value: topBottom });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownHover(hoverTopBottom)).toBe(true);
    expect(menu?.querySelector<HTMLElement>('[data-cf-panel="highlight"]')?.hidden).toBe(true);
    expect(highlight?.getAttribute('aria-expanded')).toBe('false');
    expect(menu?.querySelector<HTMLElement>('[data-cf-panel="topBottom"]')?.hidden).toBe(false);
    expect(topBottom?.classList.contains('fc-tb__menu-item--active')).toBe(true);
    expect(topBottom?.getAttribute('aria-expanded')).toBe('true');

    tb.dispose();
  });

  it('opens and closes Conditional Formatting submenus from keyboard events', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    const conditionalButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="conditional"]',
    );
    conditionalButton?.click();

    const menu = host.querySelector<HTMLElement>('#menu-conditional');
    const colorScale = menu?.querySelector<HTMLElement>('[data-cf-submenu="colorScale"]');
    const keyOpen = new KeyboardEvent('keydown', {
      bubbles: true,
      cancelable: true,
      key: 'ArrowRight',
    });
    Object.defineProperty(keyOpen, 'target', { value: colorScale });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownKeydown(keyOpen)).toBe(true);
    const panel = menu?.querySelector<HTMLElement>('[data-cf-panel="colorScale"]');
    expect(panel?.hidden).toBe(false);
    expect(colorScale?.getAttribute('aria-expanded')).toBe('true');
    expect(panel?.querySelector<HTMLButtonElement>('button')?.tabIndex).toBe(0);

    const firstPanelButton = panel?.querySelector<HTMLButtonElement>('button');
    const keyClose = new KeyboardEvent('keydown', {
      bubbles: true,
      cancelable: true,
      key: 'ArrowLeft',
    });
    Object.defineProperty(keyClose, 'target', { value: firstPanelButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownKeydown(keyClose)).toBe(true);
    expect(panel?.hidden).toBe(true);
    expect(colorScale?.getAttribute('aria-expanded')).toBe('false');
    expect(document.activeElement).toBe(colorScale);

    tb.dispose();
  });

  it('keeps the Conditional Formatting top-level menu order aligned with Excel 365', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    host.querySelector<HTMLButtonElement>('[data-ribbon-command="conditional"]')?.click();
    const menu = host.querySelector<HTMLElement>('#menu-conditional');
    const topLevelButtons = Array.from(menu?.children ?? []).filter(
      (child): child is HTMLButtonElement => child instanceof HTMLButtonElement,
    );

    expect(
      topLevelButtons.map((button) => button.dataset.cfSubmenu ?? button.dataset.cfAction),
    ).toEqual([
      'highlight',
      'topBottom',
      'dataBar',
      'colorScale',
      'iconSet',
      'new-rule',
      'clear',
      'manage',
    ]);
    expect(topLevelButtons.map((button) => button.textContent?.replace('▶', '').trim())).toEqual([
      'Highlight Cells Rules',
      'Top/Bottom Rules',
      'Data Bars',
      'Color Scales',
      'Icon Sets',
      'New Rule...',
      'Clear Rules',
      'Manage Rules...',
    ]);

    tb.dispose();
  });

  it('keeps the Conditional Formatting top-level menu icon and submenu affordances', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    host.querySelector<HTMLButtonElement>('[data-ribbon-command="conditional"]')?.click();
    const menu = host.querySelector<HTMLElement>('#menu-conditional');
    const topLevelButtons = Array.from(menu?.children ?? []).filter(
      (child): child is HTMLButtonElement => child instanceof HTMLButtonElement,
    );

    expect(topLevelButtons).toHaveLength(8);
    for (const button of topLevelButtons) {
      expect(button.classList.contains('fc-tb__menu-item--preset')).toBe(true);
      expect(button.querySelector('.fc-tb__cf-icon')).toBeTruthy();
      expect(button.querySelector('.fc-tb__menu-item__text')?.textContent?.trim()).not.toBe('');
    }
    const submenuButtons = topLevelButtons.filter((button) => button.dataset.cfSubmenu);
    expect(submenuButtons.map((button) => button.dataset.cfSubmenu)).toEqual([
      'highlight',
      'topBottom',
      'dataBar',
      'colorScale',
      'iconSet',
      'clear',
    ]);
    expect(
      submenuButtons.every((button) => !!button.querySelector('.fc-tb__menu-item__caret')),
    ).toBe(true);
    expect(submenuButtons.every((button) => button.getAttribute('aria-haspopup') === 'menu')).toBe(
      true,
    );
    expect(submenuButtons.every((button) => button.getAttribute('aria-expanded') === 'false')).toBe(
      true,
    );
    for (const button of submenuButtons) {
      const panelId = button.getAttribute('aria-controls');
      expect(panelId).toBe(`menu-conditional-${button.dataset.cfSubmenu}`);
      expect(menu?.querySelector<HTMLElement>(`#${panelId}`)?.dataset.cfPanel).toBe(
        button.dataset.cfSubmenu,
      );
    }
    expect(menu?.querySelectorAll<HTMLElement>('.fc-tb__submenu--cf')).toHaveLength(6);

    tb.dispose();
  });

  it('keeps Conditional Formatting submenus structured as Excel-style panels', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    host.querySelector<HTMLButtonElement>('[data-ribbon-command="conditional"]')?.click();
    const menu = host.querySelector<HTMLElement>('#menu-conditional');
    expect(menu).toBeTruthy();

    const openPanel = (key: string) => {
      const trigger = menu?.querySelector<HTMLElement>(`[data-cf-submenu="${key}"]`);
      const panel = menu?.querySelector<HTMLElement>(`[data-cf-panel="${key}"]`);
      expect(trigger).toBeTruthy();
      expect(panel).toBeTruthy();
      tb.dropdownsApi?.openDynamicConditionalSubmenu(
        menu as HTMLElement,
        key,
        trigger as HTMLElement,
      );
      expect(panel?.hidden).toBe(false);
      expect(panel?.classList.contains('fc-tb__submenu--cf')).toBe(true);
      return panel as HTMLElement;
    };

    const highlightPanel = openPanel('highlight');
    expect(
      Array.from(
        highlightPanel.querySelectorAll<HTMLButtonElement>(':scope > .fc-tb__menu-item--preset'),
      ).map((button) => button.dataset.cfAction),
    ).toEqual([
      'cell-gt',
      'cell-lt',
      'cell-between',
      'cell-eq',
      'text-contains',
      'date-occurring',
      'duplicates',
      'unique',
      'new-rule',
    ]);
    expect(
      highlightPanel.querySelectorAll(':scope > .fc-tb__menu-item--preset .fc-tb__cf-icon'),
    ).toHaveLength(9);

    const dataBarPanel = openPanel('dataBar');
    expect(
      Array.from(dataBarPanel.querySelectorAll<HTMLElement>(':scope > .fc-tb__menu-heading')).map(
        (heading) => heading.textContent,
      ),
    ).toEqual(['Gradient Fill', 'Solid Fill']);
    expect(
      Array.from(dataBarPanel.querySelectorAll<HTMLElement>(':scope > .fc-tb__cf-choice-row')).map(
        (row) => row.querySelectorAll('.fc-tb__cf-choice').length,
      ),
    ).toEqual([6, 6]);
    expect(
      dataBarPanel.querySelector<HTMLButtonElement>('[data-cf-action="new-rule"]'),
    ).toBeTruthy();

    const colorScalePanel = openPanel('colorScale');
    expect(colorScalePanel.querySelector('.fc-tb__cf-choice-grid-panel')).toBeTruthy();
    expect(colorScalePanel.querySelectorAll('.fc-tb__cf-choice')).toHaveLength(12);
    expect(
      Array.from(colorScalePanel.querySelectorAll<HTMLButtonElement>('.fc-tb__cf-choice')).every(
        (button) =>
          !!button.getAttribute('aria-label') && !!button.querySelector('.fc-tb__cf-choice-grid'),
      ),
    ).toBe(true);

    const iconSetPanel = openPanel('iconSet');
    expect(
      Array.from(iconSetPanel.querySelectorAll<HTMLElement>(':scope > .fc-tb__menu-heading')).map(
        (heading) => heading.textContent,
      ),
    ).toEqual(['Directional', 'Shapes', 'Indicators', 'Ratings']);
    expect(iconSetPanel.querySelectorAll(':scope > .fc-tb__cf-icon-panel')).toHaveLength(4);
    expect(iconSetPanel.querySelectorAll('.fc-tb__cf-icon-choice')).toHaveLength(13);

    const clearPanel = openPanel('clear');
    expect(
      Array.from(clearPanel.querySelectorAll<HTMLButtonElement>(':scope > .fc-tb__menu-item')).map(
        (button) => button.dataset.cfAction,
      ),
    ).toEqual(['clear-selection', 'clear-sheet']);
    expect(
      clearPanel.querySelectorAll(':scope > .fc-tb__menu-item .fc-tb__cf-icon--clear'),
    ).toHaveLength(2);

    tb.dispose();
  });

  it('does not dispatch disabled Conditional Formatting actions', () => {
    const applyConditionalMenuAction = vi.fn();
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: () => ({
        ...dynamicDropdownNoopOverrides(),
        applyConditionalMenuAction,
      }),
      helpers: stubHelpers(),
    });

    host.querySelector<HTMLButtonElement>('[data-ribbon-command="conditional"]')?.click();
    const greaterThanButton = host.querySelector<HTMLButtonElement>('[data-cf-action="cell-gt"]');
    expect(greaterThanButton).toBeTruthy();
    if (!greaterThanButton) throw new Error('Expected Conditional Formatting action.');
    greaterThanButton.disabled = true;
    greaterThanButton.setAttribute('aria-disabled', 'true');

    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: greaterThanButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);
    expect(applyConditionalMenuAction).not.toHaveBeenCalled();
    expect(host.querySelector<HTMLElement>('#menu-conditional')?.hidden).toBe(false);

    tb.dispose();
  });

  it('does not dispatch aria-disabled Conditional Formatting actions', () => {
    const applyConditionalMenuAction = vi.fn();
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: () => ({
        ...dynamicDropdownNoopOverrides(),
        applyConditionalMenuAction,
      }),
      helpers: stubHelpers(),
    });

    host.querySelector<HTMLButtonElement>('[data-ribbon-command="conditional"]')?.click();
    const greaterThanButton = host.querySelector<HTMLButtonElement>('[data-cf-action="cell-gt"]');
    expect(greaterThanButton).toBeTruthy();
    if (!greaterThanButton) throw new Error('Expected Conditional Formatting action.');
    greaterThanButton.disabled = false;
    greaterThanButton.setAttribute('aria-disabled', 'true');

    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: greaterThanButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);
    expect(applyConditionalMenuAction).not.toHaveBeenCalled();

    tb.dispose();
  });

  it('does not open disabled Conditional Formatting submenu triggers', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    host.querySelector<HTMLButtonElement>('[data-ribbon-command="conditional"]')?.click();
    const menu = host.querySelector<HTMLElement>('#menu-conditional');
    const trigger = menu?.querySelector<HTMLButtonElement>('[data-cf-submenu="iconSet"]');
    const panel = menu?.querySelector<HTMLElement>('[data-cf-panel="iconSet"]');
    expect(trigger).toBeTruthy();
    expect(panel?.hidden).toBe(true);
    if (!trigger) throw new Error('Expected Conditional Formatting submenu trigger.');
    trigger.disabled = true;
    trigger.setAttribute('aria-disabled', 'true');

    const clickEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(clickEvent, 'target', { value: trigger });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(clickEvent)).toBe(true);
    expect(panel?.hidden).toBe(true);

    const hoverEvent = new MouseEvent('mouseover', { bubbles: true });
    Object.defineProperty(hoverEvent, 'target', { value: trigger });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownHover(hoverEvent)).toBe(true);
    expect(panel?.hidden).toBe(true);

    const keyEvent = new KeyboardEvent('keydown', { bubbles: true, key: 'ArrowRight' });
    Object.defineProperty(keyEvent, 'target', { value: trigger });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownKeydown(keyEvent)).toBe(true);
    expect(panel?.hidden).toBe(true);

    tb.dispose();
  });

  it('keeps every rendered Conditional Formatting action backed by the shared dispatcher', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    host.querySelector<HTMLButtonElement>('[data-ribbon-command="conditional"]')?.click();
    const menu = host.querySelector<HTMLElement>('#menu-conditional');
    expect(menu).toBeTruthy();
    const renderedActions = Array.from(
      menu?.querySelectorAll<HTMLElement>('[data-cf-action]') ?? [],
    )
      .map((item) => item.dataset.cfAction)
      .filter((action): action is string => !!action && !action.startsWith('submenu-'));

    expect(renderedActions.length).toBeGreaterThan(0);
    expect([...new Set(renderedActions)].sort()).toEqual(
      [...SUPPORTED_CONDITIONAL_MENU_ACTIONS].sort(),
    );

    tb.dispose();
  });

  it('flips Conditional Formatting submenus left when the right viewport edge is tight', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    host.querySelector<HTMLButtonElement>('[data-ribbon-command="conditional"]')?.click();
    const menu = host.querySelector<HTMLElement>('#menu-conditional');
    const trigger = menu?.querySelector<HTMLElement>('[data-cf-submenu="iconSet"]');
    const panel = menu?.querySelector<HTMLElement>('[data-cf-panel="iconSet"]');
    expect(menu).toBeTruthy();
    expect(trigger).toBeTruthy();
    expect(panel).toBeTruthy();
    Object.defineProperty(window, 'innerWidth', { configurable: true, value: 900 });
    if (menu) {
      menu.getBoundingClientRect = vi.fn(
        () => ({ left: 560, right: 780, top: 20, bottom: 320, width: 220, height: 300 }) as DOMRect,
      );
    }
    if (trigger) {
      trigger.getBoundingClientRect = vi.fn(
        () => ({ left: 560, right: 780, top: 120, bottom: 146, width: 220, height: 26 }) as DOMRect,
      );
    }
    if (panel) {
      panel.getBoundingClientRect = vi.fn(
        () => ({ left: 0, right: 0, top: 0, bottom: 0, width: 260, height: 240 }) as DOMRect,
      );
    }

    tb.dropdownsApi?.openDynamicConditionalSubmenu(
      menu as HTMLElement,
      'iconSet',
      trigger as HTMLElement,
    );

    expect(panel?.hidden).toBe(false);
    expect(panel?.style.left).toBe('-259px');
    expect(panel?.style.top).toBe('96px');
    expect(menu?.style.overflowY).toBe('');

    tb.dispose();
  });

  it('clamps Conditional Formatting submenus vertically from the shared submenu opener', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    host.querySelector<HTMLButtonElement>('[data-ribbon-command="conditional"]')?.click();
    const menu = host.querySelector<HTMLElement>('#menu-conditional');
    const trigger = menu?.querySelector<HTMLElement>('[data-cf-submenu="iconSet"]');
    const panel = menu?.querySelector<HTMLElement>('[data-cf-panel="iconSet"]');
    expect(menu).toBeTruthy();
    expect(trigger).toBeTruthy();
    expect(panel).toBeTruthy();
    Object.defineProperty(window, 'innerWidth', { configurable: true, value: 1200 });
    Object.defineProperty(window, 'innerHeight', { configurable: true, value: 360 });
    if (menu) {
      menu.getBoundingClientRect = vi.fn(
        () => ({ left: 120, right: 340, top: 20, bottom: 320, width: 220, height: 300 }) as DOMRect,
      );
    }
    if (trigger) {
      trigger.getBoundingClientRect = vi.fn(
        () => ({ left: 120, right: 340, top: 300, bottom: 326, width: 220, height: 26 }) as DOMRect,
      );
    }
    if (panel) {
      panel.getBoundingClientRect = vi.fn(
        () => ({ left: 0, right: 0, top: 0, bottom: 0, width: 260, height: 420 }) as DOMRect,
      );
    }

    tb.dropdownsApi?.openDynamicConditionalSubmenu(
      menu as HTMLElement,
      'iconSet',
      trigger as HTMLElement,
    );

    expect(panel?.hidden).toBe(false);
    expect(panel?.style.left).toBe('219px');
    expect(panel?.style.top).toBe('0px');
    expect(panel?.style.maxHeight).toBe('332px');
    expect(panel?.style.overflowY).toBe('auto');
    expect(panel?.style.overscrollBehavior).toBe('contain');

    tb.dispose();
  });
});
