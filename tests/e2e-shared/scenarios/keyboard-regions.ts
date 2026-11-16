import type { Page } from '@playwright/test';
import { expect } from '@playwright/test';

import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { hidePageRibbon } from '../pages/page-ribbon.js';
import { selectRibbonTab, setRibbonDisplayMode } from '../pages/ribbon-tabs.js';
import { gridCellCenter, SpreadsheetPage } from '../pages/SpreadsheetPage.js';

type Platform = 'default' | 'mac';
type InstanceWindow = Window & { __fcInst: SpreadsheetInstance };

/** Which docked panes are open while F6 cycles, or an auto-hidden ribbon with none. */
export type RegionCycleVariant = 'none' | 'watch' | 'pivotFields' | 'autoHide' | 'noRibbon';

/** The keyboard region holding focus: `ribbon` (the active tab), `namebox`, `grid`, `sheetbar`
 *  (the active sheet tab), `statusbar`, `pane:<label>`, `body`, or a short description. */
const region = (page: Page) =>
  page.evaluate(() => {
    const active = document.activeElement as HTMLElement | null;
    if (!active || active === document.body || active === document.documentElement) return 'body';
    if (active.matches('[data-ribbon-tab][aria-selected="true"]')) return 'ribbon';
    if (active.matches('.fc-host__formulabar-tag')) return 'namebox';
    if (active.closest('.fc-host__grid')) return 'grid';
    if (active.matches('.fc-host__sheetbar-tab[aria-selected="true"]')) return 'sheetbar';
    if (active.closest('.fc-host__statusbar')) return 'statusbar';
    const pane = active.closest<HTMLElement>('[data-fc-surface="pane"]');
    if (pane) return `pane:${pane.getAttribute('aria-label') ?? pane.className.split(' ')[0]}`;
    return `${active.tagName.toLowerCase()}[${active.getAttribute('aria-label') ?? active.className}]`;
  });

/** The docked panes showing now, by label, in document order. */
const openPanes = (page: Page) =>
  page.evaluate(() =>
    Array.from(document.querySelectorAll<HTMLElement>('[data-fc-surface="pane"]'))
      .filter((el) => !el.hidden && el.getClientRects().length > 0)
      .map((el) => `pane:${el.getAttribute('aria-label') ?? el.className.split(' ')[0]}`),
  );

/** Presses `key` `count` times from the current focus and lists the region after each press. */
const cycle = async (page: Page, key: string, count: number): Promise<string[]> => {
  const seen: string[] = [];
  for (let i = 0; i < count; i += 1) {
    await page.keyboard.press(key);
    seen.push(await region(page));
  }
  return seen;
};

const focusGrid = async (page: Page, row = 1, col = 3): Promise<void> => {
  const cell = await gridCellCenter(page, row, col);
  await page.mouse.click(cell.x, cell.y);
  await expect.poll(() => region(page)).toBe('grid');
};

/** Seeds a two-field PivotTable and leaves its PivotTable Fields list docked. */
async function dockPivotFields(page: Page, platform: Platform): Promise<void> {
  await page.evaluate(() => {
    const inst = (window as unknown as InstanceWindow).__fcInst;
    inst.workbook.batch(() => {
      const sheet = inst.workbook.sheet(0);
      sheet.setText(0, 0, 'Channel');
      sheet.setText(0, 1, 'Sales');
      sheet.setText(1, 0, 'Online');
      sheet.setNumber(1, 1, 10);
      sheet.setText(2, 0, 'Retail');
      sheet.setNumber(2, 1, 20);
    });
    inst.history.clear();
    inst.selection.set('A1:B3');
  });
  await selectRibbonTab(page, 'insert');
  await page
    .locator(
      platform === 'mac'
        ? '[data-ribbon-command="mac.insert.pivotTable"]'
        : '[data-ribbon-command="pivotTableInsert"]',
    )
    .first()
    .click();
  const create = page.getByRole('dialog', { name: 'Create PivotTable' });
  await create.getByRole('button', { name: 'OK', exact: true }).click();
  await expect(create).toBeHidden();
  for (const field of ['Channel', 'Sales']) {
    await page.getByRole('checkbox', { name: field, exact: true }).first().check();
  }
}

/** F6 moves focus through the spreadsheet's regions in a fixed ring: the ribbon (title bar
 *  included) at its active tab, the Name Box, the grid, the active sheet tab, the status bar, then
 *  each open docked pane in document order, and wraps around. Shift+F6 walks the same ring
 *  backwards. A region that is not rendered takes no stop; an auto-hidden ribbon whose tab strip
 *  is on screen still takes its stop at the active tab. */
export async function runRegionCycleScenario(
  page: Page,
  platform: Platform,
  variant: RegionCycleVariant,
): Promise<void> {
  const sp = new SpreadsheetPage(page);
  const consoleErrors = sp.collectConsoleErrors();
  await sp.mount({ platform });
  await sp.expectNoStub();
  if (variant === 'watch') {
    await focusGrid(page);
    await page.evaluate(() =>
      (window as unknown as InstanceWindow).__fcInst.dialogs.open('watchWindow'),
    );
  } else if (variant === 'pivotFields') {
    await dockPivotFields(page, platform);
  } else if (variant === 'autoHide') {
    await setRibbonDisplayMode(page, 'autoHide');
  } else if (variant === 'noRibbon') {
    await hidePageRibbon(page);
  }
  const panes = await openPanes(page);
  expect(panes.length, 'panes docked for the variant').toBe(
    variant === 'watch' || variant === 'pivotFields' ? 1 : 0,
  );

  // A4 holds the pivot's first row item, which keeps its field list showing.
  const startOnGrid = () => (variant === 'pivotFields' ? focusGrid(page, 3, 0) : focusGrid(page));
  await startOnGrid();
  const ring = [
    ...(variant === 'noRibbon' ? [] : ['ribbon']),
    'namebox',
    'grid',
    'sheetbar',
    'statusbar',
    ...panes,
  ];
  const from = ring.indexOf('grid');
  const forward = ring.map((_, i) => ring[(from + 1 + i) % ring.length] as string);
  expect
    .soft(await cycle(page, 'F6', ring.length), 'F6 walks the ring from the grid')
    .toEqual(forward);
  const backward = ring.map(
    (_, i) => ring[(from - 1 - i + 2 * ring.length) % ring.length] as string,
  );
  await startOnGrid();
  expect
    .soft(await cycle(page, 'Shift+F6', ring.length), 'Shift+F6 walks it back')
    .toEqual(backward);
  expect(consoleErrors.read()).toEqual([]);
}

/** In a mount whose only showing region is the grid, F6 and Shift+F6 are left to the browser:
 *  the spreadsheet does not handle them and focus stays in the grid. */
export async function runGridOnlyRegionScenario(page: Page, platform: Platform): Promise<void> {
  const sp = new SpreadsheetPage(page);
  const consoleErrors = sp.collectConsoleErrors();
  await sp.mount({ platform });
  await sp.expectNoStub();
  await page.evaluate(() => {
    (window as unknown as InstanceWindow).__fcInst.update({ ui: { profile: 'embedded' } });
  });
  await hidePageRibbon(page);
  await page.evaluate(() => {
    const w = window as unknown as { __fcF6Handled?: boolean[] };
    w.__fcF6Handled = [];
    window.addEventListener('keydown', (e) => {
      if (e.key === 'F6') w.__fcF6Handled?.push(e.defaultPrevented);
    });
  });
  for (const key of ['F6', 'Shift+F6']) {
    await focusGrid(page, 3, 2);
    await page.keyboard.press(key);
    expect.soft(await region(page), `${key} leaves focus in the grid`).toBe('grid');
  }
  expect(
    await page.evaluate(() => (window as unknown as { __fcF6Handled?: boolean[] }).__fcF6Handled),
    'F6 and Shift+F6 are not handled',
  ).toEqual([false, false]);
  expect(consoleErrors.read()).toEqual([]);
}

const dialogHoldsFocus = (page: Page) =>
  page.evaluate(() => !!document.activeElement?.closest('[data-fc-surface="modal"]'));

/** Whether the active ribbon tab's panel shows (the ribbon is not collapsed to its tabs). */
const ribbonExpanded = (page: Page) =>
  page.evaluate(() => {
    const tab = document.querySelector<HTMLElement>('[data-ribbon-tab][aria-selected="true"]');
    const id = tab?.dataset.ribbonTab;
    const panel = id ? document.querySelector<HTMLElement>(`[data-ribbon-panel="${id}"]`) : null;
    return !!panel && !panel.hidden && panel.getClientRects().length > 0;
  });

/** Whether focus is on a text field outside the spreadsheet host, as the title-bar search. */
const titleSearchFocused = (page: Page) =>
  page.evaluate(() => {
    const active = document.activeElement;
    return (
      active instanceof HTMLInputElement &&
      !active.closest('.fc-host') &&
      !active.closest('[data-fc-surface]')
    );
  });

const openModal = async (page: Page): Promise<void> => {
  await focusGrid(page);
  await page.evaluate(() => (window as unknown as InstanceWindow).__fcInst.dialogs.open('format'));
  await expect.poll(() => dialogHoldsFocus(page)).toBe(true);
};

/** A modal dialog owns the keyboard: the region keys (F6, Shift+F6), the ribbon keys (Ctrl+F1,
 *  Alt with an auto-hidden ribbon) and the search key (Alt+Q, or Cmd+Ctrl+U on the Mac) do
 *  nothing while it is open, its own keys still work, and the keys work again once it closes.
 *  Escape closes the File backstage onto the File tab; over a dialog opened from a backstage
 *  command it closes only the dialog, which hands focus back to that command. */
export async function runModalKeyOwnershipScenario(page: Page, platform: Platform): Promise<void> {
  const sp = new SpreadsheetPage(page);
  const consoleErrors = sp.collectConsoleErrors();
  await sp.mount({ platform });
  await sp.expectNoStub();
  const searchKey = platform === 'mac' ? 'Meta+Control+u' : 'Alt+q';

  await openModal(page);
  for (const key of ['F6', 'Shift+F6', 'Control+F1', searchKey]) {
    await page.keyboard.press(key);
    expect.soft(await dialogHoldsFocus(page), `${key} leaves focus in the dialog`).toBe(true);
    expect.soft(await ribbonExpanded(page), `${key} leaves the ribbon as it was`).toBe(true);
  }
  await page.keyboard.press('Tab');
  expect.soft(await dialogHoldsFocus(page), 'Tab moves within the dialog').toBe(true);
  await page.keyboard.press('Escape');
  await expect.soft.poll(() => dialogHoldsFocus(page), 'Escape closes the dialog').toBe(false);

  await page.keyboard.press('Control+F1');
  await expect.soft.poll(() => ribbonExpanded(page), 'Ctrl+F1 collapses the ribbon').toBe(false);
  await page.keyboard.press('Control+F1');
  await expect.soft.poll(() => ribbonExpanded(page), 'Ctrl+F1 expands it again').toBe(true);
  await focusGrid(page);
  await page.keyboard.press(searchKey);
  await expect.soft.poll(() => titleSearchFocused(page), `${searchKey} focuses search`).toBe(true);

  if (platform === 'default') {
    await setRibbonDisplayMode(page, 'autoHide');
    await openModal(page);
    await page.keyboard.press('Alt');
    expect.soft(await dialogHoldsFocus(page), 'Alt leaves focus in the dialog').toBe(true);
    await page.keyboard.press('Escape');
    await focusGrid(page);
    await page.keyboard.press('Alt');
    await expect.soft
      .poll(() => region(page), 'Alt shows the auto-hidden ribbon at its active tab')
      .toBe('ribbon');
    await page.keyboard.press('Escape');
    await setRibbonDisplayMode(page, 'full');

    const fileTab = page.locator('[data-ribbon-tab="file"]');
    const backstage = page.locator('.fc-tb__backstage-navitem').first();
    await fileTab.click();
    await expect(backstage).toBeVisible();
    await page.keyboard.press('Escape');
    await expect.soft(backstage, 'Escape closes the backstage').toBeHidden();
    await expect.soft(fileTab, 'and focuses the File tab').toBeFocused();
    if (!(await backstage.isVisible())) await fileTab.click();
    await expect(backstage).toBeVisible();
    const card = page.locator('.fc-tb__backstage-card').filter({ hasText: 'Page Setup' });
    await card.focus();
    await page.keyboard.press('Enter');
    const pageSetup = page.getByRole('dialog', { name: 'Page Setup' });
    await expect(pageSetup).toBeVisible();
    await page.keyboard.press('Escape');
    await expect.soft(pageSetup, 'Escape closes Page Setup').toBeHidden();
    await expect.soft(backstage, 'Escape over the backstage leaves it open').toBeVisible();
    await expect.soft(card, 'focus returns to the backstage command that opened it').toBeFocused();
    await page.keyboard.press('Escape');
    await expect.soft(backstage, 'a second Escape closes the backstage').toBeHidden();
    await expect.soft(fileTab, 'and focuses the File tab again').toBeFocused();
  }
  expect(consoleErrors.read()).toEqual([]);
}

/** The sheet's tab stop is the grid, not the mount element: the mount has no tabindex, Tab from
 *  the formula bar enters the grid, and Tab past the last stop of a `tabBoundary: 'leave'`
 *  viewport moves to the next tab stop after the grid. */
export async function runSheetFocusTargetScenario(page: Page, platform: Platform): Promise<void> {
  const sp = new SpreadsheetPage(page);
  const consoleErrors = sp.collectConsoleErrors();
  await sp.mount({ platform });
  await sp.expectNoStub();
  const host = page.locator('.fc-host').first();
  expect.soft(await host.getAttribute('tabindex'), 'the mount element has no tabindex').toBeNull();

  await page.locator('.fc-host__formulabar-input').first().focus();
  await page.keyboard.press('Tab');
  await expect.soft
    .poll(() => region(page), 'Tab from the formula bar enters the grid')
    .toBe('grid');

  await page.evaluate(async () => {
    const inst = (window as unknown as InstanceWindow).__fcInst;
    await inst.update({ viewport: { mode: 'fixed', range: 'A1:B1', tabBoundary: 'leave' } });
    inst.selection.set('A1');
  });
  await focusGrid(page, 0, 0);
  await page.keyboard.press('Tab');
  await page.keyboard.press('Tab');
  expect
    .soft(
      await page.evaluate(() => {
        const grid = document.querySelector('.fc-host__grid');
        const active = document.activeElement;
        if (!grid || !active || active === document.body || grid.contains(active)) return false;
        return !!(grid.compareDocumentPosition(active) & Node.DOCUMENT_POSITION_FOLLOWING);
      }),
      "Tab past the last stop with 'leave' moves to the next stop after the grid",
    )
    .toBe(true);
  expect(consoleErrors.read()).toEqual([]);
}
