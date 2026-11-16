import { expect, type Page } from '@playwright/test';

import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import type { ViewportConfig } from '../../../packages/formulon-cell/src/api/spreadsheet/types.js';
import { hidePageRibbon } from '../pages/page-ribbon.js';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

type DemoWindow = Window & { __fcInst: SpreadsheetInstance };

/** The address the name box shows for the active cell. */
const active = (page: Page) => page.locator('.fc-host__formulabar-tag').inputValue();

/** The active cell, read from the instance (a grid-only mount has no Name Box). */
const activeCell = (page: Page) =>
  page.evaluate(() => {
    const { row, col } = (window as unknown as DemoWindow).__fcInst.selection.get().active;
    return `${String.fromCharCode(65 + col)}${row + 1}`;
  });

/** Applies `viewport` (or clears it with `null`), then focuses the grid on `start`. */
const applyViewport = async (
  page: Page,
  viewport: ViewportConfig | 'predicate-block-row-2' | null,
  start: string,
): Promise<void> => {
  await page.evaluate(
    async ({ viewport, start }) => {
      const inst = (window as unknown as DemoWindow).__fcInst;
      const next =
        viewport === 'predicate-block-row-2'
          ? {
              mode: 'fixed' as const,
              range: 'A1:H7',
              selectable: ({ addr }: { addr: { row: number } }) => addr.row !== 1,
            }
          : viewport;
      await inst.update({ viewport: next });
      inst.selection.set(start);
      inst.focus();
    },
    { viewport, start },
  );
  await expect.poll(() => activeCell(page)).toBe(start);
};

/** Presses `key` once per expected address and reads the active cell after each. */
const walk = async (page: Page, key: string, expected: readonly string[]): Promise<void> => {
  const seen: string[] = [];
  for (let i = 0; i < expected.length; i++) {
    await page.keyboard.press(key);
    seen.push(await active(page));
  }
  expect(seen, `${key} x${expected.length}`).toEqual(expected);
};

/** Whether keyboard focus is on the grid, the sheet's focus target. */
const gridFocused = (page: Page) =>
  page.evaluate(() => !!document.activeElement?.closest('.fc-host__grid'));

/** Whether keyboard focus sits on an element that follows the grid in document order. */
const focusFollowsGrid = (page: Page) =>
  page.evaluate(() => {
    const grid = document.querySelector('.fc-host__grid');
    const active = document.activeElement;
    if (!grid || !active || active === document.body || grid.contains(active)) return false;
    return !!(grid.compareDocumentPosition(active) & Node.DOCUMENT_POSITION_FOLLOWING);
  });

/** The viewport's `selectable` keeps the cursor on the listed cells, or on the cells its
 *  predicate allows, in either mode, and refuses a move onto any other. Tab walks the fixed range
 *  or the list's bounding range in reading order and stops at its end, where `tabBoundary`
 *  decides whether Tab leaves the grid. Clearing the viewport frees the cursor. */
export async function runViewportNavigationScenario(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ fixture: 'empty' });
  await page.setViewportSize({ width: 1440, height: 900 });
  await sp.expectNoStub();

  // A fixed range with a list: only B2:C3 can be reached; Tab walks it in reading order and
  // stops on the last cell.
  await applyViewport(page, { mode: 'fixed', range: 'A1:H7', selectable: ['B2:C3'] }, 'B2');
  await walk(page, 'Tab', ['C2', 'B3', 'C3', 'C3']);
  // Arrows stop at the list's edge: column A and row 1 are outside it.
  await walk(page, 'ArrowLeft', ['B3', 'B3']);
  await walk(page, 'ArrowUp', ['B2', 'B2']);
  await applyViewport(page, { mode: 'fixed', range: 'A1:H7', selectable: ['B2:C3'] }, 'C3');
  await walk(page, 'ArrowDown', ['C3']);
  await walk(page, 'ArrowRight', ['C3']);

  // A predicate: a move onto row 2 is refused, not skipped.
  await applyViewport(page, 'predicate-block-row-2', 'A1');
  await walk(page, 'ArrowDown', ['A1', 'A1']);
  await walk(page, 'ArrowRight', ['B1']);

  // The whole sheet with a list: Tab walks the list's bounding range and stops at its end.
  await applyViewport(page, { mode: 'auto', selectable: ['B2:C3'] }, 'B2');
  await walk(page, 'Tab', ['C2', 'B3', 'C3', 'C3']);

  // tabBoundary: 'stop' keeps focus on the grid at the last stop; 'leave' lets the browser move
  // it to the next tab stop after the grid.
  await applyViewport(page, { mode: 'fixed', range: 'A1:B1' }, 'A1');
  await page.keyboard.press('Tab');
  await expect.poll(() => active(page)).toBe('B1');
  await page.keyboard.press('Tab');
  expect(await gridFocused(page)).toBe(true);
  await expect.poll(() => active(page)).toBe('B1');
  await applyViewport(page, { mode: 'fixed', range: 'A1:B1', tabBoundary: 'leave' }, 'A1');
  await page.keyboard.press('Tab');
  await expect.poll(() => active(page)).toBe('B1');
  await page.keyboard.press('Tab');
  await expect.poll(() => gridFocused(page)).toBe(false);
  expect(await focusFollowsGrid(page), 'Tab leaves to the stop after the grid').toBe(true);

  // Cleared, the cursor moves freely again.
  await applyViewport(page, null, 'A1');
  await walk(page, 'ArrowDown', ['A2']);
  await walk(page, 'Tab', ['B2']);
  await sp.expectNoConsoleErrors();
}

/** Shift+Tab mirrors Tab. With no viewport it walks left and stops at column A. With
 *  `tabNavigation: 'editable'` both keys move only between the cells the policy lets the user
 *  edit, in reading order, and stop at the first and last of them; with `tabBoundary: 'leave'`
 *  Shift+Tab at the first one moves focus off the sheet and leaves the active cell alone. */
export async function runTabWalkScenario(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ fixture: 'empty' });
  await page.setViewportSize({ width: 1440, height: 900 });
  await sp.expectNoStub();

  await applyViewport(page, null, 'C3');
  await walk(page, 'Shift+Tab', ['B3', 'A3', 'A3']);
  expect(await gridFocused(page), 'Shift+Tab at column A keeps focus on the sheet').toBe(true);

  await page.evaluate(() =>
    (window as unknown as DemoWindow).__fcInst.update({
      policy: { editable: ['A1', 'C1', 'B3'], operations: { valueEdit: true } },
    }),
  );
  await applyViewport(
    page,
    { mode: 'auto', selectable: ['A1:C3'], tabNavigation: 'editable' },
    'A1',
  );
  await walk(page, 'Tab', ['C1', 'B3', 'B3']);
  expect(await gridFocused(page), 'Tab at the last editable cell keeps focus').toBe(true);
  await walk(page, 'Shift+Tab', ['C1', 'A1', 'A1']);
  expect(await gridFocused(page), 'Shift+Tab at the first editable cell keeps focus').toBe(true);

  await applyViewport(
    page,
    { mode: 'auto', selectable: ['A1:C3'], tabNavigation: 'editable', tabBoundary: 'leave' },
    'C1',
  );
  await walk(page, 'Shift+Tab', ['A1']);
  await page.keyboard.press('Shift+Tab');
  await expect.poll(() => gridFocused(page), 'Shift+Tab leaves the sheet').toBe(false);
  expect(await active(page)).toBe('A1');
  await sp.expectNoConsoleErrors();
}

/** Records whether each Tab keydown that reaches the window was handled (default prevented). */
const recordTabHandling = (page: Page) =>
  page.evaluate(() => {
    const w = window as unknown as { __fcTabHandled?: boolean[] };
    w.__fcTabHandled = [];
    window.addEventListener('keydown', (e) => {
      if (e.key === 'Tab') w.__fcTabHandled?.push(e.defaultPrevented);
    });
  });

const tabHandling = (page: Page) =>
  page.evaluate(() => (window as unknown as { __fcTabHandled?: boolean[] }).__fcTabHandled);

/** Tab never traps focus. In the formula bar Tab commits an edit and returns to the grid, with or
 *  without an edit; under a read-only policy the bar leaves Tab to the browser. In a grid-only
 *  mount an unset `tabBoundary` lets Tab leave at the last cell and Shift+Tab at the first, while
 *  an explicit `'stop'` keeps focus in the grid. */
export async function runTabNeverTrapsScenario(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ fixture: 'empty' });
  await page.setViewportSize({ width: 1440, height: 900 });
  await sp.expectNoStub();
  const bar = page.locator('.fc-host__formulabar-input').first();

  await applyViewport(page, null, 'A1');
  await bar.click();
  await page.keyboard.type('5');
  await page.keyboard.press('Tab');
  await expect
    .poll(() => gridFocused(page), 'Tab after an edit in the bar returns to the grid')
    .toBe(true);
  expect
    .soft(await sp.readValue('A1'), 'Tab commits the edit')
    .toEqual({ kind: 'number', value: 5 });

  await applyViewport(page, null, 'C3');
  await bar.click();
  await page.keyboard.press('Tab');
  await expect
    .poll(() => gridFocused(page), 'Tab in the bar without an edit returns to the grid')
    .toBe(true);
  expect
    .soft(await sp.readValue('C3'), 'Tab without an edit writes nothing')
    .toEqual({ kind: 'blank' });

  await page.evaluate(() =>
    (window as unknown as DemoWindow).__fcInst.update({ policy: { readOnly: true } }),
  );
  await recordTabHandling(page);
  await bar.click();
  await page.keyboard.press('Tab');
  expect
    .soft(await tabHandling(page), 'a read-only bar leaves Tab to the browser')
    .toEqual([false]);
  await page.evaluate(() => (window as unknown as DemoWindow).__fcInst.update({ policy: null }));

  await page.evaluate(() =>
    (window as unknown as DemoWindow).__fcInst.update({
      ui: { profile: 'embedded', features: { shortcuts: true } },
    }),
  );
  await hidePageRibbon(page);
  await applyViewport(page, { mode: 'fixed', range: 'A1:B2' }, 'B2');
  await page.keyboard.press('Tab');
  await expect.soft
    .poll(() => gridFocused(page), 'grid-only: Tab at the last cell leaves')
    .toBe(false);
  expect.soft(await activeCell(page)).toBe('B2');
  await applyViewport(page, { mode: 'fixed', range: 'A1:B2' }, 'A1');
  await page.keyboard.press('Shift+Tab');
  await expect
    .poll(() => gridFocused(page), 'grid-only: Shift+Tab at the first cell leaves')
    .toBe(false);
  expect.soft(await activeCell(page)).toBe('A1');
  await applyViewport(page, { mode: 'fixed', range: 'A1:B2', tabBoundary: 'stop' }, 'B2');
  await page.keyboard.press('Tab');
  expect.soft(await gridFocused(page), "grid-only: 'stop' keeps focus at the last cell").toBe(true);
  expect.soft(await activeCell(page)).toBe('B2');
  await sp.expectNoConsoleErrors();
}
