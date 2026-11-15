import { expect, type Page } from '@playwright/test';

import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import type { ViewportConfig } from '../../../packages/formulon-cell/src/api/spreadsheet/types.js';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

type DemoWindow = Window & { __fcInst: SpreadsheetInstance };

/** The address the name box shows for the active cell. */
const active = (page: Page) => page.locator('.fc-host__formulabar-tag').inputValue();

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
      inst.host.focus();
    },
    { viewport, start },
  );
  await expect.poll(() => active(page)).toBe(start);
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

/** Whether keyboard focus is on the grid itself (the host element). */
const gridFocused = (page: Page) =>
  page.evaluate(() => document.activeElement === (window as unknown as DemoWindow).__fcInst.host);

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
  // it to the next tab stop (here the name box).
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
  await expect(page.locator('.fc-host__formulabar-tag')).toBeFocused();

  // Cleared, the cursor moves freely again.
  await applyViewport(page, null, 'A1');
  await walk(page, 'ArrowDown', ['A2']);
  await walk(page, 'Tab', ['B2']);
  await sp.expectNoConsoleErrors();
}

/** Whether keyboard focus is on the sheet: the host itself or inside the grid. */
const sheetFocused = (page: Page) =>
  page.evaluate(() => {
    const el = document.activeElement;
    const host = (window as unknown as DemoWindow).__fcInst.host;
    return el === host || !!el?.closest('.fc-host__grid');
  });

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
  expect(await sheetFocused(page), 'Shift+Tab at column A keeps focus on the sheet').toBe(true);

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
  expect(await sheetFocused(page), 'Tab at the last editable cell keeps focus').toBe(true);
  await walk(page, 'Shift+Tab', ['C1', 'A1', 'A1']);
  expect(await sheetFocused(page), 'Shift+Tab at the first editable cell keeps focus').toBe(true);

  await applyViewport(
    page,
    { mode: 'auto', selectable: ['A1:C3'], tabNavigation: 'editable', tabBoundary: 'leave' },
    'C1',
  );
  await walk(page, 'Shift+Tab', ['A1']);
  await page.keyboard.press('Shift+Tab');
  await expect.poll(() => sheetFocused(page), 'Shift+Tab leaves the sheet').toBe(false);
  expect(await active(page)).toBe('A1');
  await sp.expectNoConsoleErrors();
}
