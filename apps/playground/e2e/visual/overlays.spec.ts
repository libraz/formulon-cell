import { expect, test } from '@playwright/test';

import { mountVisualPage } from './helpers.js';

/** B8 — visual baseline of a stacked-overlay snapshot.
 *
 *  Opens the format dialog (dialog tier) on top of the grid + active selection
 *  (grid tier) so the resulting screenshot captures the runtime composition
 *  of multiple z-index tiers. The image is brittle by nature; it lives in the
 *  visual project (Linux baseline, maxDiffPixels=50) to flag silent layering
 *  regressions across CSS refactors. */
test('@visual stacked overlays — format dialog over the grid', async ({ page }) => {
  await mountVisualPage(page, '/?fixture=selection');

  // Open the format dialog via Mod+1.
  await page
    .locator('.fc-host')
    .first()
    .click({ position: { x: 200, y: 200 } });
  await page.keyboard.press('Control+1');

  // Wait for the dialog to settle.
  await expect(page.locator('[class="fc-fmtdlg"]')).toBeVisible({ timeout: 5_000 });
  await page.waitForTimeout(250);

  // Snapshot the whole viewport — we want the overlap region between grid
  // and dialog visible in the diff.
  await expect(page).toHaveScreenshot('overlays-stacked.png', {
    maxDiffPixels: 200, // generous: the dialog includes anti-aliased glyphs
    animations: 'disabled',
  });
});

test('@visual cell context menu — mini toolbar', async ({ page }) => {
  await mountVisualPage(page, '/?fixture=selection&locale=en');

  await page
    .locator('.fc-host')
    .first()
    .click({ position: { x: 200, y: 200 } });
  await page
    .locator('.fc-host')
    .first()
    .click({ button: 'right', position: { x: 200, y: 200 } });

  const menu = page.locator('.fc-ctxmenu').first();
  await expect(menu).toBeVisible({ timeout: 5_000 });
  await expect(menu.locator('.fc-ctxmenu__mini')).toBeVisible();

  await expect(menu).toHaveScreenshot('overlays-context-menu-mini-toolbar.png', {
    maxDiffPixels: 100,
    animations: 'disabled',
  });
});

test('@visual auto fill options — date menu', async ({ page }) => {
  await mountVisualPage(page, '/?fixture=empty&locale=en');

  await page.evaluate(() => {
    const inst = (
      window as unknown as {
        __fcInst?: {
          workbook: {
            setNumber(addr: { sheet: number; row: number; col: number }, value: number): void;
            recalc(): void;
            cells(sheet: number): Iterable<{
              addr: { sheet: number; row: number; col: number };
              value: unknown;
              formula: string | null;
            }>;
          };
          store: {
            setState(fn: (state: Record<string, unknown>) => Record<string, unknown>): void;
          };
        };
      }
    ).__fcInst;
    if (!inst) throw new Error('missing spreadsheet instance');
    const serial = Date.UTC(2024, 0, 31) / 86_400_000 + 25569;
    inst.workbook.setNumber({ sheet: 0, row: 0, col: 0 }, serial);
    inst.workbook.recalc();
    inst.store.setState((state) => {
      const data = state.data as { cells: Map<string, unknown> };
      const format = state.format as { formats: Map<string, unknown> };
      const cells = new Map<string, unknown>();
      for (const entry of inst.workbook.cells(0)) {
        cells.set(`${entry.addr.sheet}:${entry.addr.row}:${entry.addr.col}`, {
          value: entry.value,
          formula: entry.formula,
        });
      }
      const formats = new Map(format.formats);
      formats.set('0:0:0', { numFmt: { kind: 'date', pattern: 'yyyy-mm-dd' } });
      return { ...state, data: { ...data, cells }, format: { formats } };
    });
    const grid = document.querySelector('.fc-host__grid');
    if (!grid) throw new Error('missing grid');
    grid.dispatchEvent(
      new CustomEvent('fc:autofilloptions', {
        bubbles: true,
        detail: {
          src: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
          dest: { sheet: 0, r0: 0, c0: 0, r1: 5, c1: 0 },
          mode: 'series',
          clientX: 210,
          clientY: 130,
        },
      }),
    );
  });

  const button = page.locator('.fc-autofill-options__button');
  await expect(button).toBeVisible({ timeout: 5_000 });
  await button.click();
  const menu = page.locator('.fc-autofill-options__menu');
  await expect(menu).toBeVisible();
  await expect(menu).toHaveScreenshot('overlays-auto-fill-options-date-menu.png', {
    maxDiffPixels: 100,
    animations: 'disabled',
  });
});

test('@visual paste options — menu', async ({ page }) => {
  await mountVisualPage(page, '/?fixture=empty&locale=en');

  await page.evaluate(() => {
    const host = document.querySelector('.fc-host');
    if (!host) throw new Error('missing host');
    const cell = {
      value: { kind: 'number', value: 7 },
      formula: null,
      format: { bold: true, fill: '#fff2cc' },
    };
    host.dispatchEvent(
      new CustomEvent('fc:pasteoptions', {
        bubbles: true,
        detail: {
          source: {
            range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
            rows: 1,
            cols: 1,
            cells: [[cell]],
          },
          before: {
            range: { sheet: 0, r0: 2, c0: 2, r1: 2, c1: 2 },
            rows: 1,
            cols: 1,
            cells: [[{ ...cell, value: { kind: 'number', value: 2 }, format: undefined }]],
          },
          range: { sheet: 0, r0: 2, c0: 2, r1: 2, c1: 2 },
        },
      }),
    );
  });

  const button = page.locator('.fc-paste-options__button');
  await expect(button).toBeVisible({ timeout: 5_000 });
  await button.evaluate((el) => (el as HTMLButtonElement).click());
  const menu = page.locator('.fc-paste-options__menu');
  await expect(menu).toBeVisible();
  await menu.evaluate((el) => {
    const node = el as HTMLElement;
    node.style.left = '80px';
    node.style.top = '80px';
  });
  const box = await menu.boundingBox();
  if (!box) throw new Error('missing paste options menu bounding box');
  await expect(
    await page.screenshot({
      clip: {
        x: Math.floor(box.x),
        y: Math.floor(box.y),
        width: Math.ceil(box.width),
        height: Math.ceil(box.height),
      },
      animations: 'disabled',
    }),
  ).toMatchSnapshot('overlays-paste-options-menu.png', { maxDiffPixels: 100 });
});

test('@visual quick analysis — button and panel', async ({ page }) => {
  await mountVisualPage(page, '/?fixture=empty&locale=en');

  await page.evaluate(() => {
    const inst = (
      window as unknown as {
        __fcInst?: {
          workbook: {
            setNumber(addr: { sheet: number; row: number; col: number }, value: number): void;
            recalc(): void;
            cells(sheet: number): Iterable<{
              addr: { sheet: number; row: number; col: number };
              value: unknown;
              formula: string | null;
            }>;
          };
          store: {
            setState(fn: (state: Record<string, unknown>) => Record<string, unknown>): void;
          };
        };
      }
    ).__fcInst;
    if (!inst) throw new Error('missing spreadsheet instance');
    [12, 19, 14, 23].forEach((value, col) => {
      inst.workbook.setNumber({ sheet: 0, row: 0, col }, value);
    });
    inst.workbook.recalc();
    inst.store.setState((state) => {
      const data = state.data as { cells: Map<string, unknown> };
      const cells = new Map<string, unknown>();
      for (const entry of inst.workbook.cells(0)) {
        cells.set(`${entry.addr.sheet}:${entry.addr.row}:${entry.addr.col}`, {
          value: entry.value,
          formula: entry.formula,
        });
      }
      return {
        ...state,
        data: { ...data, cells },
        selection: {
          active: { sheet: 0, row: 0, col: 0 },
          anchor: { sheet: 0, row: 0, col: 0 },
          range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 3 },
          extraRanges: [],
        },
      };
    });
  });

  const button = page.locator('.fc-quick__button');
  await expect(button).toBeVisible({ timeout: 5_000 });
  await button.click();
  const panel = page.locator('.fc-quick');
  await expect(panel).toBeVisible();
  await expect(panel).toHaveScreenshot('overlays-quick-analysis-panel.png', {
    maxDiffPixels: 100,
    animations: 'disabled',
  });
});
