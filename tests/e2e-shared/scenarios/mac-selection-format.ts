import { expect, type Page } from '@playwright/test';
import { a1Address, UserJourneyPage } from '../pages/UserJourneyPage.js';

const targets = ['A1', 'B1', 'C1', 'A2', 'C2', 'A3', 'B3', 'C3', 'E5'];

const engineBold = (page: Page, ref: string): Promise<boolean> =>
  page.evaluate((addr) => {
    const instance = (
      window as unknown as {
        __fcInst: {
          workbook: {
            getCellXfIndex(sheet: number, row: number, col: number): number;
            getCellXf(index: number): { fontIndex: number } | null;
            getFontRecord(index: number): { bold: boolean } | null;
          };
        };
      }
    ).__fcInst;
    const wb = instance.workbook;
    const xf = wb.getCellXf(wb.getCellXfIndex(addr.sheet, addr.row, addr.col));
    if (!xf) throw new Error('Missing engine XF');
    const font = wb.getFontRecord(xf.fontIndex);
    if (!font) throw new Error('Missing engine font');
    return font.bold;
  }, a1Address(ref));

async function commandClick(page: Page, row: number, col: number): Promise<void> {
  const box = await page.locator('.fc-host__grid').boundingBox();
  if (!box) throw new Error('Missing grid');
  await page.keyboard.down('Meta');
  await page.mouse.click(box.x + 26 + 75 * (col + 0.5), box.y + 20 + 20 * (row + 0.5));
  await page.keyboard.up('Meta');
}

const engineBorderStyles = (page: Page, ref: string): Promise<number[]> =>
  page.evaluate((addr) => {
    const wb = (
      window as unknown as {
        __fcInst: {
          workbook: {
            getCellXfIndex(sheet: number, row: number, col: number): number;
            getCellXf(index: number): { borderIndex: number } | null;
            getBorderRecord(index: number): {
              top: { style: number };
              right: { style: number };
              bottom: { style: number };
              left: { style: number };
            } | null;
          };
        };
      }
    ).__fcInst.workbook;
    const xf = wb.getCellXf(wb.getCellXfIndex(addr.sheet, addr.row, addr.col));
    if (!xf) throw new Error('Missing engine XF');
    const border = wb.getBorderRecord(xf.borderIndex);
    if (!border) throw new Error('Missing engine border');
    return [border.top.style, border.right.style, border.bottom.style, border.left.style];
  }, a1Address(ref));

async function selectWithHole(sp: UserJourneyPage, page: Page): Promise<void> {
  await sp.goTo('A1:C3');
  await commandClick(page, 1, 1);
  await commandClick(page, 4, 4);
}

async function historyKey(page: Page, key: string): Promise<void> {
  await page.locator('.fc-host').focus();
  await page.keyboard.press(key);
}

export async function runMacSelectionFormatWithHole(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac' });
  await sp.expectNoStub();
  for (const ref of ['A1', 'B2', 'D4']) {
    await sp.enter(ref, '7');
    await sp.goTo(ref);
    await sp.clickRibbon('bold');
  }
  await sp.enter('C3', '=1+2');
  await selectWithHole(sp, page);
  await sp.clickRibbon('bold');
  for (const ref of [...targets, 'B2', 'D4'])
    await expect.poll(async () => (await sp.readFormat(ref))?.bold).toBe(true);
  for (const ref of targets) await expect.poll(() => engineBold(page, ref)).toBe(true);
  await historyKey(page, 'Meta+z');
  for (const ref of targets)
    await expect.poll(async () => (await sp.readFormat(ref))?.bold ?? false).toBe(ref === 'A1');
  for (const ref of targets) await expect.poll(() => engineBold(page, ref)).toBe(ref === 'A1');
  await historyKey(page, 'Meta+Shift+z');
  for (const ref of targets)
    await expect.poll(async () => (await sp.readFormat(ref))?.bold).toBe(true);
  // A uniformly bold union must turn off together, including additional areas.
  await sp.clickRibbon('bold');
  for (const ref of targets)
    await expect.poll(async () => (await sp.readFormat(ref))?.bold).toBe(false);
  for (const ref of ['B2', 'D4'])
    await expect.poll(async () => (await sp.readFormat(ref))?.bold).toBe(true);
  await sp.clickRibbon('alignC');
  for (const ref of targets)
    await expect.poll(async () => (await sp.readFormat(ref))?.align).toBe('center');
  for (const ref of ['B2', 'D4'])
    await expect.poll(async () => (await sp.readFormat(ref))?.align).toBeUndefined();
  await historyKey(page, 'Meta+z');
  for (const ref of targets)
    await expect.poll(async () => (await sp.readFormat(ref))?.align).toBeUndefined();
  await historyKey(page, 'Meta+i');
  for (const ref of targets)
    await expect.poll(async () => (await sp.readFormat(ref))?.italic).toBe(true);
  await historyKey(page, 'Meta+z');
  for (const ref of targets)
    await expect.poll(async () => (await sp.readFormat(ref))?.italic ?? false).toBe(false);
  await expect.poll(() => sp.readFormula('C3')).toBe('=1+2');
  await expect.poll(() => sp.readValue('C3')).toEqual({ kind: 'number', value: 3 });
  await sp.expectNoConsoleErrors();
}

export async function runMacSelectionClearFormatsWithHole(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac' });
  await sp.expectNoStub();
  await sp.enter('C3', '=1+2');
  await sp.goTo('A1:C3');
  await sp.clickRibbon('bold');
  await sp.enter('E5', '7');
  await sp.goTo('E5');
  await sp.clickRibbon('bold');
  await sp.enter('D4', '99');
  await sp.goTo('D4');
  await sp.clickRibbon('bold');
  await selectWithHole(sp, page);
  await sp.clickRibbon('clearFormat');
  await page.locator('#menu-clear [data-clear="formats"]').click();
  for (const ref of targets)
    await expect.poll(async () => (await sp.readFormat(ref))?.bold ?? false).toBe(false);
  for (const ref of targets) await expect.poll(() => engineBold(page, ref)).toBe(false);
  for (const ref of ['B2', 'D4'])
    await expect.poll(async () => (await sp.readFormat(ref))?.bold).toBe(true);
  await expect.poll(() => sp.readValue('E5')).toEqual({ kind: 'number', value: 7 });
  await expect.poll(() => sp.readFormula('C3')).toBe('=1+2');
  await expect.poll(() => sp.readValue('C3')).toEqual({ kind: 'number', value: 3 });
  await historyKey(page, 'Meta+z');
  for (const ref of targets)
    await expect.poll(async () => (await sp.readFormat(ref))?.bold).toBe(true);
  for (const ref of targets) await expect.poll(() => engineBold(page, ref)).toBe(true);
  await historyKey(page, 'Meta+Shift+z');
  for (const ref of targets)
    await expect.poll(async () => (await sp.readFormat(ref))?.bold ?? false).toBe(false);
  // The primary is unformatted; a formatted additional area enables Clear Formats.
  await sp.goTo('B2');
  await commandClick(page, 5, 5);
  await sp.clickRibbon('clearFormat');
  const clear = page.locator('#menu-clear [data-clear="formats"]');
  await expect(clear).toBeEnabled();
  await clear.click();
  await expect.poll(async () => (await sp.readFormat('B2'))?.bold ?? false).toBe(false);
  await expect.poll(async () => (await sp.readFormat('D4'))?.bold).toBe(true);
  await sp.expectNoConsoleErrors();
}

export async function runMacSelectionFormatRepeat(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac' });
  await sp.expectNoStub();
  for (const ref of ['A1', 'C1', 'A3', 'B3', 'C3']) await sp.enter(ref, '7');
  await sp.goTo('A1');
  await commandClick(page, 0, 2);
  await sp.clickRibbon('alignC');
  for (const ref of ['A1', 'C1'])
    await expect.poll(async () => (await sp.readFormat(ref))?.align).toBe('center');
  await sp.goTo('A3:C3');
  await commandClick(page, 2, 1);
  await historyKey(page, 'F4');
  for (const ref of ['A3', 'C3'])
    await expect.poll(async () => (await sp.readFormat(ref))?.align).toBe('center');
  await expect.poll(async () => (await sp.readFormat('B3'))?.align).toBeUndefined();
  await historyKey(page, 'Meta+z');
  for (const ref of ['A3', 'B3', 'C3'])
    await expect.poll(async () => (await sp.readFormat(ref))?.align).toBeUndefined();
  for (const ref of ['A1', 'C1'])
    await expect.poll(async () => (await sp.readFormat(ref))?.align).toBe('center');
  await sp.expectNoConsoleErrors();
}

export async function runMacSelectionBordersWithHole(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac' });
  await sp.expectNoStub();
  await sp.enter('C3', '=1+2');
  await selectWithHole(sp, page);
  await sp.clickRibbon('borders');
  await page.locator('#menu-borders [data-border-preset="all"]').click();
  for (const ref of targets) {
    await expect
      .poll(async () => (await sp.readFormat(ref))?.borders)
      .toEqual({
        top: { style: 'thin' },
        right: { style: 'thin' },
        bottom: { style: 'thin' },
        left: { style: 'thin' },
      });
  }
  for (const ref of ['B2', 'D4'])
    await expect.poll(async () => (await sp.readFormat(ref))?.borders).toBeUndefined();
  for (const ref of targets)
    await expect.poll(() => engineBorderStyles(page, ref)).toEqual([1, 1, 1, 1]);
  for (const ref of ['B2', 'D4'])
    await expect.poll(() => engineBorderStyles(page, ref)).toEqual([0, 0, 0, 0]);
  await historyKey(page, 'Meta+z');
  for (const ref of targets)
    await expect.poll(async () => (await sp.readFormat(ref))?.borders).toBeUndefined();
  for (const ref of targets)
    await expect.poll(() => engineBorderStyles(page, ref)).toEqual([0, 0, 0, 0]);
  await historyKey(page, 'Meta+Shift+z');
  await expect
    .poll(async () => (await sp.readFormat('E5'))?.borders?.top)
    .toEqual({ style: 'thin' });
  await sp.goTo('A7:C7');
  await commandClick(page, 6, 1);
  await historyKey(page, 'F4');
  for (const ref of ['A7', 'C7'])
    await expect
      .poll(async () => (await sp.readFormat(ref))?.borders?.top)
      .toEqual({ style: 'thin' });
  await expect.poll(async () => (await sp.readFormat('B7'))?.borders).toBeUndefined();
  await historyKey(page, 'Meta+z');
  for (const ref of ['A7', 'B7', 'C7'])
    await expect.poll(async () => (await sp.readFormat(ref))?.borders).toBeUndefined();
  await expect.poll(() => sp.readFormula('C3')).toBe('=1+2');
  await expect.poll(() => sp.readValue('C3')).toEqual({ kind: 'number', value: 3 });
  await sp.expectNoConsoleErrors();
}

export async function runMacSelectionClearAllWithHole(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac' });
  await sp.expectNoStub();
  for (const ref of [...targets, 'B2', 'D4']) await sp.enter(ref, ref === 'C3' ? '=1+2' : '7');
  await sp.goTo('A1:C3');
  await sp.clickRibbon('bold');
  for (const ref of ['E5', 'D4']) {
    await sp.goTo(ref);
    await sp.clickRibbon('bold');
  }
  await selectWithHole(sp, page);
  await sp.clickRibbon('clearFormat');
  await page.locator('#menu-clear [data-clear="all"]').click();
  for (const ref of targets) {
    await expect.poll(() => sp.readValue(ref)).toEqual({ kind: 'blank' });
    await expect.poll(() => sp.readFormula(ref)).toBeNull();
    await expect.poll(() => engineBold(page, ref)).toBe(false);
  }
  for (const ref of ['B2', 'D4']) {
    await expect.poll(() => sp.readValue(ref)).toEqual({ kind: 'number', value: 7 });
    await expect.poll(() => engineBold(page, ref)).toBe(true);
  }
  await historyKey(page, 'Meta+z');
  for (const ref of targets) {
    await expect
      .poll(() => sp.readValue(ref))
      .toEqual({ kind: 'number', value: ref === 'C3' ? 3 : 7 });
    await expect.poll(() => engineBold(page, ref)).toBe(true);
  }
  await expect.poll(() => sp.readFormula('C3')).toBe('=1+2');
  await historyKey(page, 'Meta+Shift+z');
  for (const ref of targets) {
    await expect.poll(() => sp.readValue(ref)).toEqual({ kind: 'blank' });
    await expect.poll(() => engineBold(page, ref)).toBe(false);
  }
  for (const ref of ['B2', 'D4'])
    await expect.poll(() => sp.readValue(ref)).toEqual({ kind: 'number', value: 7 });
  await sp.expectNoConsoleErrors();
}

const engineHorizontalAlign = (page: Page, ref: string): Promise<number> =>
  page.evaluate((addr) => {
    const wb = (
      window as unknown as {
        __fcInst: {
          workbook: {
            getCellXfIndex(sheet: number, row: number, col: number): number;
            getCellXf(index: number): { horizontalAlign: number } | null;
          };
        };
      }
    ).__fcInst.workbook;
    const xf = wb.getCellXf(wb.getCellXfIndex(addr.sheet, addr.row, addr.col));
    if (!xf) throw new Error('Missing engine XF');
    return xf.horizontalAlign;
  }, a1Address(ref));

export async function runMacSelectionFormatDialogWithHole(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac' });
  await sp.expectNoStub();
  for (const ref of [...targets, 'B2', 'D4']) await sp.enter(ref, ref === 'C3' ? '=1+2' : '7');
  for (const ref of ['A1', 'E5', 'B2', 'D4']) {
    await sp.goTo(ref);
    await sp.clickRibbon('bold');
  }
  await selectWithHole(sp, page);
  await historyKey(page, 'Meta+1');
  const dialog = page.getByRole('dialog', { name: 'Format Cells', exact: true });
  await expect(dialog).toBeVisible();
  await dialog.locator('button[data-fc-tab="font"]').click();
  await expect
    .poll(() =>
      dialog
        .locator('input[data-fc-check="bold"]')
        .evaluate((input) => (input as HTMLInputElement).indeterminate),
    )
    .toBe(true);
  await expect(dialog.locator('[data-fc-font-style][aria-selected="true"]')).toHaveCount(0);
  await dialog.locator('button[data-fc-tab="align"]').click();
  await expect(dialog.locator('input[data-fc-check="mergeCells"]')).toBeDisabled();
  await dialog.locator('select[data-fc-select="align"]').selectOption('center');
  await dialog.locator('.fc-fmtdlg__btn--primary').click();
  await expect(dialog).toBeHidden();
  for (const ref of targets) {
    await expect.poll(async () => (await sp.readFormat(ref))?.align).toBe('center');
    await expect.poll(() => engineHorizontalAlign(page, ref)).toBe(2);
    await expect.poll(() => engineBold(page, ref)).toBe(ref === 'A1' || ref === 'E5');
  }
  for (const ref of ['B2', 'D4']) {
    await expect.poll(async () => (await sp.readFormat(ref))?.align).toBeUndefined();
    await expect.poll(() => engineHorizontalAlign(page, ref)).toBe(0);
    await expect.poll(() => engineBold(page, ref)).toBe(true);
  }
  await expect.poll(() => sp.readFormula('C3')).toBe('=1+2');
  await expect.poll(() => sp.readValue('C3')).toEqual({ kind: 'number', value: 3 });
  await historyKey(page, 'Meta+z');
  for (const ref of targets) {
    await expect.poll(async () => (await sp.readFormat(ref))?.align).toBeUndefined();
    await expect.poll(() => engineHorizontalAlign(page, ref)).toBe(0);
  }
  await historyKey(page, 'Meta+Shift+z');
  for (const ref of targets) {
    await expect.poll(async () => (await sp.readFormat(ref))?.align).toBe('center');
    await expect.poll(() => engineHorizontalAlign(page, ref)).toBe(2);
  }
  await sp.goTo('G1:G2');
  await commandClick(page, 0, 8);
  await historyKey(page, 'F4');
  for (const ref of ['G1', 'G2', 'I1']) {
    await expect.poll(async () => (await sp.readFormat(ref))?.align).toBe('center');
    await expect.poll(() => engineHorizontalAlign(page, ref)).toBe(2);
    await expect.poll(() => engineBold(page, ref)).toBe(false);
  }
  await expect.poll(async () => (await sp.readFormat('H1'))?.align).toBeUndefined();
  await expect.poll(() => engineHorizontalAlign(page, 'H1')).toBe(0);
  await historyKey(page, 'Meta+z');
  for (const ref of ['G1', 'G2', 'I1']) {
    await expect.poll(async () => (await sp.readFormat(ref))?.align).toBeUndefined();
    await expect.poll(() => engineHorizontalAlign(page, ref)).toBe(0);
  }
  for (const ref of targets) {
    await expect.poll(async () => (await sp.readFormat(ref))?.align).toBe('center');
    await expect.poll(() => engineHorizontalAlign(page, ref)).toBe(2);
  }
  await sp.expectNoConsoleErrors();
}

const engineFill = (page: Page, ref: string): Promise<[number, number]> =>
  page.evaluate((addr) => {
    const wb = (
      window as unknown as {
        __fcInst: {
          workbook: {
            getCellXfIndex(sheet: number, row: number, col: number): number;
            getCellXf(index: number): { fillIndex: number } | null;
            getFillRecord(index: number): { pattern: number; fgArgb: number } | null;
          };
        };
      }
    ).__fcInst.workbook;
    const xf = wb.getCellXf(wb.getCellXfIndex(addr.sheet, addr.row, addr.col));
    if (!xf) throw new Error('Missing engine XF');
    const fill = wb.getFillRecord(xf.fillIndex);
    if (!fill) throw new Error('Missing engine fill');
    return [fill.pattern, fill.fgArgb & 0xffffff];
  }, a1Address(ref));

export async function runMacSelectionCellStyleWithHole(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac' });
  await sp.expectNoStub();
  for (const ref of [...targets, 'B2', 'D4']) await sp.enter(ref, ref === 'C3' ? '=1+2' : '7');
  for (const ref of ['A1:C3', 'E5', 'D4']) {
    await sp.goTo(ref);
    await sp.clickRibbon('bold');
    await sp.clickRibbon('italic');
    await sp.clickRibbon('alignC');
  }
  await selectWithHole(sp, page);
  await sp.clickRibbon('cellStyles');
  const menu = page.locator('#menu-cell-styles-home');
  await expect(menu.locator('[data-cell-style="normal"]')).toHaveAttribute('aria-checked', 'true');
  await menu.locator('[data-cell-style="good"]').click();
  for (const ref of targets) {
    await expect.poll(async () => (await sp.readFormat(ref))?.cellStyle).toBe('good');
    await expect.poll(async () => (await sp.readFormat(ref))?.bold ?? false).toBe(false);
    await expect.poll(async () => (await sp.readFormat(ref))?.italic ?? false).toBe(false);
    await expect.poll(async () => (await sp.readFormat(ref))?.fill).toBe('#c6efce');
    await expect.poll(() => engineFill(page, ref)).toEqual([1, 0xc6efce]);
    await expect.poll(() => engineBold(page, ref)).toBe(false);
    await expect.poll(() => engineHorizontalAlign(page, ref)).toBe(2);
  }
  for (const ref of ['B2', 'D4']) {
    await expect.poll(() => engineBold(page, ref)).toBe(true);
    await expect.poll(async () => (await sp.readFormat(ref))?.fill).toBeUndefined();
    await expect.poll(async () => (await engineFill(page, ref))[0]).toBe(0);
  }
  await expect.poll(() => sp.readFormula('C3')).toBe('=1+2');
  await sp.clickRibbon('cellStyles');
  await expect(menu.locator('[data-cell-style="good"]')).toHaveAttribute('aria-checked', 'true');
  await page.keyboard.press('Escape');
  await historyKey(page, 'Meta+z');
  for (const ref of targets) {
    await expect.poll(() => engineBold(page, ref)).toBe(true);
    await expect.poll(async () => (await sp.readFormat(ref))?.fill).toBeUndefined();
    await expect.poll(async () => (await engineFill(page, ref))[0]).toBe(0);
  }
  await historyKey(page, 'Meta+Shift+z');
  for (const ref of targets) await expect.poll(() => engineBold(page, ref)).toBe(false);
  await sp.goTo('G1:G2');
  await commandClick(page, 0, 8);
  await historyKey(page, 'F4');
  for (const ref of ['G1', 'G2', 'I1']) {
    await expect.poll(async () => (await sp.readFormat(ref))?.cellStyle).toBe('good');
    await expect.poll(async () => (await sp.readFormat(ref))?.fill).toBe('#c6efce');
    await expect.poll(() => engineFill(page, ref)).toEqual([1, 0xc6efce]);
    await expect.poll(() => engineHorizontalAlign(page, ref)).toBe(0);
  }
  await expect.poll(async () => (await sp.readFormat('H1'))?.fill).toBeUndefined();
  await historyKey(page, 'Meta+z');
  for (const ref of ['G1', 'G2', 'I1'])
    await expect.poll(async () => (await sp.readFormat(ref))?.cellStyle).toBeUndefined();
  for (const ref of targets)
    await expect.poll(async () => (await sp.readFormat(ref))?.cellStyle).toBe('good');
  await sp.expectNoConsoleErrors();
}

export async function runMacSelectionContextClearWithHole(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac' });
  await sp.expectNoStub();
  for (const ref of [...targets, 'B2', 'D4']) await sp.enter(ref, ref === 'C3' ? '=1+2' : '7');
  await sp.goTo('A1:C3');
  await sp.clickRibbon('bold');
  await selectWithHole(sp, page);
  const box = await page.locator('.fc-host__grid').boundingBox();
  if (!box) throw new Error('Missing grid');
  await page.mouse.click(box.x + 26 + 75 / 2, box.y + 20 + 20 / 2, { button: 'right' });
  const menu = page.locator('.fc-ctxmenu:not(.fc-ctxmenu__sub)');
  await expect(menu).toBeVisible();
  await menu.locator('[data-fc-action="clear"]').click();
  await expect(menu).toBeHidden();
  for (const ref of targets) {
    await expect.poll(() => sp.readValue(ref)).toEqual({ kind: 'blank' });
    await expect.poll(() => sp.readFormula(ref)).toBeNull();
  }
  for (const ref of ['B2', 'D4'])
    await expect.poll(() => sp.readValue(ref)).toEqual({ kind: 'number', value: 7 });
  await expect.poll(() => engineBold(page, 'A1')).toBe(true);
  await historyKey(page, 'Meta+z');
  for (const ref of targets)
    await expect
      .poll(() => sp.readValue(ref))
      .toEqual({
        kind: 'number',
        value: ref === 'C3' ? 3 : 7,
      });
  await expect.poll(() => sp.readFormula('C3')).toBe('=1+2');
  await historyKey(page, 'Meta+Shift+z');
  for (const ref of targets) await expect.poll(() => sp.readValue(ref)).toEqual({ kind: 'blank' });
  for (const ref of ['B2', 'D4'])
    await expect.poll(() => sp.readValue(ref)).toEqual({ kind: 'number', value: 7 });
  await expect.poll(() => engineBold(page, 'A1')).toBe(true);
  await sp.expectNoConsoleErrors();
}

export async function runMacSelectionAccentStyles(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ locale: 'ja', platform: 'mac' });
  await sp.expectNoStub();
  for (const ref of ['A1', 'B1', 'C1']) await sp.enter(ref, '7');

  await sp.goTo('A1');
  await sp.clickRibbon('cellStyles');
  const menu = page.locator('#menu-cell-styles-home').first();
  await expect(menu).toBeVisible();
  await expect(menu.locator('[data-cell-style]')).toHaveCount(47);
  await expect(menu.locator('[data-cell-style="accent1_40"]')).toHaveAttribute(
    'aria-label',
    '40% - アクセント 1',
  );
  await expect(menu.locator('[data-cell-style="accent6_60"]')).toHaveAttribute(
    'aria-label',
    '60% - アクセント 6',
  );

  const finalAccent = menu.locator('[data-cell-style="accent6_60"]');
  const readAccentGeometry = async (): Promise<{
    clientHeight: number;
    inPageViewport: boolean;
    inScrollViewport: boolean;
    scrollHeight: number;
    scrollTop: number;
  }> =>
    finalAccent.evaluate((element) => {
      const scroll = element.closest<HTMLElement>('.fc-tb__cellstyle-scroll');
      const chip = element.getBoundingClientRect();
      const body = scroll?.getBoundingClientRect();
      const viewport = element.ownerDocument.defaultView;
      const inScrollViewport =
        body !== undefined &&
        chip.top >= body.top &&
        chip.bottom <= body.bottom &&
        chip.left >= body.left &&
        chip.right <= body.right;
      const inPageViewport =
        viewport !== null &&
        chip.top >= 0 &&
        chip.bottom <= viewport.innerHeight &&
        chip.left >= 0 &&
        chip.right <= viewport.innerWidth;
      return {
        clientHeight: scroll?.clientHeight ?? 0,
        inPageViewport,
        inScrollViewport,
        scrollHeight: scroll?.scrollHeight ?? 0,
        scrollTop: scroll?.scrollTop ?? 0,
      };
    });
  const beforeFocus = await readAccentGeometry();
  await finalAccent.focus();
  await expect(finalAccent).toBeFocused();
  await expect.poll(async () => (await readAccentGeometry()).inScrollViewport).toBe(true);
  await expect.poll(async () => (await readAccentGeometry()).inPageViewport).toBe(true);
  const afterFocus = await readAccentGeometry();
  expect(afterFocus.scrollHeight).toBeGreaterThanOrEqual(afterFocus.clientHeight);
  if (!beforeFocus.inScrollViewport) {
    await expect
      .poll(async () => (await readAccentGeometry()).scrollTop)
      .toBeGreaterThan(beforeFocus.scrollTop);
  }

  await menu.locator('[data-cell-style="accent1_40"]').click();
  await expect.poll(async () => (await sp.readFormat('A1'))?.fill).toBe('#83cceb');
  await expect.poll(() => engineFill(page, 'A1')).toEqual([1, 0x83cceb]);

  await sp.goTo('B1');
  await historyKey(page, 'F4');
  await expect.poll(async () => (await sp.readFormat('B1'))?.fill).toBe('#83cceb');
  await expect.poll(() => engineFill(page, 'B1')).toEqual([1, 0x83cceb]);

  await sp.goTo('C1');
  await historyKey(page, 'F4');
  await expect.poll(async () => (await sp.readFormat('C1'))?.fill).toBe('#83cceb');
  await expect.poll(() => engineFill(page, 'C1')).toEqual([1, 0x83cceb]);

  await historyKey(page, 'Meta+z');
  await expect.poll(async () => (await sp.readFormat('B1'))?.fill).toBe('#83cceb');
  await expect.poll(async () => (await sp.readFormat('C1'))?.fill).toBeUndefined();
  await expect.poll(() => engineFill(page, 'C1')).toEqual([0, 0]);
  await historyKey(page, 'Meta+Shift+z');
  await expect.poll(async () => (await sp.readFormat('C1'))?.fill).toBe('#83cceb');
  await expect.poll(() => engineFill(page, 'C1')).toEqual([1, 0x83cceb]);
  await sp.expectNoConsoleErrors();
}
