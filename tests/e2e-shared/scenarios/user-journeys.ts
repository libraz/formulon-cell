import { expect, type Page, test } from '@playwright/test';

import { UserJourneyPage } from '../pages/UserJourneyPage.js';

const start = async (page: Page): Promise<UserJourneyPage> => {
  const sp = new UserJourneyPage(page);
  await sp.mount({ fixture: 'empty' });
  await sp.expectNoStub();
  return sp;
};

const expectValue = async (
  sp: UserJourneyPage,
  ref: string,
  value: { kind: string; value?: unknown },
  sheet = 0,
): Promise<void> => {
  await expect.poll(() => sp.readValue(ref, sheet), { timeout: 2_000 }).toEqual(value);
};

const expectFormula = async (
  sp: UserJourneyPage,
  ref: string,
  formula: string,
  sheet = 0,
): Promise<void> => {
  await expect.poll(() => sp.readFormula(ref, sheet), { timeout: 2_000 }).toBe(formula);
};

const chooseCurrency = async (sp: UserJourneyPage): Promise<void> => {
  const menu = await sp.openRibbonSplitMenu('currency', 'menu-currency-home');
  await menu.locator('[data-currency-preset="$"]').click();
};

const chooseMergeAction = async (
  sp: UserJourneyPage,
  action: 'mergeCells' | 'unmergeCells',
): Promise<void> => {
  const menu = await sp.openRibbonSplitMenu('merge', 'menu-merge');
  await menu.locator(`[data-merge-action="${action}"]`).click();
};

/** A household budget: enter amounts, sum them, format the money, then revise
 * a precedent while checking undo and redo restore the exact totals. */
export async function runBudgetEntryScenario(page: Page): Promise<void> {
  const sp = await start(page);

  await test.step('enter budget labels and amounts', async () => {
    await sp.enter('A1', 'Rent');
    await sp.enter('A2', 'Food');
    await sp.enter('B1', '1200');
    await sp.enter('B2', '300');
    await sp.enter('B3', '=SUM(B1:B2)');
    await expectValue(sp, 'B3', { kind: 'number', value: 1500 });
    await expectFormula(sp, 'B3', '=SUM(B1:B2)');
  });

  await test.step('apply currency formatting from the ribbon', async () => {
    await sp.goTo('B1:B3');
    await chooseCurrency(sp);
    await expect
      .poll(() => sp.readFormat('B3'), { timeout: 2_000 })
      .toEqual(expect.objectContaining({ numFmt: expect.objectContaining({ kind: 'currency' }) }));
  });

  await test.step('edit a precedent and recompute the sum', async () => {
    await sp.enter('B2', '450');
    await expectValue(sp, 'B2', { kind: 'number', value: 450 });
    await expectValue(sp, 'B3', { kind: 'number', value: 1650 });
    await expectFormula(sp, 'B3', '=SUM(B1:B2)');
  });

  await test.step('undo and redo the precedent edit', async () => {
    await sp.shortcut('z');
    await expectValue(sp, 'B2', { kind: 'number', value: 300 });
    await expectValue(sp, 'B3', { kind: 'number', value: 1500 });
    await sp.shortcut('y');
    await expectValue(sp, 'B2', { kind: 'number', value: 450 });
    await expectValue(sp, 'B3', { kind: 'number', value: 1650 });
  });
  await sp.expectNoConsoleErrors();
}

/** A small order list: calculate a line total, then fill the relative formula
 * down for a second item. */
export async function runOrderFillScenario(page: Page): Promise<void> {
  const sp = await start(page);

  await test.step('enter order quantities, prices, and first total', async () => {
    await sp.enter('A1', 'Item');
    await sp.enter('B1', 'Qty');
    await sp.enter('C1', 'Unit price');
    await sp.enter('D1', 'Total');
    await sp.enter('A2', 'Notebook');
    await sp.enter('B2', '2');
    await sp.enter('C2', '12.5');
    await sp.enter('D2', '=B2*C2');
    await sp.enter('A3', 'Pen');
    await sp.enter('B3', '3');
    await sp.enter('C3', '7.5');
    await expectValue(sp, 'D2', { kind: 'number', value: 25 });
    await expectFormula(sp, 'D2', '=B2*C2');
  });

  await test.step('fill the line total formula down', async () => {
    await sp.goTo('D2:D3');
    await sp.clickRibbon('fillHome');
    await sp.page.locator('#menu-fill [data-fill="down"]').click();
    await expectValue(sp, 'D3', { kind: 'number', value: 22.5 });
    await expectFormula(sp, 'D3', '=B3*C3');
  });
  await test.step('reuse a line formula for a new item with relative references', async () => {
    await sp.enter('A4', 'Folder');
    await sp.enter('B4', '5');
    await sp.enter('C4', '4');
    await sp.goTo('D2');
    await sp.shortcut('c');
    await sp.goTo('D4');
    await sp.shortcut('v');
    await expectFormula(sp, 'D4', '=B4*C4');
    await expectValue(sp, 'D4', { kind: 'number', value: 20 });
    await sp.enter('B4', '6');
    await expectValue(sp, 'D4', { kind: 'number', value: 24 });
  });
  await sp.expectNoConsoleErrors();
}

/** A contact list: add a row, attach a link through the dialog, and edit the
 * link after the inserted row has shifted the original contact. */
export async function runContactListScenario(page: Page): Promise<void> {
  const sp = await start(page);

  await test.step('enter the initial contacts', async () => {
    await sp.enter('A1', 'Name');
    await sp.enter('B1', 'Email');
    await sp.enter('A2', 'Alice');
    await sp.enter('B2', 'alice@example.com');
    await sp.enter('A3', 'Bob');
    await sp.enter('B3', 'bob@example.com');
  });

  await test.step('link Alice before inserting a row', async () => {
    await sp.goTo('B2');
    const dialog = await sp.openHyperlinkDialog();
    await dialog.locator('input[type="url"]').fill('https://example.com/old-alice');
    await dialog.getByRole('button', { name: 'OK', exact: true }).click();
    await expect
      .poll(() => sp.readFormat('B2'), { timeout: 2_000 })
      .toEqual(expect.objectContaining({ hyperlink: 'https://example.com/old-alice' }));
  });

  await test.step('insert a row above Alice', async () => {
    await sp.goTo('A2');
    await sp.insertRowsOrColumns('rows');
    await expectValue(sp, 'A3', { kind: 'text', value: 'Alice' });
    await expectValue(sp, 'B3', { kind: 'text', value: 'alice@example.com' });
    await expectValue(sp, 'A4', { kind: 'text', value: 'Bob' });
    await expect
      .poll(() => sp.readFormat('B3'), { timeout: 2_000 })
      .toEqual(expect.objectContaining({ hyperlink: 'https://example.com/old-alice' }));
  });

  await test.step('add and link the new contact', async () => {
    await sp.enter('A2', 'Carol');
    await sp.enter('B2', 'carol@example.com');
    await sp.goTo('B2');
    const dialog = await sp.openHyperlinkDialog();
    await dialog.locator('input[type="url"]').fill('https://example.com/carol');
    await dialog.getByRole('button', { name: 'OK', exact: true }).click();
    await expect
      .poll(() => sp.readFormat('B2'), { timeout: 2_000 })
      .toEqual(expect.objectContaining({ hyperlink: 'https://example.com/carol' }));
  });

  await test.step('edit Alice link after the row move', async () => {
    await sp.goTo('B3');
    const dialog = await sp.openHyperlinkDialog();
    await expect(dialog.locator('input[type="url"]')).toHaveValue('https://example.com/old-alice');
    await dialog.locator('input[type="url"]').fill('https://example.com/alice');
    await dialog.getByRole('button', { name: 'OK', exact: true }).click();
    await expect
      .poll(() => sp.readFormat('B3'), { timeout: 2_000 })
      .toEqual(expect.objectContaining({ hyperlink: 'https://example.com/alice' }));
  });
  await sp.expectNoConsoleErrors();
}

/** A report header exercises the Format Cells dialog's font, alignment, and
 * merge controls, then uses the ribbon menu to unmerge it again. */
export async function runReportHeaderScenario(page: Page): Promise<void> {
  const sp = await start(page);

  await test.step('enter a report header', async () => {
    await sp.enter('A1', 'Monthly report');
    await sp.goTo('A1:C1');
  });

  await test.step('format, center, and merge the header', async () => {
    const dialog = await sp.openFormatDialog();
    await dialog.locator('[data-fc-tab="font"][role="tab"]').click();
    await dialog.locator('[data-fc-input="family"]').fill('Arial');
    await dialog.locator('.fc-fmtdlg__font-size-row input').fill('16');
    await dialog.locator('[data-fc-check="bold"]').check();
    await dialog.locator('[data-fc-tab="align"][role="tab"]').click();
    await dialog.locator('[data-fc-select="align"]').selectOption('center');
    await dialog.locator('[data-fc-check="mergeCells"]').check();
    await dialog.getByRole('button', { name: 'OK', exact: true }).click();

    await expect
      .poll(() => sp.readFormat('A1'), { timeout: 2_000 })
      .toEqual(
        expect.objectContaining({ bold: true, fontFamily: 'Arial', fontSize: 16, align: 'center' }),
      );
    await expect
      .poll(() => sp.readMerge('A1'), { timeout: 2_000 })
      .toEqual({ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 2 });
  });

  await test.step('unmerge the header through the ribbon menu', async () => {
    await chooseMergeAction(sp, 'unmergeCells');
    await expect.poll(() => sp.readMerge('A1'), { timeout: 2_000 }).toBeNull();
  });
  await sp.expectNoConsoleErrors();
}

/** Build a summary from a range that includes a merged numeric cell. */
export async function runMergedSummaryScenario(page: Page): Promise<void> {
  const sp = await start(page);

  await test.step('create a merged opening balance above two amounts', async () => {
    await sp.enter('A1', '100');
    await sp.enter('A2', '20');
    await sp.enter('B2', '30');
    await sp.goTo('A1:B1');
    await chooseMergeAction(sp, 'mergeCells');
    await expect
      .poll(() => sp.readMerge('A1'))
      .toEqual({
        sheet: 0,
        r0: 0,
        c0: 0,
        r1: 0,
        c1: 1,
      });
    await expectValue(sp, 'A1', { kind: 'number', value: 100 });
  });

  await test.step('navigate past the merged cell and select the complete summary range', async () => {
    await sp.goTo('A1:B1');
    await page.keyboard.press('ArrowRight');
    await expect(page.getByRole('textbox', { name: 'Name box', exact: true })).toHaveValue('C1');
    await page.keyboard.press('ArrowLeft');
    await expect(page.getByRole('textbox', { name: 'Name box', exact: true })).toHaveValue('A1');
    await page.keyboard.press('Shift+ArrowDown');
    await expect(page.getByRole('textbox', { name: 'Name box', exact: true })).toHaveValue('A1:B2');
  });

  await test.step('AutoSum counts the merged balance once and places the total below the range', async () => {
    await sp.clickRibbon('autosum');
    await expectFormula(sp, 'A3', '=SUM(A1:B2)');
    await expectValue(sp, 'A3', { kind: 'number', value: 150 });
    await expectValue(sp, 'B1', { kind: 'blank' });
  });

  await test.step('insert an average using the function arguments dialog', async () => {
    await sp.goTo('D1');
    await page.locator('[data-ribbon-tab="formulas"]').click();
    await sp.clickRibbon('avg');
    const dialog = page.locator('.fc-fxdialog');
    await expect(dialog).toBeVisible();
    await dialog.locator('.fc-fxdialog__arg-input').first().fill('A2:B2');
    await dialog.getByRole('button', { name: 'Insert', exact: true }).click();
    await expect(dialog).toBeHidden();
    await expectFormula(sp, 'D1', '=AVERAGE(A2:B2)');
    await expectValue(sp, 'D1', { kind: 'number', value: 25 });
  });
  await sp.expectNoConsoleErrors();
}

/** Copy a formatted two-column block, insert the copied cells, and exercise
 * cell, row, and column insertion/deletion on a table-like sheet. */
export async function runTableOrganizationScenario(page: Page): Promise<void> {
  const sp = await start(page);

  await test.step('create and format a table block', async () => {
    await sp.enter('A1', 'Product');
    await sp.enter('B1', '10');
    await sp.enter('A2', 'Widget');
    await sp.enter('B2', '4');
    await sp.enter('D3', 'tail');
    await sp.enter('C3', 'keep');
    await sp.enter('A5', 'row-marker');
    await sp.enter('H1', 'column-marker');

    await sp.goTo('A1:B1');
    await sp.clickRibbon('bold');
    await sp.goTo('B1:B2');
    await chooseCurrency(sp);
  });

  await test.step('copy values and formats to a new block', async () => {
    await sp.goTo('A1:B2');
    await sp.shortcut('c');
    await sp.goTo('D1');
    await sp.shortcut('v');
    await expectValue(sp, 'D1', { kind: 'text', value: 'Product' });
    await expectValue(sp, 'E1', { kind: 'number', value: 10 });
    await expect
      .poll(() => sp.readFormat('D1'), { timeout: 2_000 })
      .toEqual(expect.objectContaining({ bold: true }));
    await expect
      .poll(() => sp.readFormat('E1'), { timeout: 2_000 })
      .toEqual(expect.objectContaining({ numFmt: expect.objectContaining({ kind: 'currency' }) }));
  });

  await test.step('insert the copied cells and shift the existing tail', async () => {
    await sp.goTo('A1:B2');
    await sp.shortcut('c');
    await sp.goTo('D3');
    await sp.insertCopiedCells();
    await expectValue(sp, 'D3', { kind: 'text', value: 'Product' });
    await expectValue(sp, 'E3', { kind: 'number', value: 10 });
    await expectValue(sp, 'F3', { kind: 'text', value: 'tail' });
  });

  await test.step('insert cells downward and restore the marker', async () => {
    await sp.goTo('C3');
    await sp.insertRowsOrColumns('cells');
    await expectValue(sp, 'C3', { kind: 'blank' });
    await expectValue(sp, 'C4', { kind: 'text', value: 'keep' });
  });

  await test.step('insert and delete a table row', async () => {
    await sp.goTo('A5');
    await sp.insertRowsOrColumns('rows');
    await expectValue(sp, 'A6', { kind: 'text', value: 'row-marker' });
    await sp.goTo('A5');
    await sp.deleteRowsOrColumns('rows');
    await expectValue(sp, 'A5', { kind: 'text', value: 'row-marker' });
  });

  await test.step('insert and delete a table column', async () => {
    await sp.goTo('H1');
    await sp.insertRowsOrColumns('cols');
    await expectValue(sp, 'I1', { kind: 'text', value: 'column-marker' });
    await sp.goTo('H1');
    await sp.deleteRowsOrColumns('cols');
    await expectValue(sp, 'H1', { kind: 'text', value: 'column-marker' });
  });
  await test.step('paste the copied block once with Enter', async () => {
    await sp.goTo('A1:B2');
    await sp.shortcut('c');
    await sp.goTo('D6');
    await page.keyboard.press('Enter');
    await expectValue(sp, 'D6', { kind: 'text', value: 'Product' });
    await expectValue(sp, 'E6', { kind: 'number', value: 10 });
    await page.keyboard.press('Enter');
    await expect(page.getByRole('textbox', { name: 'Name box', exact: true })).toHaveValue('E8');
    await expectValue(sp, 'E8', { kind: 'blank' });
  });
  await sp.expectNoConsoleErrors();
}

/** Rename and add sheets, keep their edits isolated, and paste a formatted
 * value across sheets through the same copy/paste UI a user would use. */
export async function runMonthlySheetsScenario(page: Page): Promise<void> {
  const sp = await start(page);

  await test.step('name the monthly sheets', async () => {
    await sp.renameSelectedSheet('Plan');
    await sp.page.locator('.fc-host__sheetbar-add').click();
    await expect(sp.page.locator('.fc-host__sheetbar-tab')).toHaveCount(2);
    await sp.renameSelectedSheet('Actual');
    await expect.poll(() => sp.sheetCount(), { timeout: 2_000 }).toBe(2);
    await expect.poll(() => sp.sheetName(0), { timeout: 2_000 }).toBe('Plan');
    await expect.poll(() => sp.sheetName(1), { timeout: 2_000 }).toBe('Actual');
  });

  await test.step('enter values independently on each sheet', async () => {
    await sp.chooseSheet('Plan');
    await sp.enter('A1', '100');
    await sp.enter('A2', '50');
    await sp.enter('A3', '=SUM(A1:A2)');
    await expectValue(sp, 'A3', { kind: 'number', value: 150 });
    await sp.chooseSheet('Actual');
    await sp.enter('A1', '999');
    await expectValue(sp, 'A1', { kind: 'number', value: 999 }, 1);
    await sp.chooseSheet('Plan');
    await expectValue(sp, 'A1', { kind: 'number', value: 100 }, 0);
    await expectFormula(sp, 'A3', '=SUM(A1:A2)');
  });

  await test.step('copy a formatted plan value to Actual', async () => {
    await sp.chooseSheet('Plan');
    await sp.goTo('A2');
    await chooseCurrency(sp);
    await sp.goTo('A2');
    await sp.shortcut('c');
    await sp.chooseSheet('Actual');
    await sp.goTo('B1');
    await sp.shortcut('v');
    await expectValue(sp, 'B1', { kind: 'number', value: 50 }, 1);
    await expect
      .poll(() => sp.readFormat('B1', 1), { timeout: 2_000 })
      .toEqual(expect.objectContaining({ numFmt: expect.objectContaining({ kind: 'currency' }) }));
    await sp.chooseSheet('Plan');
    await expectValue(sp, 'A1', { kind: 'number', value: 100 }, 0);
  });
  await sp.expectNoConsoleErrors();
}

/** Save through the demo title bar, load the downloaded file into a fresh page,
 * and continue editing after checking values, formulas, formats, merges, and
 * sheet names survived the xlsx round trip. */
export async function runSaveOpenContinueScenario(page: Page): Promise<void> {
  const sp = await start(page);

  await test.step('build a workbook worth saving', async () => {
    await sp.renameSelectedSheet('Stored');
    await sp.page.locator('.fc-host__sheetbar-add').click();
    await expect(sp.page.locator('.fc-host__sheetbar-tab')).toHaveCount(2);
    await sp.renameSelectedSheet('Archive');
    await sp.chooseSheet('Stored');
    await sp.enter('D1', 'Reference');
    await sp.goTo('D1');
    const linkDialog = await sp.openHyperlinkDialog();
    await linkDialog.locator('input[type="url"]').fill('https://example.com/reference');
    await linkDialog.getByRole('button', { name: 'OK', exact: true }).click();
    await sp.enter('A1', '21');
    await sp.enter('A2', '=A1*2');
    await sp.goTo('A1');
    await chooseCurrency(sp);
    await expect
      .poll(() => sp.readFormat('A1'), { timeout: 2_000 })
      .toEqual(expect.objectContaining({ numFmt: expect.objectContaining({ kind: 'currency' }) }));
    await sp.enter('B1', 'Saved heading');
    await sp.goTo('B1');
    await sp.clickRibbon('bold');
    await expect.poll(() => sp.readFormat('B1')).toEqual(expect.objectContaining({ bold: true }));
    await sp.goTo('B1:C1');
    await chooseMergeAction(sp, 'mergeCells');
    await expect
      .poll(() => sp.readMerge('B1'), { timeout: 2_000 })
      .toEqual({ sheet: 0, r0: 0, c0: 1, r1: 0, c1: 2 });
    await expect.poll(() => sp.readFormat('B1')).toEqual(expect.objectContaining({ bold: true }));
  });

  const downloadPath = await test.step('save from the title bar', async () => sp.saveDownload());
  const fresh = await page.context().newPage();
  const reopened = new UserJourneyPage(fresh);
  try {
    await test.step('open the saved file on a fresh page', async () => {
      await reopened.mount({ fixture: 'empty' });
      await reopened.expectNoStub();
      await reopened.openDownload(downloadPath);
      await expect.poll(() => reopened.sheetCount(), { timeout: 10_000 }).toBe(2);
      await expect.poll(() => reopened.sheetName(0), { timeout: 10_000 }).toBe('Stored');
      await expect.poll(() => reopened.sheetName(1), { timeout: 10_000 }).toBe('Archive');
    });

    await test.step('verify persisted values, formulas, formats, and merge', async () => {
      await reopened.chooseSheet('Stored');
      await expectValue(reopened, 'A1', { kind: 'number', value: 21 });
      await expectValue(reopened, 'A2', { kind: 'number', value: 42 });
      await expectFormula(reopened, 'A2', '=A1*2');
      await expectValue(reopened, 'B1', { kind: 'text', value: 'Saved heading' });
      await expect
        .poll(() => reopened.readFormat('B1'))
        .toEqual(expect.objectContaining({ bold: true }));
      await expect
        .poll(() => reopened.readFormat('D1'))
        .toEqual(expect.objectContaining({ hyperlink: 'https://example.com/reference' }));
      await expect
        .poll(() => reopened.readFormat('A1'), { timeout: 5_000 })
        .toEqual(
          expect.objectContaining({ numFmt: expect.objectContaining({ kind: 'currency' }) }),
        );
      await expect
        .poll(() => reopened.readMerge('B1'), { timeout: 5_000 })
        .toEqual({ sheet: 0, r0: 0, c0: 1, r1: 0, c1: 2 });
    });

    await test.step('continue editing after reopening', async () => {
      await reopened.enter('A1', '30');
      await expectValue(reopened, 'A1', { kind: 'number', value: 30 });
      await expectValue(reopened, 'A2', { kind: 'number', value: 60 });
      await expectFormula(reopened, 'A2', '=A1*2');
    });
    await test.step('open another workbook without carrying over the previous formats', async () => {
      const blankPage = await page.context().newPage();
      const blank = new UserJourneyPage(blankPage);
      try {
        await blank.mount({ fixture: 'empty' });
        await blank.expectNoStub();
        const blankPath = await blank.saveDownload();
        await reopened.openDownload(blankPath);
        await expect.poll(() => reopened.sheetCount()).toBe(1);
        await expect.poll(() => reopened.sheetName(0)).toBe('Sheet1');
        await expectValue(reopened, 'A1', { kind: 'blank' });
        await expect.poll(() => reopened.readFormat('A1')).toBeNull();
        await expect.poll(() => reopened.readFormat('B1')).toBeNull();
        await expect.poll(() => reopened.readFormat('D1')).toBeNull();
        await expect.poll(() => reopened.readMerge('B1')).toBeNull();
        await blank.expectNoConsoleErrors();
      } finally {
        await blankPage.close();
      }
    });
    await reopened.expectNoConsoleErrors();
    await sp.expectNoConsoleErrors();
  } finally {
    await fresh.close();
  }
}
