import { expect, type Page } from '@playwright/test';
import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/mount/types.js';
import { gridCellCenter } from '../pages/SpreadsheetPage.js';
import { a1Address, UserJourneyPage } from '../pages/UserJourneyPage.js';

export type FormulaTransactionSurface = 'inline' | 'bar';
export type FormulaTransactionPlatform = 'default' | 'mac';

export async function runR1C1ReferenceTransactions(
  page: Page,
  surface: FormulaTransactionSurface,
  platform: FormulaTransactionPlatform,
): Promise<void> {
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ locale: 'en', platform });
  await sheet.expectNoStub();
  await page.evaluate(() => {
    const instance = (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst;
    instance.applyChanges([
      { addr: { sheet: 0, row: 3, col: 3 }, input: '=B2+SUM(B3:C3)+LOG10(100)' },
      { addr: { sheet: 0, row: 1, col: 1 }, input: '2' },
      { addr: { sheet: 0, row: 2, col: 1 }, input: '3' },
      { addr: { sheet: 0, row: 2, col: 2 }, input: '4' },
    ]);
  });
  await sheet.goTo('D4');
  await page.evaluate(() => {
    const instance = (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst;
    instance.store.setState((state) => ({ ...state, ui: { ...state.ui, r1c1: true } }));
  });
  const input = page.locator(
    surface === 'inline' ? '.fc-host__editor' : '.fc-host__formulabar-input',
  );
  if (surface === 'inline') await page.keyboard.press('F2');
  else await input.focus();
  const suffix = '+SUM(R[-1]C[-2]:R[-1]C[-1])+LOG10(100)';
  await expect(input).toHaveValue(`=R[-2]C[-2]${suffix}`);
  const refs = () =>
    page.evaluate(
      () =>
        (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst.store.getState().ui
          .editorRefs,
    );
  const expectedRefs = [
    { r0: 1, c0: 1, r1: 1, c1: 1, colorIndex: 0 },
    { r0: 2, c0: 1, r1: 2, c1: 2, colorIndex: 1 },
  ];
  await expect.poll(refs).toEqual(expectedRefs);
  await input.evaluate((element) => (element as HTMLTextAreaElement).setSelectionRange(11, 11));
  const undo = platform === 'mac' ? 'Meta+z' : 'Control+z';
  const redo = platform === 'mac' ? 'Meta+Shift+z' : 'Control+y';
  for (const [index, reference] of ['R2C2', 'R2C[-2]', 'R[-2]C2', 'R[-2]C[-2]'].entries()) {
    await page.keyboard.press('F4');
    await expect(input).toHaveValue(`=${reference}${suffix}`);
    await expect.poll(refs).toEqual(expectedRefs);
    if (index === 0) {
      await page.keyboard.press(undo);
      await expect(input).toHaveValue(`=R[-2]C[-2]${suffix}`);
      await page.keyboard.press(redo);
      await expect(input).toHaveValue(`=${reference}${suffix}`);
    }
  }
  await page.keyboard.press(platform === 'mac' ? 'Meta+t' : 'F4');
  await expect(input).toHaveValue(`=R2C2${suffix}`);
  await page.keyboard.press('Enter');
  await expect.poll(() => sheet.readFormula('D4')).toBe('=$B$2+SUM(B3:C3)+LOG10(100)');
  expect(await sheet.readValue('D4')).toEqual({ kind: 'number', value: 11 });

  await page.evaluate((rc5) => {
    const instance = (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst;
    instance.applyChanges([
      { addr: rc5, input: '9' },
      { addr: { sheet: 0, row: 3, col: 4 }, input: '17' },
      { addr: { sheet: 0, row: 3, col: 5 }, input: '23' },
    ]);
    instance.store.setState((state) => ({ ...state, ui: { ...state.ui, r1c1: false } }));
  }, a1Address('RC5'));
  await sheet.goTo('D4');
  if (surface === 'inline') await page.keyboard.press('F2');
  else await input.focus();
  await input.fill('=RC5');
  await page.keyboard.press('Enter');
  await expect.poll(() => sheet.readFormula('D4')).toBe('=RC5');
  expect(await sheet.readValue('D4')).toEqual({ kind: 'number', value: 9 });

  await sheet.goTo('D4');
  await page.evaluate(() => {
    const instance = (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst;
    instance.store.setState((state) => ({ ...state, ui: { ...state.ui, r1c1: true } }));
  });
  if (surface === 'inline') await page.keyboard.press('F2');
  else await input.focus();
  await input.fill('=RC5');
  await page.keyboard.press('Enter');
  await expect.poll(() => sheet.readFormula('D4')).toBe('=$E4');
  expect(await sheet.readValue('D4')).toEqual({ kind: 'number', value: 17 });

  await page.evaluate(() => {
    const instance = (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst;
    instance.store.setState((state) => ({ ...state, ui: { ...state.ui, r1c1: false } }));
  });
  await sheet.goTo('D4:E4');
  await page.evaluate(() => {
    const instance = (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst;
    instance.store.setState((state) => ({ ...state, ui: { ...state.ui, r1c1: true } }));
  });
  if (surface === 'inline') await page.keyboard.press('F2');
  else await input.focus();
  await input.fill('=RC6');
  await page.keyboard.press(platform === 'mac' ? 'Meta+Enter' : 'Control+Enter');
  for (const ref of ['D4', 'E4']) {
    await expect.poll(() => sheet.readFormula(ref)).toBe('=$F4');
    expect(await sheet.readValue(ref)).toEqual({ kind: 'number', value: 23 });
  }
  await sheet.expectNoConsoleErrors();
}

export async function runFormulaDraftTransactions(
  page: Page,
  surface: FormulaTransactionSurface,
  platform: FormulaTransactionPlatform,
): Promise<void> {
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ locale: 'en', platform });
  await sheet.expectNoStub();
  await page.evaluate(() => {
    const instance = (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst;
    instance.applyChanges([{ addr: { sheet: 0, row: 0, col: 0 }, input: '=1+2' }]);
    instance.applyChanges([{ addr: { sheet: 0, row: 0, col: 1 }, input: '6' }], {
      history: 'record',
    });
  });
  const input = page.locator(
    surface === 'inline' ? '.fc-host__editor' : '.fc-host__formulabar-input',
  );
  const begin = async () => {
    await sheet.goTo('A1');
    if (surface === 'inline') await page.keyboard.press('F2');
    else await input.focus();
  };
  const undo = platform === 'mac' ? 'Meta+z' : 'Control+z';
  const redo = platform === 'mac' ? 'Meta+Shift+z' : 'Control+y';
  const workbookState = () =>
    page.evaluate(() => {
      const instance = (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst;
      return {
        formula: instance.workbook.cellFormula({ sheet: 0, row: 0, col: 0 }),
        marker: instance.workbook.getValue({ sheet: 0, row: 0, col: 1 }),
        undo: instance.history.canUndo(),
        redo: instance.history.canRedo(),
      };
    });
  const baseline = await workbookState();
  expect(baseline).toEqual({
    formula: '=1+2',
    marker: { kind: 'number', value: 6 },
    undo: true,
    redo: false,
  });

  await begin();
  await input.fill('=SU');
  await page
    .locator('.fc-autocomplete__item')
    .filter({
      has: page.locator('.fc-autocomplete__name', { hasText: /^SUM$/ }),
    })
    .click();
  await page.keyboard.press('Tab');
  await expect(input).toHaveValue('=SUM(');
  await page.keyboard.press(undo);
  await expect(input).toHaveValue('=SU');
  expect(await input.evaluate((el) => (el as HTMLTextAreaElement).selectionStart)).toBe(3);
  await page.keyboard.press(redo);
  await expect(input).toHaveValue('=SUM(');
  expect(await workbookState()).toEqual(baseline);
  await page.keyboard.press('Escape');

  await begin();
  await input.fill('=SUM(');
  const from = await gridCellCenter(page, 3, 2);
  const to = await gridCellCenter(page, 4, 3);
  await page.mouse.move(from.x, from.y);
  await page.mouse.down();
  await page.mouse.move(to.x, to.y, { steps: 5 });
  await page.mouse.up();
  await expect(input).toHaveValue('=SUM(C4:D5');
  await page.keyboard.press(undo);
  await expect(input).toHaveValue('=SUM(');
  expect(await input.evaluate((el) => (el as HTMLTextAreaElement).selectionStart)).toBe(5);
  await page.keyboard.press(redo);
  await expect(input).toHaveValue('=SUM(C4:D5');
  expect(await workbookState()).toEqual(baseline);
  await page.keyboard.press('Escape');

  await begin();
  await input.fill('=C4+2');
  await input.evaluate((el) => (el as HTMLTextAreaElement).setSelectionRange(3, 3));
  await page.keyboard.press('F4');
  await expect(input).toHaveValue('=$C$4+2');
  await page.keyboard.press(undo);
  await expect(input).toHaveValue('=C4+2');
  await page.keyboard.press(redo);
  await expect(input).toHaveValue('=$C$4+2');
  expect(await workbookState()).toEqual(baseline);
  await page.keyboard.press('Escape');

  await begin();
  await input.fill('=1+2');
  await input.evaluate((el) => (el as HTMLTextAreaElement).setSelectionRange(1, 4));
  await page.keyboard.press('F9');
  await expect(input).toHaveValue('=3');
  await page.keyboard.press(undo);
  await expect(input).toHaveValue('=1+2');
  expect(
    await input.evaluate((el) => {
      const textarea = el as HTMLTextAreaElement;
      return [textarea.selectionStart, textarea.selectionEnd];
    }),
  ).toEqual([1, 4]);
  await page.keyboard.press(redo);
  await expect(input).toHaveValue('=3');
  expect(await workbookState()).toEqual(baseline);
  await page.keyboard.press('Escape');

  expect(
    await page.evaluate(() =>
      (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst.undo(),
    ),
  ).toBe(true);
  expect(await sheet.readValue('B1')).toEqual({ kind: 'blank' });
  expect(await sheet.readFormula('A1')).toBe('=1+2');
  await sheet.expectNoConsoleErrors();
}
