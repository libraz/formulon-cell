import { expect, type Page } from '@playwright/test';
import type { E2eProbes } from '../../../apps/demo-shared/e2e-probes.js';
import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
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
    instance.workbook.batch(() => {
      const sheet = instance.workbook.sheet(0);
      sheet.setFormula(3, 3, '=B2+SUM(B3:C3)+LOG10(100)');
      sheet.setNumber(1, 1, 2);
      sheet.setNumber(2, 1, 3);
      sheet.setNumber(2, 2, 4);
    });
    instance.history.clear();
  });
  await sheet.goTo('D4');
  await page.evaluate(() => {
    const w = window as unknown as { __fcInst: SpreadsheetInstance; __fcProbes: E2eProbes };
    const { store } = w.__fcProbes['mount/internals'].spreadsheetInternals(w.__fcInst);
    store.setState((state) => ({ ...state, ui: { ...state.ui, r1c1: true } }));
  });
  const input = page.locator(
    surface === 'inline' ? '.fc-host__editor' : '.fc-host__formulabar-input',
  );
  if (surface === 'inline') await page.keyboard.press('F2');
  else await input.focus();
  const suffix = '+SUM(R[-1]C[-2]:R[-1]C[-1])+LOG10(100)';
  await expect(input).toHaveValue(`=R[-2]C[-2]${suffix}`);
  const refs = () =>
    page.evaluate(() => {
      const w = window as unknown as { __fcInst: SpreadsheetInstance; __fcProbes: E2eProbes };
      return w.__fcProbes['mount/internals'].spreadsheetInternals(w.__fcInst).store.getState().ui
        .editorRefs;
    });
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
    const w = window as unknown as { __fcInst: SpreadsheetInstance; __fcProbes: E2eProbes };
    w.__fcInst.workbook.batch(() => {
      const sheet = w.__fcInst.workbook.sheet(0);
      sheet.setNumber(rc5.row, rc5.col, 9);
      sheet.setNumber(3, 4, 17);
      sheet.setNumber(3, 5, 23);
    });
    w.__fcInst.history.clear();
    const { store } = w.__fcProbes['mount/internals'].spreadsheetInternals(w.__fcInst);
    store.setState((state) => ({ ...state, ui: { ...state.ui, r1c1: false } }));
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
    const w = window as unknown as { __fcInst: SpreadsheetInstance; __fcProbes: E2eProbes };
    const { store } = w.__fcProbes['mount/internals'].spreadsheetInternals(w.__fcInst);
    store.setState((state) => ({ ...state, ui: { ...state.ui, r1c1: true } }));
  });
  if (surface === 'inline') await page.keyboard.press('F2');
  else await input.focus();
  await input.fill('=RC5');
  await page.keyboard.press('Enter');
  await expect.poll(() => sheet.readFormula('D4')).toBe('=$E4');
  expect(await sheet.readValue('D4')).toEqual({ kind: 'number', value: 17 });

  await page.evaluate(() => {
    const w = window as unknown as { __fcInst: SpreadsheetInstance; __fcProbes: E2eProbes };
    const { store } = w.__fcProbes['mount/internals'].spreadsheetInternals(w.__fcInst);
    store.setState((state) => ({ ...state, ui: { ...state.ui, r1c1: false } }));
  });
  await sheet.goTo('D4:E4');
  await page.evaluate(() => {
    const w = window as unknown as { __fcInst: SpreadsheetInstance; __fcProbes: E2eProbes };
    const { store } = w.__fcProbes['mount/internals'].spreadsheetInternals(w.__fcInst);
    store.setState((state) => ({ ...state, ui: { ...state.ui, r1c1: true } }));
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
    instance.workbook.batch(() => instance.workbook.sheet(0).setFormula(0, 0, '=1+2'));
    instance.history.clear();
    instance.applyChanges([{ sheet: 0, at: { row: 0, col: 1 }, input: '6' }], {
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
      const w = window as unknown as { __fcInst: SpreadsheetInstance; __fcProbes: E2eProbes };
      // The tagged cell value has no facade equivalent; read it from the engine handle.
      const handle = w.__fcProbes['api/workbook'].workbookHandleOf(w.__fcInst.workbook);
      return {
        formula: handle.cellFormula({ sheet: 0, row: 0, col: 0 }),
        marker: handle.getValue({ sheet: 0, row: 0, col: 1 }),
        undo: w.__fcInst.history.canUndo,
        redo: w.__fcInst.history.canRedo,
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
    await page.evaluate(() => {
      const { history } = (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst;
      const available = history.canUndo;
      history.undo();
      return available;
    }),
  ).toBe(true);
  expect(await sheet.readValue('B1')).toEqual({ kind: 'blank' });
  expect(await sheet.readFormula('A1')).toBe('=1+2');
  await sheet.expectNoConsoleErrors();
}
