import { expect, type Page } from '@playwright/test';
import type { E2eProbes } from '../../../apps/demo-shared/e2e-probes.js';
import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

const url = 'https://example.com/formulon-native-check';

async function linkState(page: Page) {
  return page.evaluate(() => {
    const w = window as unknown as { __fcInst: SpreadsheetInstance; __fcProbes: E2eProbes };
    const { layoutForView } = w.__fcProbes['render/geometry'];
    const { store } = w.__fcProbes['mount/internals'].spreadsheetInternals(w.__fcInst);
    // Resolved fonts, the display font and the theme have no facade member.
    const wb = w.__fcProbes['api/workbook'].workbookHandleOf(w.__fcInst.workbook);
    const style = wb.getEffectiveStyle(0, 0, 0);
    const xf = style && wb.getCellXf(style.xfIndex);
    const font = xf && wb.getResolvedFontRecord(xf.fontIndex);
    const normal = wb.workbookDisplayFont;
    const hlink = wb.getTheme()?.colors[10];
    if (!font || !normal || hlink === undefined) throw new Error('Workbook font/theme unavailable');
    const format = store.getState().format.formats.get('0:0:0');
    const layout = layoutForView(store.getState());
    return {
      rowHeight: layout.rowHeights.get(0) ?? layout.defaultRowHeight,
      value: wb.getValue({ sheet: 0, row: 0, col: 0 }),
      target: format?.hyperlink ?? null,
      nativeTarget:
        wb.getHyperlinks(0).find((link) => link.row === 0 && link.col === 0)?.target ?? null,
      points: font.sizePt,
      family: font.name,
      normalPoints: normal.size,
      normalFamily: normal.name,
      underline: font.underline,
      color: (font.color?.argb ?? 0) >>> 0,
      hlink: hlink >>> 0,
    };
  });
}

export async function runTypedUrlAutoformat(
  page: Page,
  platform: 'default' | 'mac',
  capture?: (name: string) => Promise<void>,
): Promise<void> {
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ locale: 'en', platform });
  await sheet.expectNoStub();
  await page.evaluate(() => {
    const w = window as unknown as { __fcInst: SpreadsheetInstance; __fcProbes: E2eProbes };
    const { store } = w.__fcProbes['mount/internals'].spreadsheetInternals(w.__fcInst);
    const handle = w.__fcProbes['api/workbook'].workbookHandleOf(w.__fcInst.workbook);
    const normal = handle.workbookDisplayFont;
    if (!normal) throw new Error('Default font unavailable');
    const { mutators } = w.__fcProbes['store/store'];
    mutators.setColWidth(store, 0, 480);
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      {
        fontSize: (normal.size + 13) * (4 / 3),
        color: '#aa0000',
        underline: false,
      },
    );
  });
  const before = await linkState(page);
  // CSS-pixel conversion can leave a few floating-point bits in native point values.
  expect(before.points).toBeCloseTo(before.normalPoints + 13, 10);
  await sheet.enter('A1', url);
  await expect.poll(async () => (await linkState(page)).target).toBe(url);
  const linked = await linkState(page);
  expect(linked.value).toEqual({ kind: 'text', value: url });
  expect(linked.nativeTarget).toBe(url);
  expect(linked.points).toBe(linked.normalPoints);
  expect(linked.family).toBe(linked.normalFamily);
  expect(linked.underline).toBe(1);
  expect(linked.color).toBe(linked.hlink);
  expect(linked.rowHeight).toBeLessThan(before.rowHeight);
  await capture?.('linked');

  await sheet.shortcut('z');
  await expect.poll(async () => (await linkState(page)).target).toBeNull();
  const undoneLink = await linkState(page);
  expect(undoneLink.value).toEqual({ kind: 'text', value: url });
  expect(undoneLink.points).toBe(before.points);
  expect(undoneLink.color).toBe(before.color);
  expect(undoneLink.underline).toBe(0);
  expect(undoneLink.rowHeight).toBe(before.rowHeight);
  await capture?.('undo-conversion');
  await sheet.shortcut('z');
  await expect.poll(() => sheet.readValue('A1')).toEqual({ kind: 'blank' });
  await sheet.shortcut('Shift+z');
  await expect.poll(() => sheet.readValue('A1')).toEqual({ kind: 'text', value: url });
  expect((await linkState(page)).target).toBeNull();
  await sheet.shortcut('Shift+z');
  await expect.poll(async () => (await linkState(page)).nativeTarget).toBe(url);
  expect((await linkState(page)).rowHeight).toBe(linked.rowHeight);
  await capture?.('redo-conversion');

  await sheet.enter('B1', `'${url}`);
  expect(await sheet.readValue('B1')).toEqual({ kind: 'text', value: url });
  expect((await sheet.readFormat('B1'))?.hyperlink).toBeUndefined();
  await sheet.goTo('C1');
  const formulaBar = page.locator('.fc-host__formulabar-input').first();
  await formulaBar.fill(url);
  await formulaBar.press('Enter');
  await expect.poll(() => sheet.readValue('C1')).toEqual({ kind: 'text', value: url });
  expect((await sheet.readFormat('C1'))?.hyperlink).toBeUndefined();

  await sheet.goTo('A1');
  await sheet.shortcut('k');
  const dialog = page.locator('.fc-hldlg');
  await expect(dialog).toBeVisible();
  // The Mac dialog's URL field is a plain text input labelled "URL".
  const urlField = dialog
    .locator('input[type="url"]')
    .or(dialog.getByRole('textbox', { name: 'URL', exact: true }));
  await expect(urlField).toHaveValue(url);
  await dialog.getByRole('button', { name: 'Remove link', exact: true }).click();
  await expect(dialog).toBeHidden();
  await expect.poll(async () => (await linkState(page)).nativeTarget).toBeNull();
  expect(await sheet.readValue('A1')).toEqual({ kind: 'text', value: url });
  await sheet.shortcut('z');
  await expect.poll(async () => (await linkState(page)).nativeTarget).toBe(url);
  await sheet.expectNoConsoleErrors();
}
