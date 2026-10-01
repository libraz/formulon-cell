import { expect, type Page } from '@playwright/test';
import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/mount/types.js';
import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

type DemoWindow = Window & { __fcInst?: SpreadsheetInstance };

export async function runRestrictedEmbeddingScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount();
  await sp.expectNoStub();
  await page.waitForFunction(() => !!(window as DemoWindow).__fcInst);
  await page.evaluate(() => {
    const instance = (window as DemoWindow).__fcInst;
    if (!instance) throw new Error('Spreadsheet instance is unavailable');
    instance.setUi({ profile: 'embedded', features: { shortcuts: true, clipboard: true } });
    instance.setViewportOptions({ range: { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 2 } });
    instance.setPolicy({
      operations: { valueEdit: true, clear: true, paste: true, fill: true },
      editable: [{ sheet: 0, r0: 0, c0: 0, r1: 3, c1: 0 }],
    });
    instance.setContextMenu({
      mode: 'builtIn',
      items: ['copy', 'paste', 'clear', 'rowInsertAbove'],
    });
    instance.applyChanges([
      { addr: { sheet: 0, row: 0, col: 0 }, input: '1' },
      { addr: { sheet: 0, row: 0, col: 1 }, input: '=A1*2' },
    ]);
    const dialog = document.createElement('dialog');
    dialog.id = 'embedded-host-dialog';
    dialog.style.cssText = 'width:900px;height:500px';
    document.body.append(dialog);
    dialog.append(instance.host);
    dialog.showModal();
    instance.setOverlayOptions({ root: dialog });
    instance.host.focus();
  });
  // Actual keyboard input passes through the restricted editor and real WASM.
  await page.keyboard.type('7');
  await page.keyboard.press('Enter');
  await expect
    .poll(() =>
      page.evaluate(() =>
        (window as DemoWindow).__fcInst?.workbook.getValue({ sheet: 0, row: 0, col: 1 }),
      ),
    )
    .toEqual({ kind: 'number', value: 14 });

  await page.keyboard.press('Shift+F10');
  const menu = page.locator('#embedded-host-dialog .fc-ctxmenu:not(.fc-ctxmenu__sub)');
  await expect(menu).toBeVisible();
  await expect(menu.getByText('Insert', { exact: false })).toHaveCount(0);
  await page.keyboard.press('Escape');
  await page.evaluate(() => {
    const dialog = document.getElementById('embedded-host-dialog');
    if (dialog) dialog.style.transform = 'translate(25px, 10px) scale(0.9)';
  });
  const canvas = await page.locator('#embedded-host-dialog canvas').first().boundingBox();
  if (!canvas) throw new Error('Embedded canvas is unavailable');
  const menuPoint = { x: canvas.x + 80, y: canvas.y + 35 };
  await page.mouse.click(menuPoint.x, menuPoint.y, { button: 'right' });
  await expect(menu).toBeVisible();
  const menuBox = await menu.boundingBox();
  if (!menuBox) throw new Error('Embedded menu is unavailable');
  expect(menuBox.x).toBeCloseTo(menuPoint.x, 0);
  expect(menuBox.y).toBeCloseTo(menuPoint.y, 0);
  await page.keyboard.press('Escape');
  await page.evaluate(() => {
    const instance = (window as DemoWindow).__fcInst;
    if (!instance) throw new Error('Spreadsheet instance is unavailable');
    instance.setPolicy({ readOnly: true });
    instance.host.focus();
  });
  await page.keyboard.type('99');
  await page.keyboard.press('Enter');
  await expect
    .poll(() =>
      page.evaluate(() =>
        (window as DemoWindow).__fcInst?.workbook.getValue({ sheet: 0, row: 1, col: 0 }),
      ),
    )
    .toEqual({ kind: 'blank' });
  // The default portal follows a real fullscreen boundary after leaving the modal.
  await page.evaluate(() => {
    const instance = (window as DemoWindow).__fcInst;
    if (!instance) throw new Error('Spreadsheet instance is unavailable');
    const dialog = document.getElementById('embedded-host-dialog') as HTMLDialogElement;
    dialog.close();
    const fullscreen = document.createElement('div');
    fullscreen.id = 'embedded-fullscreen';
    fullscreen.style.cssText = 'width:100%;height:100%';
    const button = document.createElement('button');
    button.textContent = 'Enter embedded fullscreen';
    button.addEventListener('click', () => {
      void fullscreen.requestFullscreen();
    });
    fullscreen.append(button, instance.host);
    document.body.append(fullscreen);
    instance.setOverlayOptions(undefined);
  });
  await page.getByRole('button', { name: 'Enter embedded fullscreen' }).click();
  await expect
    .poll(() => page.evaluate(() => document.fullscreenElement?.id))
    .toBe('embedded-fullscreen');
  await page.evaluate(() => (window as DemoWindow).__fcInst?.host.focus());
  await page.keyboard.press('Shift+F10');
  await expect(
    page.locator('#embedded-fullscreen .fc-ctxmenu:not(.fc-ctxmenu__sub)'),
  ).toBeVisible();
  await page.keyboard.press('Escape');
  await page.evaluate(() =>
    document.fullscreenElement ? document.exitFullscreen() : Promise.resolve(),
  );
}
