import { expect, type Page } from '@playwright/test';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

export async function runNestedNameDialogKeyboard(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount();
  await sp.expectNoStub();
  await sp.goTo('A1');
  await page.keyboard.press('Control+F3');
  const manager = page.locator('.fc-namedlg');
  const editor = page.locator('.fc-namedlg-editor');
  await expect(manager).toBeVisible();
  const newButton = manager.getByRole('button', { name: 'New...', exact: true });
  await newButton.click();
  await expect(editor).toBeVisible();
  const name = editor.getByRole('textbox', { name: 'Name', exact: true });
  const comment = editor.getByRole('textbox', { name: 'Comment', exact: true });
  const scope = editor.getByRole('combobox', { name: 'Scope', exact: true });
  const formula = editor.getByRole('textbox', { name: 'Refers to', exact: true });
  await expect(name).toBeFocused();
  await name.fill('Revenue');
  await page.keyboard.press('Tab');
  await expect(scope).toBeFocused();
  await page.keyboard.press('Tab');
  await expect(comment).toBeFocused();
  await page.keyboard.press('Tab');
  await expect(formula).toBeFocused();
  await editor.getByRole('button', { name: 'Cancel', exact: true }).focus();
  await page.keyboard.press('Tab');
  await expect(name).toBeFocused();
  await page.keyboard.press('Shift+Tab');
  await expect(editor.getByRole('button', { name: 'Cancel', exact: true })).toBeFocused();
  await page.keyboard.press('Escape');
  await expect(editor).toBeHidden();
  await expect(manager).toBeVisible();
  await expect(newButton).toBeFocused();
  await newButton.click();
  await name.fill('Revenue');
  await formula.fill('=Sheet1!$A$1');
  await editor.getByRole('button', { name: 'OK', exact: true }).click();
  await expect(editor).toBeHidden();
  await expect(manager).toBeVisible();
  await expect(manager.locator('.fc-namedlg__row-name')).toHaveText('Revenue');
  await manager.getByRole('button', { name: 'Delete', exact: true }).click();
  const confirm = page.locator('.fc-namedlg-confirm');
  await expect(confirm).toBeVisible();
  await page.keyboard.press('Escape');
  await expect(confirm).toBeHidden();
  await expect(manager).toBeVisible();
  await expect(manager.locator('.fc-namedlg__row-name')).toHaveText('Revenue');
  await expect(manager.getByRole('button', { name: 'Delete', exact: true })).toBeFocused();
  await page.keyboard.press('Escape');
  await expect(manager).toBeHidden();
  await sp.expectNoConsoleErrors();
}

export async function runNameDeleteFocus(page: Page): Promise<void> {
  const name = `Revenue_${'Column_'.repeat(24)}`;
  const sp = new UserJourneyPage(page);
  await sp.mount();
  await sp.expectNoStub();
  await sp.goTo('A1');
  await page.keyboard.press('Control+F3');
  const manager = page.locator('.fc-namedlg');
  const editor = page.locator('.fc-namedlg-editor');
  await expect(manager).toBeVisible();
  const newButton = manager.getByRole('button', { name: 'New...', exact: true });
  await newButton.click();
  await expect(editor).toBeVisible();
  await editor.getByRole('textbox', { name: 'Name', exact: true }).fill(name);
  await editor.getByRole('textbox', { name: 'Refers to', exact: true }).fill('=Sheet1!$A$1');
  await editor.getByRole('button', { name: 'OK', exact: true }).click();
  await expect(editor).toBeHidden();
  await expect(manager.locator('.fc-namedlg__row-name')).toHaveText(name);

  const confirm = page.locator('.fc-namedlg-confirm');
  const deleteAction = manager.getByRole('button', { name: 'Delete', exact: true });
  const row = manager.locator('.fc-namedlg__item');
  await deleteAction.click();
  await expect(confirm).toBeVisible();
  const confirmationSpacing = await confirm
    .locator('.fc-namedlg-confirm__panel')
    .evaluate((panel) => {
      const body = panel.querySelector<HTMLElement>('.fc-namedlg-confirm__body');
      const buttons = panel.querySelector<HTMLElement>('.fc-namedlg-confirm__buttons');
      const message = panel.querySelector<HTMLElement>('.fc-namedlg-confirm__message');
      if (!body || !buttons || !message) throw new Error('Missing confirmation content');
      const panelStyle = getComputedStyle(panel);
      const bodyStyle = getComputedStyle(body);
      return {
        trailing: panel.getBoundingClientRect().bottom - buttons.getBoundingClientRect().bottom,
        messageRight: message.getBoundingClientRect().right,
        contentRight:
          body.getBoundingClientRect().right - Number.parseFloat(bodyStyle.paddingRight),
        padding:
          Number.parseFloat(bodyStyle.paddingBottom) +
          Number.parseFloat(panelStyle.borderBottomWidth),
      };
    });
  expect.soft(confirmationSpacing.trailing).toBeCloseTo(confirmationSpacing.padding, 0);
  expect.soft(confirmationSpacing.messageRight).toBeCloseTo(confirmationSpacing.contentRight, 0);
  await page.keyboard.press('Escape');
  await expect(confirm).toBeHidden();
  await expect(deleteAction).toBeFocused();

  await row.focus();
  await page.keyboard.press('Delete');
  await expect(confirm).toBeVisible();
  await page.keyboard.press('Escape');
  await expect(confirm).toBeHidden();
  await expect(row).toBeFocused();

  await page.keyboard.press('Backspace');
  await expect(confirm).toBeVisible();
  await confirm.getByRole('button', { name: 'Cancel', exact: true }).click();
  await expect(confirm).toBeHidden();
  await expect(row).toBeFocused();

  await deleteAction.click();
  await expect(confirm).toBeVisible();
  await confirm.getByRole('button', { name: 'OK', exact: true }).click();
  await expect(confirm).toBeHidden();
  await expect(manager.locator('.fc-namedlg__item')).toHaveCount(0);
  await expect(newButton).toBeFocused();

  await page.keyboard.press('Escape');
  await expect(manager).toBeHidden();
  await expect(page.locator('.fc-host')).toBeFocused();

  await page.keyboard.press('Control+z');
  await page.keyboard.press('Control+F3');
  await expect(manager).toBeVisible();
  await expect(manager.locator('.fc-namedlg__row-name')).toHaveText(name);
  await page.keyboard.press('Escape');
  await expect(manager).toBeHidden();
  await sp.expectNoConsoleErrors();
}
