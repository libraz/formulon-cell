import { expect, type Page } from '@playwright/test';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

async function openEditor(page: Page): Promise<UserJourneyPage> {
  const sp = new UserJourneyPage(page);
  await sp.mount();
  await sp.expectNoStub();
  await sp.goTo('A1');
  await page.keyboard.press('Control+F3');
  await page.locator('.fc-namedlg').getByRole('button', { name: 'New...', exact: true }).click();
  await expect(page.locator('.fc-namedlg-editor')).toBeVisible();
  return sp;
}

export async function runSelectHomeEndAndDeparture(page: Page): Promise<void> {
  const sp = await openEditor(page);
  const editor = page.locator('.fc-namedlg-editor');
  const combo = editor.getByRole('combobox', { name: 'Scope', exact: true });
  const native = editor.locator('select');
  await combo.focus();
  await page.keyboard.press('End');
  await page.keyboard.press('Enter');
  await expect(native).toHaveValue('0');
  await expect(combo).toContainText('Sheet1');
  await page.keyboard.press('Home');
  await page.keyboard.press('Enter');
  await expect(native).toHaveValue('-1');
  // WebKit can blur the trigger with a null relatedTarget before an option click.
  await page.keyboard.press('Enter');
  await editor.getByRole('option', { name: 'Sheet1', exact: true }).click();
  await expect(native).toHaveValue('0');
  await expect(combo).toHaveAttribute('aria-expanded', 'false');
  await combo.focus();
  await page.keyboard.press('Home');
  await page.keyboard.press('Enter');
  await expect(native).toHaveValue('-1');
  await page.keyboard.press('ArrowDown');
  await expect(combo).toHaveAttribute('aria-expanded', 'true');
  await page.keyboard.press('Tab');
  await expect(combo).toHaveAttribute('aria-expanded', 'false');
  await expect(editor.getByRole('textbox', { name: 'Comment', exact: true })).toBeFocused();
  await expect(native).toHaveValue('-1');
  await editor.getByRole('textbox', { name: 'Name', exact: true }).fill('ParityScope');
  await editor.locator('.fc-namedlg-editor__input').last().fill('=Sheet1!$A$1');
  await editor.getByRole('button', { name: 'OK', exact: true }).click();
  await expect(editor).toBeHidden();
  await page.locator('.fc-namedlg').getByRole('button', { name: 'Edit...', exact: true }).click();
  await expect(editor).toBeVisible();
  await expect(combo).toBeDisabled();
  await editor.getByRole('textbox', { name: 'Name', exact: true }).focus();
  await page.keyboard.press('Tab');
  await expect(editor.getByRole('textbox', { name: 'Comment', exact: true })).toBeFocused();
  await sp.expectNoConsoleErrors();
}

export async function runSelectEscapeLayers(page: Page): Promise<void> {
  const sp = await openEditor(page);
  const editor = page.locator('.fc-namedlg-editor');
  const combo = editor.getByRole('combobox', { name: 'Scope', exact: true });
  await combo.focus();
  await page.keyboard.press('ArrowDown');
  await expect(combo).toHaveAttribute('aria-expanded', 'true');
  await page.keyboard.press('Escape');
  await expect(combo).toHaveAttribute('aria-expanded', 'false');
  await expect(editor).toBeVisible();
  await page.keyboard.press('Escape');
  await expect(editor).toBeHidden();
  const manager = page.locator('.fc-namedlg');
  await expect(manager).toBeVisible();
  await expect(manager.getByRole('button', { name: 'New...', exact: true })).toBeFocused();
  await page.keyboard.press('Escape');
  await expect(manager).toBeHidden();
  await sp.expectNoConsoleErrors();
}
