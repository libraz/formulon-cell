import { expect, type Page, test } from '@playwright/test';

async function mount(page: Page): Promise<void> {
  await page.goto('/?fixture=selection');
  await page.waitForSelector('.fc-host', { state: 'attached', timeout: 30_000 });
  await page.waitForFunction(
    () => {
      const host = document.querySelector('.fc-host') as HTMLElement | null;
      const state = host?.dataset.fcEngineState;
      return state === 'ready' || state === 'ready-stub';
    },
    { timeout: 30_000 },
  );
}

test('S01: status summary menu supports checkbox menu keys and Escape focus return', async ({
  page,
}) => {
  await mount(page);

  const status = page.locator('#status-metric');
  await expect(status).toContainText('Sum');
  await status.evaluate((el) => (el as HTMLElement).focus());
  await expect(status).toBeFocused();
  await status.click({ button: 'right', force: true });

  const menu = page.getByRole('menu', { name: 'Selection summary' });
  await expect(menu).toBeVisible();
  const sum = menu.getByRole('menuitemcheckbox', { name: /SUM/ });
  await expect(sum).toBeFocused();
  await expect(sum).toHaveAttribute('aria-checked', 'true');

  await page.keyboard.press('ArrowDown');
  await expect(menu.getByRole('menuitemcheckbox', { name: /AVG/ })).toBeFocused();
  await page.keyboard.press('Home');
  await expect(sum).toBeFocused();
  await page.keyboard.press('Space');
  await expect(sum).toHaveAttribute('aria-checked', 'false');
  await expect(menu).toBeVisible();

  await page.keyboard.press('Escape');
  await expect(menu).toHaveCount(0);
  await expect(status).toBeFocused();
});
