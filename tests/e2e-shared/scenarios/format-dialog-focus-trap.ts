import { expect, type Page } from '@playwright/test';

import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

/** B7 — format dialog focus stays inside the active tab panel.
 *
 *  Each tab in the format dialog renders as a sibling `[role=tabpanel]` with
 *  `hidden=true` on inactive ones. Focus must not escape through Tab key
 *  navigation into a hidden tab's inputs — `hidden` is the documented escape
 *  hatch for the focus tree, and breaking it (e.g. switching to `aria-hidden`
 *  alone) regresses screen-reader and keyboard users. */
export async function runFormatDialogFocusTrapScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount();
  await sp.expectNoStub();

  await sp.focusHost();
  await sp.shortcut('1');

  const dialog = page.locator('[class="fc-fmtdlg"]');
  await expect(dialog).toBeVisible({ timeout: 2_000 });

  // Snapshot which panel is currently active and confirm it's the only one
  // not [hidden].
  const panelSummary = await page.evaluate(() => {
    const panels = Array.from(document.querySelectorAll<HTMLElement>('.fc-fmtdlg__panel-tab'));
    return panels.map((p) => ({ id: p.dataset.fcTab ?? '', hidden: p.hidden }));
  });
  const visible = panelSummary.filter((p) => !p.hidden);
  expect(visible).toHaveLength(1);
  const visibleTabId = visible[0]?.id;

  // Verify ALL inputs/buttons inside the hidden panels have a zero layout
  // box — meaning Tab navigation skips them entirely.
  const focusableInHidden = await page.evaluate(() => {
    const hiddenPanels = Array.from(
      document.querySelectorAll<HTMLElement>('.fc-fmtdlg__panel-tab[hidden]'),
    );
    let count = 0;
    for (const panel of hiddenPanels) {
      const candidates = panel.querySelectorAll<HTMLElement>(
        'button, input, select, textarea, [tabindex]:not([tabindex="-1"])',
      );
      for (const c of candidates) {
        if (c instanceof HTMLButtonElement && c.disabled) continue;
        if (c instanceof HTMLInputElement && c.disabled) continue;
        if (c.offsetParent !== null) count += 1;
      }
    }
    return count;
  });
  expect(
    focusableInHidden,
    'no input in a hidden format-dialog tab panel should be in the layout tree',
  ).toBe(0);

  // Switch to a different tab and re-check — the previously-active panel
  // becomes hidden and its descendants leave the focus tree.
  // Most demos default to "number"; switching to "font" is a reliable jump.
  const targetTab = visibleTabId === 'font' ? 'number' : 'font';
  const targetBtn = page.locator(`[role="tab"][data-fc-tab="${targetTab}"]`);
  if ((await targetBtn.count()) > 0) {
    await targetBtn.click();
    await page.waitForTimeout(100);

    const afterSwitch = await page.evaluate(() => {
      return Array.from(document.querySelectorAll<HTMLElement>('.fc-fmtdlg__panel-tab')).map(
        (p) => ({ id: p.dataset.fcTab ?? '', hidden: p.hidden }),
      );
    });
    const visibleAfter = afterSwitch.filter((p) => !p.hidden);
    expect(visibleAfter).toHaveLength(1);
    expect(visibleAfter[0]?.id).toBe(targetTab);

    // Re-verify the hidden side stays focus-inert.
    const focusableInHiddenAfter = await page.evaluate(() => {
      const hiddenPanels = Array.from(
        document.querySelectorAll<HTMLElement>('.fc-fmtdlg__panel-tab[hidden]'),
      );
      let count = 0;
      for (const panel of hiddenPanels) {
        const candidates = panel.querySelectorAll<HTMLElement>(
          'button, input, select, textarea, [tabindex]:not([tabindex="-1"])',
        );
        for (const c of candidates) {
          if (c instanceof HTMLButtonElement && c.disabled) continue;
          if (c instanceof HTMLInputElement && c.disabled) continue;
          if (c.offsetParent !== null) count += 1;
        }
      }
      return count;
    });
    expect(focusableInHiddenAfter).toBe(0);
  }

  // Close cleanly.
  await page.keyboard.press('Escape');
  await expect(dialog).toBeHidden({ timeout: 2_000 });
}
