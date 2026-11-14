import { expect, type Page } from '@playwright/test';

import { DIALOG_LAYOUT_PREFIX, installDialogLayoutCheck } from '../pages/dialog-layout.js';

const FIXTURE_PATH = '/__dialog-layout-fixture';

// A list row whose first cell ellipsises a long name, as the Name Manager's does, beside a dialog
// whose two captions really do overlap.
const FIXTURE = `<style>
  .row { display: grid; grid-template-columns: 120px 120px 120px; font: 12px sans-serif; }
  .row > span { min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  .stack { position: relative; height: 40px; font: 12px sans-serif; }
  .stack > span { position: absolute; left: 0; top: 0; }
  .stack > .popup { position: absolute; left: 0; top: 20px; display: grid; width: 160px; background: #fff; }
  .below { display: block; width: 160px; }
  .stack > .left { position: absolute; left: 0; top: 0; width: 120px; }
  .stack > .right { position: absolute; left: 80px; top: 0; width: 120px; }
</style>
<div role="dialog" aria-label="Clipped">
  <div class="row">
    <span>Revenue_${'Column_'.repeat(24)}</span><span>=Sheet1!$A$1</span><span>Workbook</span>
  </div>
</div>
<div role="dialog" aria-label="Overlapping">
  <div class="stack"><span>First caption</span><span>Second caption</span></div>
</div>
<div role="dialog" aria-label="Open list">
  <div class="stack">
    <button role="combobox" aria-haspopup="listbox" aria-expanded="true">Kind</button>
    <div role="listbox" class="popup"><button role="option">Whole number</button><button role="option">List</button></div>
  </div>
  <input class="below" aria-label="Value">
</div>
<div role="dialog" aria-label="Closed selects">
  <div class="stack">
    <button role="combobox" aria-haspopup="listbox" aria-expanded="false" class="left">Operator</button>
    <button role="combobox" aria-haspopup="listbox" aria-expanded="false" class="right">Type</button>
    <div role="listbox" hidden><button role="option">Percent</button></div>
  </div>
</div>`;

/** The dialog layout check passes a text run its own element clips and an open combobox list
 *  over the controls below it, and still reports two captions drawn on top of each other and two
 *  closed selects that overlap. */
export async function runDialogLayoutCheckScenario(page: Page): Promise<void> {
  const errors: string[] = [];
  page.on('console', (msg) => {
    if (msg.type() === 'error' && msg.text().startsWith(DIALOG_LAYOUT_PREFIX))
      errors.push(msg.text());
  });
  await installDialogLayoutCheck(page);
  await page.route(`**${FIXTURE_PATH}`, (route) =>
    route.fulfill({ contentType: 'text/html', body: '<!doctype html><title>fixture</title>' }),
  );
  await page.goto(FIXTURE_PATH);
  // The check runs on DOM changes, so the dialogs are inserted after load as the app's are.
  await page.evaluate((html) => document.body.insertAdjacentHTML('beforeend', html), FIXTURE);

  // One pass checks every visible dialog, so the control's report means the clipped row was checked.
  await expect.poll(() => errors.some((e) => e.includes(' Overlapping: '))).toBe(true);
  await expect.poll(() => errors.some((e) => e.includes(' Closed selects: '))).toBe(true);
  expect(errors.filter((e) => e.includes(' Clipped: '))).toEqual([]);
  expect(errors.filter((e) => e.includes(' Open list: '))).toEqual([]);
}
