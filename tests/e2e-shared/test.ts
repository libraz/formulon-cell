import { test as base } from '@playwright/test';

import { DIALOG_LAYOUT_PREFIX, installDialogLayoutCheck } from './pages/dialog-layout.js';

/** Whether a test fails on the dialog layout faults its page logs. Only the layout check's own
 *  fixture, which draws faults on purpose, sets `expected`. */
export type DialogLayoutFaults = 'fail' | 'expected';

/**
 * The `test` every e2e spec imports. It watches every dialog the test's page shows and fails the
 * test if any of them logged a layout fault, so no spec has to remember the check. Other console
 * errors stay with the scenarios that collect them.
 */
export const test = base.extend<{ dialogLayoutFaults: DialogLayoutFaults; dialogLayout: void }>({
  dialogLayoutFaults: ['fail', { option: true }],
  dialogLayout: [
    async ({ page, dialogLayoutFaults }, use) => {
      await installDialogLayoutCheck(page);
      const faults: string[] = [];
      page.on('console', (msg) => {
        if (msg.type() === 'error' && msg.text().startsWith(DIALOG_LAYOUT_PREFIX))
          faults.push(msg.text());
      });
      await use();
      if (dialogLayoutFaults === 'fail') {
        base.expect(faults, 'dialogs should lay out without overlaps').toEqual([]);
      }
    },
    { auto: true },
  ],
});
