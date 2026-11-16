import {
  DIALOG_ESCAPE_FIXTURES,
  expectDialogEscapeClassification,
  runDialogEscapeScenario,
} from '../../../tests/e2e-shared/scenarios/dialog-escape.js';
import { test } from '../../../tests/e2e-shared/test.js';

for (const platform of ['default', 'mac'] as const) {
  test(`D06 (vue-demo, ${platform}): every dialog id is a modal, a floating panel, a docked pane or absent`, () => {
    expectDialogEscapeClassification(platform);
  });

  for (const fixture of DIALOG_ESCAPE_FIXTURES) {
    test(`D06 (vue-demo, ${platform}, ${fixture}): Escape closes dialogs and panels back to their opener and leaves panes open`, async ({
      page,
    }) => {
      test.setTimeout(180_000);
      await runDialogEscapeScenario(page, platform, fixture);
    });
  }
}
