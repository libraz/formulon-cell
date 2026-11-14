import { runContextualGroupTitlesScenario } from '../../../tests/e2e-shared/scenarios/contextual-group-titles.js';
import { test } from '../../../tests/e2e-shared/test.js';

for (const [platform, locale] of [
  ['mac', 'en'],
  ['mac', 'ja'],
  ['mac', 'de'],
  ['default', 'en'],
  ['default', 'de'],
] as const) {
  test(`Contextual tabs title their groups and place their commands (${platform} chrome, ${locale})`, async ({
    page,
  }) => {
    await runContextualGroupTitlesScenario(page, platform, locale);
  });
}
