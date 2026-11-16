import { describe, expect, it } from 'vitest';
import { dictionaries } from '../../../src/i18n/strings.js';
import { buildRibbonModel } from '../../../src/toolbar/ribbon-model.js';

const labels = (tabs: ReturnType<typeof buildRibbonModel>): string[] =>
  tabs.flatMap((tab) => [
    tab.label,
    ...tab.groups.flatMap((group) => [group.title, ...group.commands.map((c) => c.label)]),
  ]);

describe('the ribbon locale comes from the caller', () => {
  it('builds the Mac profile for a cloned zh dictionary in zh', () => {
    const clone = structuredClone(dictionaries.zh);
    const expected = buildRibbonModel('zh', { profile: 'excel365Mac' });
    const actual = buildRibbonModel('zh', { profile: 'excel365Mac', strings: clone });
    expect(labels(actual)).toEqual(labels(expected));
  });
});
