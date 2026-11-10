import { describe, expect, it } from 'vitest';
import {
  BUILT_IN_COMMAND_IDS,
  OTHER_BUILT_IN_COMMAND_IDS,
} from '../../../src/commands/built-in-command-ids.js';
import { BUILTIN_ITEM_IDS } from '../../../src/interact/context-menu-spec.js';
import { MAC_RIBBON_MENU_ITEMS } from '../../../src/toolbar/ribbon/mac/menu-manifest.js';
import {
  buildRibbonModel,
  CONTEXTUAL_RIBBON_TABS,
  RIBBON_TABS,
} from '../../../src/toolbar/ribbon-model.js';

const LAYOUT_KINDS = new Set(['break', 'label']);

const modelIds = (): Set<string> => {
  const ids = new Set<string>();
  for (const profile of ['default', 'excel365Mac'] as const) {
    const models = [
      ...buildRibbonModel('en', { tabs: [...RIBBON_TABS, ...CONTEXTUAL_RIBBON_TABS], profile }),
      ...buildRibbonModel('en', { profile, contextualTabs: CONTEXTUAL_RIBBON_TABS }),
    ];
    for (const tab of models) {
      for (const group of tab.groups) {
        for (const command of group.commands) {
          if (!LAYOUT_KINDS.has(command.kind ?? 'button') && !command.disabled) ids.add(command.id);
        }
      }
    }
  }
  return ids;
};

const menuIds = (): Set<string> => {
  const ids = new Set<string>();
  for (const items of Object.values(MAC_RIBBON_MENU_ITEMS)) {
    for (const item of items) {
      if (!item.heading && !item.disabled && !item.id.startsWith('mac.function.')) ids.add(item.id);
    }
  }
  return ids;
};

describe('built-in command ids', () => {
  it('are the enabled ribbon commands, Mac menu leaves, context-menu built-ins and the others', () => {
    const ids = new Set([
      ...modelIds(),
      ...menuIds(),
      ...BUILTIN_ITEM_IDS,
      ...OTHER_BUILT_IN_COMMAND_IDS,
    ]);
    expect([...BUILT_IN_COMMAND_IDS]).toEqual([...ids].sort());
  });

  it('keep the others out of the generated sources', () => {
    const generated = new Set([...modelIds(), ...menuIds(), ...BUILTIN_ITEM_IDS]);
    expect(OTHER_BUILT_IN_COMMAND_IDS.filter((id) => generated.has(id))).toEqual([]);
  });
});
