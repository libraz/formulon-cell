import { describe, expect, it } from 'vitest';
import { CELL_COMMAND_IDS, COMMAND_IDS } from '../../../src/api/spreadsheet/command-ids.js';
import { isBuiltInCommandId } from '../../../src/commands/built-in-command-ids.js';
import {
  buildRibbonModel,
  CONTEXTUAL_RIBBON_TABS,
  RIBBON_TABS,
} from '../../../src/toolbar/ribbon-model.js';

const VALUE_KINDS = new Set(['select', 'color', 'input', 'break', 'label']);

describe('command ids', () => {
  it('are exactly the ribbon commands that take no value and the cell commands', () => {
    const ids = new Set<string>(CELL_COMMAND_IDS);
    for (const profile of ['default', 'excel365Mac'] as const) {
      const models = [
        ...buildRibbonModel('en', { tabs: [...RIBBON_TABS, ...CONTEXTUAL_RIBBON_TABS], profile }),
        ...buildRibbonModel('en', { profile, contextualTabs: CONTEXTUAL_RIBBON_TABS }),
      ];
      for (const tab of models) {
        for (const group of tab.groups) {
          for (const command of group.commands) {
            if (!VALUE_KINDS.has(command.kind ?? 'button')) ids.add(command.id);
          }
        }
      }
    }
    expect([...COMMAND_IDS]).toEqual([...ids].sort());
  });

  it('name a built-in command for every cell command, so policy and protection decide it', () => {
    for (const id of CELL_COMMAND_IDS) expect(isBuiltInCommandId(id), id).toBe(true);
  });
});
