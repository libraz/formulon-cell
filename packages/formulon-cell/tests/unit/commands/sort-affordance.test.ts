import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import {
  builtInCommandIntent,
  canExecuteBuiltIn,
} from '../../../src/commands/built-in-command-policy.js';
import { interactionControllerFor } from '../../../src/commands/interaction-controller.js';
import type { InteractionPolicy } from '../../../src/commands/interaction-policy.js';
import { setProtectedSheet } from '../../../src/commands/protection.js';
import { en } from '../../../src/i18n/strings/en.js';
import { attachFilterDropdown } from '../../../src/interact/filter-dropdown.js';
import { createEditingDropdownDefaults } from '../../../src/mount/dynamic-dropdown-defaults/editing.js';
import { mutators } from '../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/mount.js';

const SORT_IDS = [
  'sortAsc',
  'sortDesc',
  'sortData',
  'sortFilterHome',
  'mac.data.sortAsc',
  'mac.data.sortDesc',
  'mac.data.sortCustom',
] as const;

describe('sort availability on a protected sheet', () => {
  let sheet: MountedStubSheet;

  beforeEach(async () => {
    sheet = await mountStubSheet({ locale: 'en' });
    const { store } = sheet.instance;
    sheet.workbook.setText({ sheet: 0, row: 0, col: 0 }, 'Qty');
    for (const [row, value] of [
      [1, 3],
      [2, 1],
      [3, 2],
    ] as const) {
      sheet.workbook.setNumber({ sheet: 0, row, col: 0 }, value);
    }
    sheet.workbook.recalc();
    mutators.replaceCells(store, sheet.workbook.cells(0));
    mutators.setActive(store, { sheet: 0, row: 1, col: 0 });
    mutators.setRange(store, { sheet: 0, r0: 1, c0: 0, r1: 1, c1: 0 });
  });
  afterEach(() => sheet.dispose());

  const column = (): unknown[] =>
    [1, 2, 3].map((row) => {
      const value = sheet.workbook.getValue({ sheet: 0, row, col: 0 });
      return value.kind === 'number' ? value.value : value.kind;
    });
  const unlockData = (): void => {
    for (const row of [1, 2, 3])
      mutators.setCellFormat(sheet.instance.store, { sheet: 0, row, col: 0 }, { locked: false });
  };
  /** What the ribbon (canExecuteBuiltIn) and the context menu (the built-in intent) show. */
  const availability = () => {
    const { store } = sheet.instance;
    const controller = interactionControllerFor(store);
    return SORT_IDS.map((id) => {
      const intent = builtInCommandIntent(store, id, 'contextMenu');
      return {
        id,
        ribbon: canExecuteBuiltIn(store, id, 'ribbon').allowed,
        contextMenu: intent ? controller?.canExecute(intent).allowed : undefined,
      };
    });
  };
  const all = (value: boolean) => SORT_IDS.map((id) => ({ id, ribbon: value, contextMenu: value }));

  for (const [label, policy] of [
    ['without a policy', undefined],
    ['under an allow-all policy', { defaultOperation: 'allow' }],
  ] as const satisfies readonly (readonly [string, InteractionPolicy | undefined])[]) {
    it(`disables every sort command while the Sort option is off ${label}`, () => {
      interactionControllerFor(sheet.instance.store)?.setPolicy(policy);
      unlockData();
      setProtectedSheet(sheet.instance.store, 0, true);
      expect(availability()).toEqual(all(false));
    });

    it(`enables the sort commands with the Sort option on, locked or not ${label}`, () => {
      interactionControllerFor(sheet.instance.store)?.setPolicy(policy);
      setProtectedSheet(sheet.instance.store, 0, true, { permissions: { sort: true } });
      expect(availability()).toEqual(all(true));
      unlockData();
      expect(availability()).toEqual(all(true));
    });
  }

  it('disables the filter dropdown sort items while the Sort option is off', () => {
    const { store, history } = sheet.instance;
    const host = document.createElement('div');
    host.className = 'fc-host';
    host.dataset.fcPlatform = 'mac';
    document.body.appendChild(host);
    const handle = attachFilterDropdown({
      store,
      host,
      strings: en,
      getWb: () => sheet.workbook,
      history,
    });
    const sortButtons = (): (boolean | undefined)[] => {
      handle.open({ sheet: 0, r0: 0, c0: 0, r1: 3, c1: 0 }, 0, { x: 10, y: 20, h: 24 });
      const disabled = ['asc', 'desc'].map(
        (direction) =>
          document.querySelector<HTMLButtonElement>(`.fc-filter-dropdown__sort-${direction}`)
            ?.disabled,
      );
      handle.close();
      return disabled;
    };
    try {
      setProtectedSheet(store, 0, true);
      expect(sortButtons()).toEqual([true, true]);
      setProtectedSheet(store, 0, true, { permissions: { sort: true } });
      expect(sortButtons()).toEqual([false, false]);
    } finally {
      handle.detach();
      host.remove();
    }
  });

  it('keeps a policy that denies sort in charge of availability', () => {
    interactionControllerFor(sheet.instance.store)?.setPolicy({
      defaultOperation: 'allow',
      operations: { sort: false },
    });
    expect(availability()).toEqual(all(false));
  });

  it('does nothing at all when an enabled sort meets locked cells', async () => {
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => undefined);
    try {
      setProtectedSheet(sheet.instance.store, 0, true, { permissions: { sort: true } });
      sheet.instance.history.clear();
      const before = document.body.innerHTML;
      await createEditingDropdownDefaults(sheet.instance).applySortMenuAction('asc');
      await new Promise((resolve) => setTimeout(resolve, 0));
      expect(column()).toEqual([3, 1, 2]);
      expect(sheet.instance.history.canUndo()).toBe(false);
      expect(warn).not.toHaveBeenCalled();
      expect(document.body.innerHTML).toBe(before);
    } finally {
      warn.mockRestore();
    }
  });

  it('sorts unlocked cells with the Sort option on', async () => {
    unlockData();
    setProtectedSheet(sheet.instance.store, 0, true, { permissions: { sort: true } });
    await createEditingDropdownDefaults(sheet.instance).applySortMenuAction('asc');
    expect(column()).toEqual([1, 2, 3]);
  });
});
