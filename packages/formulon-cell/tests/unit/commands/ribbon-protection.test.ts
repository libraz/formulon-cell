import { describe, expect, it } from 'vitest';

import { allowedEditRangesForSheet } from '../../../src/commands/protection.js';
import { executeRibbonProtectionAction } from '../../../src/commands/ribbon-protection.js';
import { defaultStrings } from '../../../src/i18n/strings.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

describe('executeRibbonProtectionAction', () => {
  it('adds the active selection as an allowed edit range and returns a host report', () => {
    const store = createSpreadsheetStore();
    store.setState((state) => ({
      ...state,
      selection: {
        ...state.selection,
        range: { sheet: 0, r0: 1, c0: 1, r1: 3, c1: 2 },
      },
    }));

    const report = executeRibbonProtectionAction({
      store,
      action: 'allow-edit-range',
      strings: defaultStrings.ribbonMenu,
    });

    expect(allowedEditRangesForSheet(store.getState(), 0)).toEqual([
      expect.objectContaining({
        title: 'B2:C4',
        range: { sheet: 0, r0: 1, c0: 1, r1: 3, c1: 2 },
      }),
    ]);
    expect(report).toEqual({
      title: defaultStrings.ribbonMenu.allowEditRangesDialogTitle,
      items: [
        {
          severity: 'info',
          label: defaultStrings.ribbonMenu.allowEditRangesCommand,
          detail: defaultStrings.ribbonMenu.allowedEditRangeAddedStatus.replace('{range}', 'B2:C4'),
        },
      ],
    });
  });

  it('clears allowed edit ranges only on the active sheet and returns a host report', () => {
    const store = createSpreadsheetStore();
    store.setState((state) => ({
      ...state,
      data: { ...state.data, sheetIndex: 1 },
      selection: { ...state.selection, range: { sheet: 1, r0: 0, c0: 0, r1: 0, c1: 0 } },
    }));
    executeRibbonProtectionAction({
      store,
      action: 'allow-edit-range',
      strings: defaultStrings.ribbonMenu,
    });
    store.setState((state) => ({
      ...state,
      data: { ...state.data, sheetIndex: 0 },
      selection: { ...state.selection, range: { sheet: 0, r0: 2, c0: 0, r1: 2, c1: 0 } },
    }));
    executeRibbonProtectionAction({
      store,
      action: 'allow-edit-range',
      strings: defaultStrings.ribbonMenu,
    });
    expect(allowedEditRangesForSheet(store.getState(), 0)).toHaveLength(1);
    expect(allowedEditRangesForSheet(store.getState(), 1)).toHaveLength(1);

    const report = executeRibbonProtectionAction({
      store,
      action: 'clear-allowed-edit-ranges',
      strings: defaultStrings.ribbonMenu,
    });

    expect(allowedEditRangesForSheet(store.getState(), 0)).toEqual([]);
    expect(allowedEditRangesForSheet(store.getState(), 1)).toHaveLength(1);
    expect(report).toEqual({
      title: defaultStrings.ribbonMenu.allowEditRangesDialogTitle,
      items: [
        {
          severity: 'info',
          label: defaultStrings.ribbonMenu.allowEditRangesClearCommand,
          detail: defaultStrings.ribbonMenu.allowedEditRangesClearedStatus,
        },
      ],
    });
  });
});
