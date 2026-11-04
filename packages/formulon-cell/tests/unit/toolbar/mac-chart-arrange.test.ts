import { describe, expect, it } from 'vitest';
import { arrangeStackPosition } from '../../../src/commands/session-illustration.js';
import { createI18nController } from '../../../src/i18n/controller.js';
import type { SpreadsheetInstance } from '../../../src/mount/types.js';
import type { SessionChart } from '../../../src/store/store.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';
import { projectMacRibbonState } from '../../../src/toolbar/ribbon/mac/state-projection.js';

describe('Mac chart Format tab arrange buttons', () => {
  it('disables Bring Forward / Send Backward when there is nothing to reorder past', () => {
    const instance = {
      store: createSpreadsheetStore(),
      i18n: createI18nController({ locale: 'ja' }),
      workbook: {
        calcMode: () => 'auto',
        getExternalLinks: () => [],
        getPivotTables: () => [],
        definedNames: function* () {},
      },
      isSheetProtected: () => false,
    } as unknown as SpreadsheetInstance;
    const host = document.createElement('div');
    const button = (id: string) => {
      const el = document.createElement('button');
      el.dataset.ribbonCommand = id;
      host.appendChild(el);
      return el;
    };
    const forward = button('objectBringForward');
    const backward = button('objectSendBackward');
    const setCharts = (ids: string[], selected: string) =>
      instance.store.setState((s) => ({
        ...s,
        charts: {
          ...s.charts,
          charts: ids.map((id) => ({ id, sheet: 0 }) as unknown as SessionChart),
        },
        ui: { ...s.ui, selectedObject: { id: selected, name: selected } },
      }));
    const disabled = () => {
      projectMacRibbonState(host, instance);
      return [forward.disabled, backward.disabled];
    };

    setCharts(['only'], 'only');
    expect(disabled()).toEqual([true, true]);
    setCharts(['back', 'front'], 'back');
    expect(disabled()).toEqual([false, true]);
    setCharts(['back', 'front'], 'front');
    expect(disabled()).toEqual([true, false]);
  });

  it('shares one front-most / back-most rule with the arrange menu', () => {
    const stack = [{ id: 'a' }, { id: 'b' }, { id: 'c' }];
    expect(arrangeStackPosition(stack, 'a')).toEqual({
      hasTarget: true,
      atFront: false,
      atBack: true,
    });
    expect(arrangeStackPosition(stack, 'b')).toEqual({
      hasTarget: true,
      atFront: false,
      atBack: false,
    });
    expect(arrangeStackPosition(stack, 'c')).toEqual({
      hasTarget: true,
      atFront: true,
      atBack: false,
    });
    expect(arrangeStackPosition(stack, null)).toEqual({
      hasTarget: false,
      atFront: true,
      atBack: true,
    });
  });
});
