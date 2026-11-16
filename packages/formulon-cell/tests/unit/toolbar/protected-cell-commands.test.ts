import { afterEach, describe, expect, it } from 'vitest';
import { setProtectedSheet } from '../../../src/commands/protection.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { dictionaries } from '../../../src/i18n/strings.js';
import { attachRibbonInteractionPolicy } from '../../../src/mount/toolbar-interaction-policy.js';
import { mutators } from '../../../src/store/store.js';
import { projectRibbonActiveState } from '../../../src/toolbar/ribbon-active-state.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/mount.js';

// Bold and Clear Formats follow the one built-in decision on every profile: a locked active cell
// of a protected sheet greys them unless the sheet allows formatting cells.
const textFor = (locale: keyof typeof dictionaries) => ({
  background: dictionaries[locale].ribbon.background,
  sheetBackgroundClear: dictionaries[locale].ribbonMenu.sheetBackgroundClear,
});

describe('Bold and Clear on a protected Mac sheet', () => {
  let sheet: MountedStubSheet | undefined;

  afterEach(() => {
    sheet?.dispose();
    sheet = undefined;
  });

  it('are greyed by the interaction policy alone, in the UI language', async () => {
    sheet = await mountStubSheet({
      workbook: await WorkbookHandle.createDefault(),
      locale: 'ja',
      ui: { platform: 'mac' },
    });
    const instance = sheet.instance;
    const host = document.createElement('div');
    host.innerHTML =
      '<button data-ribbon-command="bold"></button><button data-ribbon-command="clearFormat"></button>';
    const buttons = [...host.querySelectorAll<HTMLButtonElement>('button')];
    const projection = attachRibbonInteractionPolicy(host, {
      getInstance: () => instance,
      closeDynamicDropdowns: () => undefined,
    });
    const project = (): void => {
      projectRibbonActiveState(host, instance, textFor('ja'));
      projection.project();
    };
    try {
      setProtectedSheet(instance.store, 0, true);
      project();
      for (const button of buttons) {
        expect(button.getAttribute('aria-disabled')).toBe('true');
        expect(button.dataset.ribbonPolicyDisabled).toMatch(/^変更しようとしているセルやグラフ/);
      }

      // Format cells opens Bold; clearing stays all or nothing over locked cells.
      setProtectedSheet(instance.store, 0, true, { permissions: { formatCells: true } });
      project();
      expect(buttons.map((button) => button.getAttribute('aria-disabled'))).toEqual([
        'false',
        'true',
      ]);

      setProtectedSheet(instance.store, 0, true);
      project();
      for (const button of buttons) expect(button.getAttribute('aria-disabled')).toBe('true');
      mutators.setCellFormat(instance.store, instance.store.getState().selection.active, {
        locked: false,
      });
      project();
      // An unlocked cell clears, but formatting it still needs Format cells, as the write does.
      expect(buttons.map((button) => button.getAttribute('aria-disabled'))).toEqual([
        'true',
        'false',
      ]);
    } finally {
      projection.detach();
    }
  });
});
