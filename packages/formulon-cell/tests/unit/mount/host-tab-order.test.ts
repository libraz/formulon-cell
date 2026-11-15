import { afterEach, describe, expect, it } from 'vitest';
import { mutators } from '../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/mount.js';
import { tabOutOf, tabStops } from '../../test-utils/tab-order.js';

const key = (target: Element, init: KeyboardEventInit): KeyboardEvent => {
  const event = new KeyboardEvent('keydown', { bubbles: true, cancelable: true, ...init });
  target.dispatchEvent(event);
  return event;
};

let sheet: MountedStubSheet;
afterEach(() => {
  sheet.dispose();
  document.body.replaceChildren();
});

const mountOn = (platform: 'default' | 'mac') =>
  mountStubSheet({ locale: 'en', toolbar: true, ui: { platform } });

describe('the sheet bar', () => {
  it.each(['default', 'mac'] as const)(
    'lets Tab and Shift+Tab leave it on the %s profile',
    async (platform) => {
      sheet = await mountOn(platform);
      const bar = sheet.instance.host.querySelector<HTMLElement>('.fc-host__sheetbar');
      if (!bar) throw new Error('No sheet bar.');
      expect(tabOutOf(bar, false, 10)).not.toBeNull();
      expect(tabOutOf(bar, true, 10)).not.toBeNull();
    },
  );
});

const KEYS = [
  { key: 'Delete' },
  { key: 'x' },
  { key: 'ArrowDown' },
  { key: 'Enter' },
  { key: 'PageDown' },
];

// A key pressed on a bar or pane control in the host belongs to that control, not to the grid.
describe('controls outside the grid', () => {
  it.each(['default', 'mac'] as const)(
    'leave the cells and the active cell alone on the %s profile',
    async (platform) => {
      sheet = await mountOn(platform);
      const store = sheet.instance.store;
      const host = sheet.instance.host;
      const controls = tabStops().filter(
        (el) =>
          host.contains(el) &&
          el !== host &&
          !el.closest('.fc-host__grid') &&
          !el.matches('input, textarea, select'),
      );
      expect(controls.some((el) => el.closest('.fc-host__ribbon'))).toBe(true);
      const reached: string[] = [];
      const reset = (): void => {
        sheet.workbook.setNumber({ sheet: 0, row: 1, col: 1 }, 7);
        mutators.replaceCells(store, sheet.workbook.cells(0));
        mutators.setActive(store, { sheet: 0, row: 1, col: 1 });
      };
      reset();
      for (const control of controls) {
        for (const init of KEYS) {
          if (!control.isConnected) break;
          control.focus();
          key(control, init);
          const state = store.getState();
          if (
            state.ui.editor.kind !== 'idle' ||
            state.selection.active.row !== 1 ||
            state.selection.active.col !== 1 ||
            sheet.workbook.getValue({ sheet: 0, row: 1, col: 1 }).kind !== 'number'
          ) {
            reached.push(`${control.className}: ${init.key}`);
            key(document.activeElement ?? host, { key: 'Escape' });
            reset();
          }
        }
      }
      expect(reached).toEqual([]);
    },
  );
});
