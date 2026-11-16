import { afterEach, describe, expect, it } from 'vitest';
import type { MountOptions } from '../../../src/mount/types.js';
import { mutators } from '../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/mount.js';
import { sheetFocusTarget } from '../../test-utils/sheet-focus.js';

let sheet: MountedStubSheet;
afterEach(() => {
  sheet.dispose();
  document.body.replaceChildren();
});

const tab = (target: Element, shiftKey = false): KeyboardEvent => {
  const event = new KeyboardEvent('keydown', {
    key: 'Tab',
    shiftKey,
    bubbles: true,
    cancelable: true,
  });
  target.dispatchEvent(event);
  return event;
};

const activeOf = () => sheet.instance.store.getState().selection.active;
const place = (row: number, col: number): void =>
  mutators.setActive(sheet.instance.store, { sheet: 0, row, col });

describe.each(['default', 'mac'] as const)(
  'Tab from the formula bar on the %s profile',
  (platform) => {
    const mountBar = async (extra: Partial<MountOptions> = {}) => {
      sheet = await mountStubSheet({ locale: 'en', ui: { platform }, ...extra });
      const input = sheet.host.querySelector<HTMLTextAreaElement>('.fc-host__formulabar-input');
      if (!input) throw new Error('No formula bar.');
      return input;
    };
    const edit = (input: HTMLTextAreaElement, text: string): void => {
      input.focus();
      input.dispatchEvent(new FocusEvent('focus'));
      input.value = text;
      input.dispatchEvent(new Event('input'));
    };

    it('commits an edited bar and moves right, then left with Shift+Tab', async () => {
      const input = await mountBar();
      place(0, 1);
      edit(input, 'one');
      expect(tab(input).defaultPrevented).toBe(true);
      expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({
        kind: 'text',
        value: 'one',
      });
      expect(activeOf()).toMatchObject({ row: 0, col: 2 });
      expect(document.activeElement).toBe(sheetFocusTarget(sheet.host));

      edit(input, 'two');
      expect(tab(input, true).defaultPrevented).toBe(true);
      expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 2 })).toEqual({
        kind: 'text',
        value: 'two',
      });
      expect(activeOf()).toMatchObject({ row: 0, col: 1 });
      expect(document.activeElement).toBe(sheetFocusTarget(sheet.host));
    });

    it('moves the cell and focuses the grid from an unedited editable bar', async () => {
      const input = await mountBar();
      place(2, 2);
      input.focus();
      expect(tab(input).defaultPrevented).toBe(true);
      expect(activeOf()).toMatchObject({ row: 2, col: 3 });
      expect(document.activeElement).toBe(sheetFocusTarget(sheet.host));
      input.focus();
      expect(tab(input, true).defaultPrevented).toBe(true);
      expect(activeOf()).toMatchObject({ row: 2, col: 2 });
    });

    it('leaves Tab alone when the policy forbids editing the cell', async () => {
      const input = await mountBar({ policy: { readOnly: true } });
      place(2, 2);
      input.focus();
      const event = tab(input);
      expect(event.defaultPrevented).toBe(false);
      expect(activeOf()).toMatchObject({ row: 2, col: 2 });
      expect(tab(input, true).defaultPrevented).toBe(false);
      expect(activeOf()).toMatchObject({ row: 2, col: 2 });
    });

    it.each([
      ['stop', true],
      ['leave', false],
    ] as const)("follows tabBoundary '%s' at the viewport edge", async (tabBoundary, prevented) => {
      const input = await mountBar({ viewport: { range: 'A1:B1', tabBoundary } });
      place(0, 1);
      edit(input, 'edge');
      expect(tab(input).defaultPrevented).toBe(prevented);
      expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({
        kind: 'text',
        value: 'edge',
      });
      expect(activeOf()).toMatchObject({ row: 0, col: 1 });
      expect(document.activeElement).toBe(sheetFocusTarget(sheet.host));
    });
  },
);

describe.each(['default', 'mac'] as const)(
  'the default tab boundary on the %s profile',
  (platform) => {
    const mountGrid = (viewport: MountOptions['viewport'], chrome: 'grid-only' | 'full') =>
      mountStubSheet({
        locale: 'en',
        viewport,
        ui: {
          platform,
          ...(chrome === 'grid-only'
            ? { features: { formulaBar: false, sheetTabs: false, statusBar: false } }
            : {}),
        },
      });
    const range = 'A1:B1';

    const gridPresses = (reverse: boolean): KeyboardEvent => {
      place(0, reverse ? 0 : 1);
      const grid = sheetFocusTarget(sheet.host);
      grid.focus();
      return tab(grid, reverse);
    };

    it('has no region besides the grid in the grid-only mount', async () => {
      sheet = await mountGrid({ range }, 'grid-only');
      expect(sheet.host.querySelector('.fc-host__formulabar')).toBeNull();
      expect(sheet.host.querySelector('.fc-host__sheetbar')).toBeNull();
      expect(sheet.host.querySelector('.fc-host__statusbar')).toBeNull();
      expect(sheet.host.querySelector('.fc-host__ribbon')).toBeNull();
    });

    it('leaves at the edge of a grid-only mount when tabBoundary is unset', async () => {
      sheet = await mountGrid({ range }, 'grid-only');
      expect(gridPresses(false).defaultPrevented).toBe(false);
      expect(gridPresses(true).defaultPrevented).toBe(false);
    });

    it('leaves from the in-cell editor at the edge of a grid-only mount', async () => {
      sheet = await mountGrid({ range }, 'grid-only');
      for (const reverse of [false, true]) {
        place(0, reverse ? 0 : 1);
        const grid = sheetFocusTarget(sheet.host);
        grid.focus();
        grid.dispatchEvent(
          new KeyboardEvent('keydown', { key: 'F2', bubbles: true, cancelable: true }),
        );
        const editor = sheet.host.querySelector<HTMLTextAreaElement | HTMLInputElement>(
          '.fc-host__editor',
        );
        if (!editor) throw new Error('No editor.');
        expect(tab(editor, reverse).defaultPrevented).toBe(false);
      }
    });

    it("stops at the edge of a grid-only mount with an explicit 'stop'", async () => {
      sheet = await mountGrid({ range, tabBoundary: 'stop' }, 'grid-only');
      expect(gridPresses(false).defaultPrevented).toBe(true);
      expect(gridPresses(true).defaultPrevented).toBe(true);
    });

    it('stops at the edge when the other regions are shown', async () => {
      sheet = await mountGrid({ range }, 'full');
      expect(gridPresses(false).defaultPrevented).toBe(true);
      expect(gridPresses(true).defaultPrevented).toBe(true);
    });

    it("leaves at the edge with an explicit 'leave' beside other regions", async () => {
      sheet = await mountGrid({ range, tabBoundary: 'leave' }, 'full');
      expect(gridPresses(false).defaultPrevented).toBe(false);
    });
  },
);
