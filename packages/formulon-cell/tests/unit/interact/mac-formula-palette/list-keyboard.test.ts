import { afterEach, describe, expect, it } from 'vitest';
import { type MountedStubSheet, mountStubSheet } from '../../../test-utils/mount.js';
import { tabOutOf } from '../../../test-utils/tab-order.js';

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

describe('the Formula Builder function list', () => {
  const openPane = async (): Promise<HTMLElement> => {
    sheet = await mountStubSheet({ locale: 'en', ui: { platform: 'mac' } });
    key(sheet.instance.host, { key: 'F3', shiftKey: true });
    const pane = document.querySelector<HTMLElement>('.fc-mac-formula-palette:not([hidden])');
    if (!pane) throw new Error('The Formula Builder did not open.');
    return pane;
  };

  it('is one tab stop, so Tab and Shift+Tab leave the pane', async () => {
    const pane = await openPane();
    const lists = pane.querySelectorAll('[role="listbox"]');
    const rows = pane.querySelectorAll('[role="option"]');
    expect(rows.length).toBeGreaterThan(30);
    for (const list of lists) {
      const stops = Array.from(list.querySelectorAll<HTMLElement>('[role="option"]')).filter(
        (row) => row.tabIndex >= 0,
      );
      expect(stops.length).toBeLessThanOrEqual(1);
    }
    expect(tabOutOf(pane, false, 20)).not.toBeNull();
    expect(tabOutOf(pane, true, 20)).not.toBeNull();
  });

  it('moves between rows with the arrow keys, Home and End', async () => {
    const pane = await openPane();
    const rows = Array.from(pane.querySelectorAll<HTMLElement>('[role="listbox"]'))
      .map((list) => Array.from(list.querySelectorAll<HTMLElement>('[role="option"]')))
      .reduce((a, b) => (b.length > a.length ? b : a), []);
    const [first, second] = rows;
    const last = rows.at(-1);
    if (!first || !second || !last) throw new Error('The list has too few rows.');
    first.focus();
    expect(key(first, { key: 'ArrowDown' }).defaultPrevented).toBe(true);
    expect(document.activeElement).toBe(second);
    expect(second.tabIndex).toBe(0);
    expect(first.tabIndex).toBe(-1);
    key(second, { key: 'ArrowUp' });
    expect(document.activeElement).toBe(first);
    key(first, { key: 'End' });
    expect(document.activeElement).toBe(last);
    key(last, { key: 'Home' });
    expect(document.activeElement).toBe(first);
  });
});
