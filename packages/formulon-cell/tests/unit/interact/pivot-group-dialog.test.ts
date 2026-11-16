import { PivotDatePeriod } from '@libraz/formulon';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { PivotAxis } from '../../../src/engine/types.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { ja } from '../../../src/i18n/strings/ja.js';
import { attachContextMenu, type ContextMenuHandle } from '../../../src/interact/context-menu.js';
import { disposeOverlayPortal } from '../../../src/interact/overlay-portal.js';
import { createSpreadsheetStore, type SpreadsheetStore } from '../../../src/store/store.js';
import { loadFixture } from '../engine/pivot-excel-parity/fixture.js';
import { item, setSelectionRanges } from './context-menu/fixtures.js';

// Sheet5 shows Date on rows 4-8 (2024-01-15 .. 03-10); its cache holds dates up to 2024-07-22.
describe('the Grouping dialog of a PivotTable date field', () => {
  let host: HTMLElement;
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;
  let history: History;
  let sheet: number;
  let detach: ContextMenuHandle;

  beforeEach(async () => {
    host = document.createElement('div');
    host.tabIndex = -1;
    document.body.appendChild(host);
    wb = await loadFixture('pivot_value_date_filters.xlsx');
    sheet = Array.from({ length: wb.sheetCount }, (_, i) => i).find(
      (i) => wb.sheetName(i) === 'Sheet5',
    ) as number;
    store = createSpreadsheetStore();
    store.setState((s) => ({ ...s, data: { ...s.data, sheetIndex: sheet } }));
    history = new History();
    detach = attachContextMenu({ getLocale: () => 'en-US', host, store, wb, strings: ja, history });
  });

  afterEach(() => {
    detach();
    disposeOverlayPortal(host);
    document.body.innerHTML = '';
    wb.dispose();
  });

  const openOn = (row: number, col: number): void => {
    setSelectionRanges(store, { sheet, r0: row, c0: col, r1: row, c1: col }, []);
    host.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'F10', shiftKey: true, bubbles: true, cancelable: true }),
    );
  };
  const dialog = () => document.querySelector<HTMLElement>('.fc-pivotgroup');
  const inputs = () =>
    Array.from(dialog()?.querySelectorAll<HTMLInputElement>('input') ?? []).filter(
      (el) => el.type !== 'checkbox',
    );
  const checks = () =>
    Array.from(dialog()?.querySelectorAll<HTMLInputElement>('input[type="checkbox"]') ?? []);
  const periods = () => dialog()?.querySelector<HTMLSelectElement>('select');
  const pick = (...values: number[]): void => {
    const select = periods();
    if (!select) throw new Error('no list');
    for (const option of Array.from(select.options)) {
      option.selected = values.includes(Number(option.value));
    }
    select.dispatchEvent(new Event('change', { bubbles: true }));
  };
  const ok = () =>
    dialog()?.querySelector<HTMLButtonElement>('.fc-pivotgroup__footer button:last-child');
  const dateGroup = () => wb.getPivotDefinition(sheet, 0)?.fields[2]?.dateGroup ?? null;

  it('opens from Group… on a date label with the measured defaults', () => {
    openOn(4, 0);
    const group = item('pivotGroup');
    expect(group?.getAttribute('aria-disabled')).not.toBe('true');
    group?.click();
    expect(dialog()?.textContent).toContain('グループ化');
    const [start, end, days] = inputs();
    expect(checks().map((c) => c.checked)).toEqual([true, true]);
    expect(start?.value).toBe('2024/1/15');
    expect(end?.value).toBe('2024/7/23');
    expect(Array.from(periods()?.selectedOptions ?? []).map((o) => o.textContent)).toEqual(['月']);
    expect(days?.value).toBe('1');
    expect(days?.disabled).toBe(true);
  });

  it('groups by every chosen period in one undo step', () => {
    openOn(4, 0);
    item('pivotGroup')?.click();
    pick(PivotDatePeriod.Months, PivotDatePeriod.Years);
    ok()?.click();
    expect(dialog()).toBeNull();
    expect(dateGroup()).toMatchObject({
      periods: PivotDatePeriod.Months | PivotDatePeriod.Years,
      startSerial: -1,
      endSerial: -1,
    });
    expect(history.undo()).toBe(true);
    expect(dateGroup()).toBeNull();
  });

  it('takes a typed start as a fixed bound and Number of days for Days alone', () => {
    openOn(4, 0);
    item('pivotGroup')?.click();
    const [start, , days] = inputs();
    if (!start || !days) throw new Error('no inputs');
    start.value = '2024/1/1';
    start.dispatchEvent(new Event('input'));
    expect(checks()[0]?.checked).toBe(false);
    pick(PivotDatePeriod.Days);
    expect(days.disabled).toBe(false);
    days.value = '7';
    ok()?.click();
    expect(dateGroup()).toMatchObject({
      periods: PivotDatePeriod.Days,
      intervalDays: 7,
      startSerial: 45292,
      endSerial: -1,
    });
  });

  it('leaves Group… greyed on a text label', () => {
    expect(wb.setPivotFieldAxis(sheet, 0, 0, PivotAxis.Row)).toBe(true);
    const region = Array.from(wb.pivotCells(sheet)).find(
      (c) => c.kind === 1 && c.fieldName === 'Region',
    );
    if (!region) throw new Error('no region label');
    openOn(region.addr.row, region.addr.col);
    expect(item('pivotGroup')?.getAttribute('aria-disabled')).toBe('true');
  });

  it('keeps OK off with no period chosen', () => {
    openOn(4, 0);
    item('pivotGroup')?.click();
    pick();
    expect(ok()?.disabled).toBe(true);
  });
});
