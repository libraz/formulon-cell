import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { createPivotTableFromRange } from '../../../src/commands/pivot-table.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { ja } from '../../../src/i18n/strings/ja.js';
import { attachContextMenu, type ContextMenuHandle } from '../../../src/interact/context-menu.js';
import { disposeOverlayPortal } from '../../../src/interact/overlay-portal.js';
import {
  groupPivotField,
  pivotAutoGroupInterval,
  pivotGroupSettings,
  pivotGroupTargetAt,
} from '../../../src/interact/pivot-group.js';
import { createSpreadsheetStore, type SpreadsheetStore } from '../../../src/store/store.js';
import { item, setSelectionRanges } from './context-menu/fixtures.js';

describe('the interval the Grouping dialog proposes', () => {
  it.each([
    [3, 1],
    [30, 10],
    [300, 100],
    [2.999, 0.1],
    [29.99, 1],
    [299.9, 10],
    [0.05, 0.01],
    [0.19, 0.01],
    [0.5, 0.1],
    [1, 0.1],
    [3.3, 1],
    [25, 1],
    [50, 10],
    [110, 10],
    [2997, 100],
    [1e6, 1e5],
    [3.6 - 0.3, 1],
  ])('takes %d to %d', (span, interval) => {
    expect(pivotAutoGroupInterval(span)).toBe(interval);
  });
});

// Amt 10..120 step 10 on rows 1-12 of K, Qty in L; the report's Amt labels run down A5:A16.
describe('PivotTable Grouping of a number field', () => {
  let host: HTMLElement;
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;
  let history: History;
  let detach: ContextMenuHandle;

  beforeEach(async () => {
    host = document.createElement('div');
    host.tabIndex = -1;
    document.body.appendChild(host);
    wb = await WorkbookHandle.createDefault();
    wb.setText({ sheet: 0, row: 0, col: 10 }, 'Amt');
    wb.setText({ sheet: 0, row: 0, col: 11 }, 'Qty');
    for (let i = 1; i <= 12; i += 1) {
      wb.setNumber({ sheet: 0, row: i, col: 10 }, i * 10);
      wb.setNumber({ sheet: 0, row: i, col: 11 }, i);
    }
    const created = createPivotTableFromRange(wb, {
      source: { sheet: 0, r0: 0, c0: 10, r1: 12, c1: 11 },
      destination: { sheet: 0, row: 2, col: 0 },
      rowField: 'Amt',
      valueField: 'Qty',
    });
    expect(created.ok).toBe(true);
    store = createSpreadsheetStore();
    history = new History();
    detach = attachContextMenu({ host, store, wb, strings: ja, history });
  });

  afterEach(() => {
    detach();
    disposeOverlayPortal(host);
    document.body.innerHTML = '';
    wb.dispose();
  });

  const label = () => {
    const cell = Array.from(wb.pivotCells(0)).find((c) => c.kind === 1);
    if (!cell) throw new Error('no label');
    return cell.addr;
  };
  const labels = () =>
    Array.from(wb.pivotCells(0))
      .filter((c) => c.kind === 1)
      .map((c) =>
        c.value.kind === 'text' || c.value.kind === 'number' ? String(c.value.value) : '',
      );
  const numericGroup = () => wb.getPivotDefinition(0, 0)?.fields[0]?.numericGroup ?? null;
  const openMenu = (): void => {
    const addr = label();
    setSelectionRanges(
      store,
      { sheet: 0, r0: addr.row, c0: addr.col, r1: addr.row, c1: addr.col },
      [],
    );
    host.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'F10', shiftKey: true, bubbles: true, cancelable: true }),
    );
  };
  const dialog = () => document.querySelector<HTMLElement>('.fc-pivotgroup');
  const texts = () =>
    Array.from(dialog()?.querySelectorAll<HTMLInputElement>('input[type="text"]') ?? []);

  it('opens on the smallest and largest value by the proposed interval', () => {
    const target = pivotGroupTargetAt(wb, label());
    expect(target?.values).toBe('number');
    expect(target && pivotGroupSettings(wb, target)).toEqual({
      kind: 'number',
      autoStart: true,
      autoEnd: true,
      start: 10,
      end: 120,
      interval: 10,
    });
  });

  it('groups into the engine-labelled ranges in one undo step', () => {
    const target = pivotGroupTargetAt(wb, label());
    const settings = target && pivotGroupSettings(wb, target);
    if (!target || !settings) throw new Error('no target');
    const ctx = {
      wb,
      history,
      store,
      origin: 'ribbon' as const,
      commandId: 'pivotGroupSelection',
      onChanged: () => undefined,
    };
    expect(groupPivotField(ctx, target, settings)).toBe(true);
    expect(numericGroup()).toEqual({ interval: 10, start: null, end: null });
    expect(labels()).toContain('10-19');
    history.undo();
    expect(numericGroup()).toBeNull();
  });

  it('runs from Group… and comes off with Ungroup', () => {
    openMenu();
    item('pivotGroup')?.click();
    expect(texts().map((input) => input.value)).toEqual(['10', '120', '10']);
    const by = texts()[2];
    if (!by) throw new Error('no By');
    by.value = '50';
    dialog()?.querySelector<HTMLButtonElement>('.fc-pivotgroup__footer button:last-child')?.click();
    expect(dialog()).toBeNull();
    expect(numericGroup()).toEqual({ interval: 50, start: null, end: null });
    openMenu();
    const ungroup = item('pivotUngroup');
    expect(ungroup?.getAttribute('aria-disabled')).not.toBe('true');
    ungroup?.click();
    expect(numericGroup()).toBeNull();
  });
});
