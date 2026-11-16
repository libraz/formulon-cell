import { afterEach, describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import {
  pivotFieldFilterModel,
  setPivotFieldItemsChecked,
} from '../../../src/commands/pivot-filter.js';
import { PivotAxis } from '../../../src/engine/types.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { ja } from '../../../src/i18n/strings/ja.js';
import { disposeOverlayPortal } from '../../../src/interact/overlay-portal.js';
import {
  collapsePivotField,
  expandPivotField,
  isInnermostRowItem,
  pivotRowItemAt,
  togglePivotItemDetail,
} from '../../../src/interact/pivot-detail.js';
import {
  openPivotShowDetailDialog,
  pivotShowDetailFields,
} from '../../../src/interact/pivot-detail-picker.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';
import { loadFixture } from '../engine/pivot-excel-parity/fixture.js';

// Rows Region (East, North, South) then Product: East on row 4, its A on row 5, total on row 12.
const books: WorkbookHandle[] = [];
afterEach(() => {
  for (const wb of books.splice(0)) wb.dispose();
});

const setup = async () => {
  const wb = await loadFixture('ja_d7_two_rows_values_cols.xlsx');
  books.push(wb);
  const history = new History();
  const ctx = {
    wb,
    history,
    store: createSpreadsheetStore(),
    origin: 'ribbon' as const,
    commandId: 'pivotCollapseField',
    onChanged: () => undefined,
  };
  const detail = (field: number) => wb.getPivotField(0, 0, field)?.items.map((i) => i.showDetail);
  return { wb, history, ctx, detail };
};

describe('PivotTable expand and collapse', () => {
  it('names the field and item of a row label, and nothing for other cells', async () => {
    const { wb } = await setup();
    const east = pivotRowItemAt(wb, { sheet: 0, row: 3, col: 0 });
    expect(east).toMatchObject({ fieldIdx: 0, itemIdx: 0, level: 0, showDetail: true });
    expect(east && isInnermostRowItem(east)).toBe(false);
    const southB = pivotRowItemAt(wb, { sheet: 0, row: 9, col: 0 });
    expect(southB).toMatchObject({ fieldIdx: 1, level: 1 });
    expect(southB && isInnermostRowItem(southB)).toBe(true);
    expect(pivotRowItemAt(wb, { sheet: 0, row: 3, col: 1 })).toBeNull();
    expect(pivotRowItemAt(wb, { sheet: 0, row: 2, col: 0 })).toBeNull();
    expect(pivotRowItemAt(wb, { sheet: 0, row: 11, col: 0 })).toBeNull();
  });

  it('turns one item over from its button, one undo step', async () => {
    const { wb, history, ctx, detail } = await setup();
    const east = pivotRowItemAt(wb, { sheet: 0, row: 3, col: 0 });
    if (!east) throw new Error('no item');
    expect(togglePivotItemDetail(ctx, east)).toBe(true);
    expect(detail(0)).toEqual([false, true, true]);
    history.undo();
    expect(detail(0)).toEqual([true, true, true]);
    history.redo();
    expect(detail(0)).toEqual([false, true, true]);
    const collapsed = pivotRowItemAt(wb, { sheet: 0, row: 3, col: 0 });
    expect(collapsed?.showDetail).toBe(false);
    if (collapsed) togglePivotItemDetail(ctx, collapsed);
    expect(detail(0)).toEqual([true, true, true]);
  });

  it('collapses the field above an innermost item, and expands an outer field again', async () => {
    const { wb, history, ctx, detail } = await setup();
    const inner = pivotRowItemAt(wb, { sheet: 0, row: 4, col: 0 });
    if (!inner) throw new Error('no item');
    expect(collapsePivotField(ctx, inner)).toBe(true);
    expect(detail(0)).toEqual([false, false, false]);
    expect(wb.getPivotField(0, 0, 0)?.showDetail).toBe(false);
    expect(detail(1)?.every(Boolean)).toBe(true);
    const outer = pivotRowItemAt(wb, { sheet: 0, row: 3, col: 0 });
    if (!outer) throw new Error('no item');
    expect(expandPivotField(ctx, outer)).toBe(true);
    expect(detail(0)).toEqual([true, true, true]);
    history.undo();
    expect(detail(0)).toEqual([false, false, false]);
    history.undo();
    expect(detail(0)).toEqual([true, true, true]);
  });

  it('keeps a collapsed item collapsed when the field filter rewrites its items', async () => {
    const { wb, ctx, detail } = await setup();
    const east = pivotRowItemAt(wb, { sheet: 0, row: 3, col: 0 });
    if (!east) throw new Error('no item');
    togglePivotItemDetail(ctx, east);
    const items = pivotFieldFilterModel(wb, 0, 0, 0)?.items ?? [];
    const hideNorth = items.map((item) => ({ ...item, checked: item.value !== 'North' }));
    expect(setPivotFieldItemsChecked(wb, 0, 0, 0, hideNorth)).toBe(true);
    expect(wb.getPivotField(0, 0, 0)?.items.map((i) => i.visible)).toEqual([true, false, true]);
    expect(detail(0)).toEqual([false, true, true]);
  });

  it('leaves Expand Field on the innermost field to the Show Detail picker', async () => {
    const { wb, ctx, detail } = await setup();
    const inner = pivotRowItemAt(wb, { sheet: 0, row: 4, col: 0 });
    if (!inner) throw new Error('no item');
    expect(expandPivotField(ctx, inner)).toBe(false);
    expect(detail(1)?.every(Boolean)).toBe(true);
  });
});

describe('the Show Detail picker of an innermost row item', () => {
  afterEach(() => {
    disposeOverlayPortal(document.body);
    document.body.innerHTML = '';
  });

  const open = async () => {
    const env = await setup();
    const host = document.createElement('div');
    document.body.appendChild(host);
    const inner = pivotRowItemAt(env.wb, { sheet: 0, row: 4, col: 0 });
    if (!inner) throw new Error('no item');
    return { ...env, host, inner, picker: { ...env.ctx, host, strings: ja } };
  };
  const dialog = () => document.querySelector<HTMLElement>('.fc-pivotdetail');
  const options = () =>
    Array.from(dialog()?.querySelectorAll<HTMLButtonElement>('[role="option"]') ?? []);
  const footerButton = (label: string) => {
    const found = Array.from(dialog()?.querySelectorAll<HTMLButtonElement>('button') ?? []).find(
      (b) => b.textContent === label,
    );
    if (!found) throw new Error(`missing ${label}`);
    return found;
  };
  const rowOrder = (wb: WorkbookHandle) => wb.getPivotDefinition(0, 0)?.rowFieldOrder;

  it('lists the fields outside the row area, in field order, with column and filter fields unpickable', async () => {
    const { wb, inner } = await open();
    expect(pivotShowDetailFields(wb, inner)).toEqual([
      { fieldIdx: 2, name: 'Amt', pickable: true },
      { fieldIdx: 3, name: 'Qty', pickable: true },
    ]);
    expect(wb.setPivotFieldAxis(0, 0, 3, PivotAxis.Col)).toBe(true);
    expect(pivotShowDetailFields(wb, inner)).toEqual([
      { fieldIdx: 2, name: 'Amt', pickable: true },
      { fieldIdx: 3, name: 'Qty', pickable: false },
    ]);
  });

  it('opens with the prompt and nothing chosen, so OK waits for a choice', async () => {
    const { picker, inner } = await open();
    expect(openPivotShowDetailDialog(picker, inner)).toBe(true);
    expect(dialog()?.getAttribute('aria-label')).toBe('詳細を表示');
    expect(dialog()?.textContent).toContain('表示する詳細を含むフィールドを選んでください。');
    expect(options().map((o) => o.textContent)).toEqual(['Amt', 'Qty']);
    expect(options().map((o) => o.getAttribute('aria-selected'))).toEqual(['false', 'false']);
    expect(footerButton('OK').disabled).toBe(true);
    expect(footerButton('キャンセル').disabled).toBe(false);
    footerButton('OK').click();
    expect(dialog()).not.toBeNull();
  });

  it('closes on Cancel and leaves the report and the undo stack alone', async () => {
    const { wb, history, picker, inner } = await open();
    openPivotShowDetailDialog(picker, inner);
    options()[0]?.click();
    footerButton('キャンセル').click();
    expect(dialog()).toBeNull();
    expect(rowOrder(wb)).toEqual([0, 1]);
    expect(history.undo()).toBe(false);
  });

  it('adds the chosen field as the next row field under every item, one undo step', async () => {
    const { wb, history, picker, inner, detail } = await open();
    openPivotShowDetailDialog(picker, inner);
    options()[0]?.click();
    expect(options()[0]?.getAttribute('aria-selected')).toBe('true');
    expect(footerButton('OK').disabled).toBe(false);
    footerButton('OK').click();
    expect(dialog()).toBeNull();
    expect(rowOrder(wb)).toEqual([0, 1, 2]);
    expect(wb.getPivotDefinition(0, 0)?.dataFields.length).toBe(2);
    expect(detail(1)?.every(Boolean)).toBe(true);
    expect(detail(2)?.every(Boolean)).toBe(true);
    expect(history.undo()).toBe(true);
    expect(rowOrder(wb)).toEqual([0, 1]);
    expect(wb.getPivotField(0, 0, 2)?.axis).not.toBe(PivotAxis.Row);
    expect(wb.getPivotDefinition(0, 0)?.dataFields.length).toBe(2);
    expect(history.undo()).toBe(false);
    expect(history.redo()).toBe(true);
    expect(rowOrder(wb)).toEqual([0, 1, 2]);
  });

  it('does not open on an outer item, whose Expand Field needs no picker', async () => {
    const { wb, picker } = await open();
    const outer = pivotRowItemAt(wb, { sheet: 0, row: 3, col: 0 });
    if (!outer) throw new Error('no item');
    expect(openPivotShowDetailDialog(picker, outer)).toBe(false);
    expect(dialog()).toBeNull();
  });
});
