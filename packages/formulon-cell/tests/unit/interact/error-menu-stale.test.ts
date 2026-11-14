import { afterEach, describe, expect, it, vi } from 'vitest';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { ja } from '../../../src/i18n/strings/ja.js';
import { attachErrorMenu } from '../../../src/interact/error-menu.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const B1 = { sheet: 0, row: 0, col: 1 };

/** Manual calculation with B1 = A1*2 left stale by a later write to A1. */
const staleBook = async () => {
  const wb = await WorkbookHandle.createDefault();
  wb.setNumber({ sheet: 0, row: 0, col: 0 }, 1);
  wb.setFormula(B1, '=A1*2');
  wb.setCalcMode(1);
  wb.setNumber({ sheet: 0, row: 0, col: 0 }, 5);
  const store = createSpreadsheetStore();
  mutators.replaceCells(store, wb.cells(0));
  return { wb, store };
};

const openMenu = (wb: WorkbookHandle, store: ReturnType<typeof createSpreadsheetStore>) => {
  const fcHost = document.createElement('div');
  fcHost.className = 'fc-host';
  fcHost.dataset.fcPlatform = 'mac';
  document.body.appendChild(fcHost);
  const onCalculate = vi.fn();
  const handle = attachErrorMenu({
    host: fcHost,
    store,
    getWb: () => wb,
    strings: ja,
    onCalculate,
  });
  handle.open(B1, 10, 10, 'error');
  const items = [...document.querySelectorAll<HTMLButtonElement>('.fc-errmenu__item')];
  return {
    heading: document.querySelector('.fc-errmenu__heading')?.textContent,
    labels: items.map((b) => b.textContent),
    click: (label: string) => items.find((b) => b.textContent === label)?.click(),
    onCalculate,
    handle,
  };
};

afterEach(() => {
  document.body.replaceChildren();
});

describe('Mac error menu on a stale formula value', () => {
  it('heads the menu Stale Value and lists the stale actions in the desktop order', async () => {
    const { wb, store } = await staleBook();
    try {
      expect(wb.isCellStale(B1)).toBe(true);
      const { heading, labels, handle } = openMenu(wb, store);
      expect(heading).toBe('古くなった値');
      expect(labels).toEqual([
        'このエラーに関するヘルプ(H)',
        '今すぐ計算(C) (F9)',
        '自動計算に切り替える(A)',
        '古くなったセルのフォーマットを無効にする(D)',
        'エラーを無視する(I)',
        '数式バーで編集(F)',
        'エラー チェック オプション...',
      ]);
      handle.detach();
    } finally {
      wb.dispose();
    }
  });

  it('runs Calculate Now, Switch to Automatic and Disable Format Stale Cells', async () => {
    const { wb, store } = await staleBook();
    try {
      let menu = openMenu(wb, store);
      menu.click('今すぐ計算(C) (F9)');
      expect(menu.onCalculate).toHaveBeenLastCalledWith('calculate-now');
      menu.handle.detach();
      menu = openMenu(wb, store);
      menu.click('自動計算に切り替える(A)');
      expect(menu.onCalculate).toHaveBeenLastCalledWith('auto');
      menu.handle.detach();
      menu = openMenu(wb, store);
      menu.click('古くなったセルのフォーマットを無効にする(D)');
      expect(store.getState().ui.formatStaleValues).toBe(false);
      menu.handle.detach();
    } finally {
      wb.dispose();
    }
  });

  it('falls back to the generic menu once the preference is off', async () => {
    const { wb, store } = await staleBook();
    try {
      mutators.setFormatStaleValues(store, false);
      const { heading, handle } = openMenu(wb, store);
      expect(heading).not.toBe('古くなった値');
      handle.detach();
    } finally {
      wb.dispose();
    }
  });
});
