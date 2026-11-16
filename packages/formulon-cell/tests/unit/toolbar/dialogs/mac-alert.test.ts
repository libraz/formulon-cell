import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { afterEach, describe, expect, it } from 'vitest';
import type { Range } from '../../../../src/engine/types.js';
import { addrKey, WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { defaultStrings, ja } from '../../../../src/i18n/strings.js';
import { attachFindReplace } from '../../../../src/interact/find-replace.js';
import { createSpreadsheetStore, type SpreadsheetStore } from '../../../../src/store/store.js';
import { confirmMergeLoseData } from '../../../../src/toolbar/dialogs/merge-confirm.js';
import { confirmMoveOverwrite } from '../../../../src/toolbar/dialogs/move-overwrite-confirm.js';
import { showMessage } from '../../../../src/toolbar/dialogs/prompt.js';
import { confirmSheetDelete } from '../../../../src/toolbar/dialogs/sheet-delete-confirm.js';

const MAC_JA = { spreadsheetProfileId: () => 'mac-ja_JP' };
const RANGE: Range = { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 };

const seed = (store: SpreadsheetStore, key: string, value: string): void => {
  store.setState((s) => {
    const cells = new Map(s.data.cells);
    cells.set(key, { value: { kind: 'text', value }, formula: null });
    return { ...s, data: { ...s.data, cells } };
  });
};

const dialog = (): HTMLElement => {
  const el = document.body.querySelector<HTMLElement>('[role="alertdialog"]');
  if (!el) throw new Error('no alert');
  return el;
};
const footerLabels = (): string[] =>
  [...dialog().querySelectorAll('.fc-fmtdlg__footer button')].map((b) => b.textContent ?? '');
const click = (label: string): void => {
  [...dialog().querySelectorAll('button')]
    .find((b) => b.textContent === label)
    ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
};
const expectAlertShape = (): void => {
  expect(dialog().classList.contains('fc-tb__dlg--alert')).toBe(true);
  expect(dialog().querySelector<HTMLElement>('.fc-fmtdlg__header')?.hidden).toBe(true);
};

afterEach(() => {
  document.body.replaceChildren();
});

describe('the Mac alert shape', () => {
  // Excel for Mac: a compact panel with no title bar, the message under an optional bold
  // headline, and its buttons filling the row.
  it('shows the merge warning as a Mac alert', async () => {
    const store = createSpreadsheetStore();
    seed(store, '0:0:0', 'a');
    seed(store, '0:0:1', 'b');
    const pending = confirmMergeLoseData(ja, store.getState(), RANGE, false, MAC_JA);
    await Promise.resolve();
    expectAlertShape();
    expect(footerLabels()).toEqual(['キャンセル', 'OK']);
    click('OK');
    expect(await pending).toBe(true);
  });

  it('keeps the titled dialog off the Mac profile', async () => {
    const store = createSpreadsheetStore();
    seed(store, '0:0:0', 'a');
    seed(store, '0:0:1', 'b');
    const pending = confirmMergeLoseData(defaultStrings, store.getState(), RANGE);
    await Promise.resolve();
    expect(dialog().classList.contains('fc-tb__dlg--alert')).toBe(false);
    expect(dialog().querySelector<HTMLElement>('.fc-fmtdlg__header')?.hidden).toBe(false);
    click(defaultStrings.ribbon.mergeLoseDataCancel);
    expect(await pending).toBe(false);
  });

  it('shows the sheet-delete and drag-overwrite warnings as Mac alerts', async () => {
    const wb = {
      spreadsheetProfileId: () => 'mac-ja_JP',
      cells: function* () {
        yield { addr: { sheet: 0, row: 0, col: 0 }, value: { kind: 'number', value: 1 } };
      },
    } as unknown as WorkbookHandle;
    const del = confirmSheetDelete(ja, wb, 0);
    await Promise.resolve();
    expectAlertShape();
    click('キャンセル');
    expect(await del).toBe(false);

    const move = confirmMoveOverwrite(ja, MAC_JA);
    await Promise.resolve();
    expectAlertShape();
    click('OK');
    expect(await move).toBe(true);
  });

  it('sets a bold headline above the message, with one button', async () => {
    const pending = showMessage({
      title: 'T',
      headline: 'Head',
      message: 'Body',
      okLabel: 'OK',
      alert: true,
    });
    expectAlertShape();
    expect(dialog().querySelector('.fc-tb__dlg__headline')?.textContent).toBe('Head');
    expect(dialog().querySelector('.fc-tb__dlg__message')?.textContent).toBe('Body');
    expect(footerLabels()).toEqual(['OK']);
    click('OK');
    await pending;
    expect(document.body.querySelector('[role="alertdialog"]')).toBeNull();
  });

  it('fills the button row: one button full width, two at equal width', () => {
    const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../../..');
    const css = readFileSync(join(root, 'src/styles/core/app/popups/alert.css'), 'utf8');
    const rule = css.match(
      /\.fc-tb__dlg--alert \.fc-fmtdlg__footer > \.fc-fmtdlg__btn\s*\{([^}]*)\}/,
    );
    expect(rule?.[1]).toMatch(/flex:\s*1 1 0/);
  });
});

describe('Find on the Mac with nothing found', () => {
  const mount = async (platform: 'mac' | 'default') => {
    const fcHost = document.createElement('div');
    fcHost.className = 'fc-host';
    fcHost.dataset.fcPlatform = platform;
    const host = document.createElement('div');
    host.tabIndex = -1;
    fcHost.appendChild(host);
    document.body.appendChild(fcHost);
    const store = createSpreadsheetStore();
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    const addr = { sheet: 0, row: 0, col: 0 };
    wb.setText(addr, 'apple');
    store.setState((s) => {
      const cells = new Map(s.data.cells);
      cells.set(addrKey(addr), { value: { kind: 'text', value: 'apple' }, formula: null });
      return { ...s, data: { ...s.data, cells } };
    });
    const handle = attachFindReplace({ host, store, wb, strings: ja, getLocale: () => 'ja-JP' });
    handle.open('find');
    const next = (query: string): void => {
      const input = host.querySelector<HTMLInputElement>('#fc-find-find-input');
      if (!input) throw new Error('no input');
      input.value = query;
      input.dispatchEvent(new Event('input', { bubbles: true }));
      input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
    };
    return { next };
  };

  // strman_0003 key 1c0, the product name dropped from the headline.
  it('says so in a Mac alert', async () => {
    const { next } = await mount('mac');
    next('zzz');
    expectAlertShape();
    expect(dialog().querySelector('.fc-tb__dlg__headline')?.textContent).toBe(
      '検索しているデータが見つかりません。',
    );
    expect(dialog().querySelector('.fc-tb__dlg__message')?.textContent).toBe(
      '現在のシートにデータが存在する場合は、入力内容を確認して、もう一度やり直してください。',
    );
    expect(footerLabels()).toEqual(['OK']);
  });

  it('shows no alert when a match is found', async () => {
    const { next } = await mount('mac');
    next('apple');
    expect(document.body.querySelector('[role="alertdialog"]')).toBeNull();
  });
});
