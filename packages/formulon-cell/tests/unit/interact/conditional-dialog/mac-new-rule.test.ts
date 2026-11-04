import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { ja } from '../../../../src/i18n/strings.js';
import { attachConditionalDialog } from '../../../../src/interact/conditional-dialog.js';
import { createSpreadsheetStore, type SpreadsheetStore } from '../../../../src/store/store.js';
import { setRange } from './fixtures.js';

// Excel for Mac (ja): New Formatting Rule offers five rule types, each with a condition popup.
describe('Mac New Formatting Rule layout', () => {
  let host: HTMLElement;
  let store: SpreadsheetStore;

  beforeEach(() => {
    host = document.createElement('div');
    host.tabIndex = -1;
    host.dataset.fcPlatform = 'mac';
    document.body.appendChild(host);
    store = createSpreadsheetStore();
    setRange(store, 1, 1, 4, 3);
  });
  afterEach(() => {
    document.body.innerHTML = '';
  });

  const open = (options: Parameters<ReturnType<typeof attachConditionalDialog>['open']>[0]) => {
    const handle = attachConditionalDialog({ host, store, strings: ja });
    handle.open({ mode: 'new', ...options });
    return handle;
  };
  const select = (name: string): HTMLSelectElement => {
    const el = document.querySelector<HTMLSelectElement>(`select[data-cf-mac-${name}]`);
    if (!el) throw new Error(`missing ${name}`);
    return el;
  };
  const labels = (el: HTMLSelectElement): string[] =>
    Array.from(el.options, (option) => option.textContent ?? '');
  const choose = (el: HTMLSelectElement, value: string): void => {
    el.value = value;
    el.dispatchEvent(new Event('change', { bubbles: true }));
  };
  const kind = (): string =>
    document.querySelector<HTMLSelectElement>('select[data-cf-kind]')?.value ?? '';

  it('offers the five rule types', () => {
    open({ kind: 'cell-value', cellValueOp: '>' });
    expect(labels(select('type'))).toEqual([
      '指定の値を含むセルだけを書式設定',
      '上位または下位に入る値だけを書式設定',
      '平均より上または下の値だけを書式設定',
      '一意の値または重複する値だけを書式設定',
      '数式を使用して、書式設定するセルを決定',
    ]);
    expect(select('type').value).toBe('cells');
  });

  it('opens a preset on its type and condition', () => {
    open({ kind: 'text-contains' });
    expect(select('type').value).toBe('cells');
    expect(select('cond').value).toBe('text-contains');
    expect(labels(select('cond'))).toEqual([
      'セルの値',
      '特定の文字列',
      '日付',
      '空白セル',
      '空白以外のセル',
      'エラー値',
      'エラーなし',
    ]);
  });

  it('maps each type and condition onto the rule model', () => {
    open({ kind: 'cell-value' });
    choose(select('type'), 'unique');
    expect(kind()).toBe('duplicates');
    expect(labels(select('cond'))).toEqual(['重複', '一意']);
    choose(select('cond'), 'unique');
    expect(kind()).toBe('unique');
    choose(select('type'), 'topBottom');
    expect(kind()).toBe('top-bottom');
    choose(select('cond'), 'bottom');
    expect(
      document.querySelector<HTMLSelectElement>('.fc-conddlg__sub:not([hidden]) select')?.value,
    ).toBe('bottom');
    choose(select('type'), 'average');
    expect(kind()).toBe('average');
    choose(select('type'), 'formula');
    expect(kind()).toBe('formula');
    expect(select('cond').hidden).toBe(true);
    choose(select('type'), 'cells');
    expect(kind()).toBe('cell-value');
  });

  it('seeds the topBottom and average presets', () => {
    open({ kind: 'top-bottom', topBottomMode: 'bottom', topBottomPercent: true });
    expect(select('type').value).toBe('topBottom');
    expect(select('cond').value).toBe('bottom');
  });

  it('hides the desktop kind row and has none of this on the desktop profile', () => {
    open({ kind: 'cell-value' });
    expect(document.querySelector<HTMLElement>('.fc-conddlg__kind-row')?.hidden).toBe(true);
    document.body.innerHTML = '';
    host = document.createElement('div');
    host.dataset.fcPlatform = 'default';
    document.body.appendChild(host);
    open({ kind: 'cell-value' });
    expect(document.querySelector<HTMLElement>('.fc-conddlg__mac-types')?.hidden).toBe(true);
    expect(document.querySelector<HTMLElement>('.fc-conddlg__kind-row')?.hidden).toBe(false);
  });
});
