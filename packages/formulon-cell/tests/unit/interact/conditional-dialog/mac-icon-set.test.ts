import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { ja } from '../../../../src/i18n/strings.js';
import { attachConditionalDialog } from '../../../../src/interact/conditional-dialog.js';
import { createSpreadsheetStore, type SpreadsheetStore } from '../../../../src/store/store.js';
import { setRange } from './fixtures.js';

// Excel for Mac (ja): New Formatting Rule with the Icon Set style.
describe('Mac icon set rule editor', () => {
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

  it('lists the twenty icon sets with Excel names and order, starting on the unrimmed signals', () => {
    const handle = attachConditionalDialog({ host, store, strings: ja });
    handle.open({ mode: 'new', kind: 'icon-set' });
    const family = document.querySelector<HTMLSelectElement>('select[data-cf-icon-family]');
    expect(family?.value).toBe('traffic3');
    expect(Array.from(family?.options ?? [], (option) => option.textContent)).toEqual([
      '3 つの矢印 (色分け)',
      '3 つの矢印 (灰色)',
      '3 種類の三角形',
      '4 つの矢印 (灰色)',
      '4 つの矢印 (色分け)',
      '5 つの矢印 (灰色)',
      '5 つの矢印 (色分け)',
      '3 つの信号 (枠なし)',
      '3 つの信号 (枠あり)',
      '3 つの図形',
      '4 つの信号',
      '赤と黒の丸',
      '3 つの記号 (丸囲み)',
      '3 つの記号 (丸囲みなし)',
      '3 つのフラグ',
      '3 種類の星',
      '4 つの評価',
      '白黒の丸',
      '5 つの評価',
      '5 種類のボックス',
    ]);
  });

  it('words the two option checkboxes as Excel does', () => {
    const handle = attachConditionalDialog({ host, store, strings: ja });
    handle.open({ mode: 'new', kind: 'icon-set' });
    const texts = Array.from(
      document.querySelectorAll<HTMLElement>('.fc-conddlg__sub .fc-fmtdlg__check span'),
    ).map((span) => span.textContent);
    expect(texts).toContain('アイコンの順序を逆にする');
    expect(texts).toContain('アイコンのみ表示');
  });

  it('keeps the arrows default on the desktop profile', () => {
    host.dataset.fcPlatform = 'default';
    const handle = attachConditionalDialog({ host, store, strings: ja });
    handle.open({ mode: 'new', kind: 'icon-set' });
    expect(document.querySelector<HTMLSelectElement>('select[data-cf-icon-family]')?.value).toBe(
      'arrows3',
    );
  });
});
