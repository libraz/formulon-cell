import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { formatAsTable } from '../../../src/commands/format-as-table.js';
import { History } from '../../../src/commands/history.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { ja } from '../../../src/i18n/strings/ja.js';
import { attachAccessibilityPane } from '../../../src/interact/accessibility-pane.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

const at = (row: number, col: number) => ({ sheet: 0, row, col });
const flush = () => new Promise<void>((resolve) => queueMicrotask(resolve));

describe('Mac accessibility checker pane', () => {
  let fcHost: HTMLElement;
  let dock: HTMLElement;
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;

  beforeEach(async () => {
    fcHost = document.createElement('div');
    fcHost.className = 'fc-host';
    fcHost.dataset.fcPlatform = 'mac';
    dock = document.createElement('div');
    dock.className = 'fc-host__taskpane-dock';
    dock.hidden = true;
    fcHost.appendChild(dock);
    document.body.appendChild(fcHost);
    store = createSpreadsheetStore();
    wb = await WorkbookHandle.createDefault();
    wb.attachStore(store);
  });

  afterEach(() => {
    wb.detachStore(store);
    wb.dispose();
    fcHost.remove();
  });

  const seedFailingBook = (): void => {
    wb.setText(at(2, 0), 'warning');
    mutators.setCellFormat(store, at(2, 0), { color: '#ff0000', fill: '#ffc0cb' });
    for (const row of [6, 7]) {
      wb.setNumber(at(row, 0), row);
      wb.setNumber(at(row, 1), row * 10);
    }
    formatAsTable(
      store,
      { sheet: 0, r0: 6, c0: 0, r1: 7, c1: 1 },
      {
        showHeader: false,
        workbook: wb,
      },
    );
  };

  const attach = () =>
    attachAccessibilityPane({
      host: fcHost,
      store,
      getWb: () => wb,
      getStrings: () => ja,
      history: new History(),
    });

  const rows = () =>
    [...dock.querySelectorAll<HTMLButtonElement>('.fc-a11ypane__rule')].map((row) => [
      row.querySelector('.fc-a11ypane__rule-name')?.textContent,
      row.querySelector('.fc-a11ypane__rule-count')?.textContent ?? 'pass',
    ]);

  it('shows the clean summary and every rule passing on a clean book', () => {
    const pane = attach();
    pane.open();
    expect(dock.hidden).toBe(false);
    expect(dock.querySelector('.fc-a11ypane__title')?.textContent).toBe('ユーザー補助アシスタント');
    expect(dock.querySelector('.fc-a11ypane__summary-text')?.textContent).toBe(
      '問題ないようです。問題は見つかりませんでした。',
    );
    expect(rows().every(([, state]) => state === 'pass')).toBe(true);
    pane.close();
    expect(dock.hidden).toBe(true);
    pane.detach();
  });

  it('draws the close button as a filled disc with an X cut out of it', () => {
    const pane = attach();
    pane.open();
    const svg = dock.querySelector('.fc-a11ypane__close svg');
    expect(svg?.querySelectorAll('path')).toHaveLength(2);
    expect(svg?.querySelector<SVGPathElement>('path[fill="none"]')?.style.stroke).toContain(
      '--fc-tb-icon-paper',
    );
    pane.detach();
  });

  it('counts the failing rules in desktop order and drills into contrast', async () => {
    seedFailingBook();
    const pane = attach();
    pane.open();
    expect(dock.querySelector('.fc-a11ypane__summary-title')?.textContent).toBe(
      '引き続き頑張ってください。',
    );
    expect([...dock.querySelectorAll('.fc-a11ypane__section')].map((h) => h.textContent)).toEqual([
      '色とコントラスト',
      'メディアとイラスト',
      'テーブル',
      'ドキュメント構造',
      'ドキュメント アクセス',
    ]);
    expect(rows()).toEqual([
      ['読み取りにくいテキストのコントラスト', '1'],
      ['赤の書式を使用しない', 'pass'],
      ['代替テキストなし', 'pass'],
      ['テーブル ヘッダーがありません', '1'],
      ['結合されたセルの使用', 'pass'],
      ['既定のシート名', 'pass'],
      ['制限付きアクセス', 'pass'],
    ]);

    dock.querySelector<HTMLButtonElement>('[data-rule="contrast"]')?.click();
    expect(dock.querySelector('.fc-a11ypane__counter')?.textContent).toBe('1/1 - 色とコントラスト');
    expect(dock.querySelector('.fc-a11ypane__card-title')?.textContent).toBe(
      '読み取りにくいテキストのコントラストです',
    );
    expect(store.getState().selection.range).toEqual({ sheet: 0, r0: 2, c0: 0, r1: 2, c1: 0 });
    const swatches = dock.querySelectorAll<HTMLButtonElement>('.fc-a11ypane__swatch');
    expect(swatches).toHaveLength(2);
    expect([...dock.querySelectorAll('.fc-a11ypane__fix')].map((b) => b.textContent)).toEqual([
      'その他のフォントの色',
      '塗りつぶしの色',
      'セルのフォントの書式設定',
    ]);

    // Applying a suggested colour fixes the finding; the pane returns to the overview.
    swatches[0]?.click();
    await flush();
    expect(dock.querySelector('.fc-a11ypane__counter')).toBeNull();
    expect(rows()[0]).toEqual(['読み取りにくいテキストのコントラスト', 'pass']);
    pane.detach();
  });

  it('selects the table body and inserts its header row', async () => {
    seedFailingBook();
    const pane = attach();
    pane.open();
    dock.querySelector<HTMLButtonElement>('[data-rule="tableHeader"]')?.click();
    expect(dock.querySelector('.fc-a11ypane__counter')?.textContent).toBe('1/1 - テーブル');
    expect(store.getState().selection.range).toEqual({ sheet: 0, r0: 6, c0: 0, r1: 7, c1: 1 });
    const fix = dock.querySelector<HTMLButtonElement>('.fc-a11ypane__fix');
    expect(fix?.textContent).toBe('ヘッダー行の挿入');
    fix?.click();
    await flush();
    expect(store.getState().tables.tables[0]?.showHeader).toBe(true);
    expect(rows()[3]).toEqual(['テーブル ヘッダーがありません', 'pass']);
    pane.detach();
  });
});
