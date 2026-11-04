import { afterEach, describe, expect, it, vi } from 'vitest';
import { ja } from '../../../src/i18n/strings.js';
import {
  type CellStyleDialogMacOptions,
  describeCellStyleFormat,
  showCellStyleDialog,
} from '../../../src/toolbar/dialogs/cell-style.js';

// Excel for Mac (ja) Style dialog: name, preview, property checklist, Format button.
const mac = (onFormat = vi.fn()): CellStyleDialogMacOptions => ({
  strings: {
    title: 'セル スタイルの変更',
    nameLabel: 'スタイル名:',
    previewLabel: 'プレビュー:',
    caption: '次のプロパティに対して、選択されたスタイルが書式設定されています:',
    propertyColumn: 'スタイル プロパティ',
    formatColumn: '書式',
    properties: {
      number: '数値',
      alignment: '配置',
      font: 'フォント',
      border: '枠線',
      fill: 'フィル',
      protection: '保護',
    },
    formatButton: '書式...',
  },
  summaries: {
    number: 'G/標準',
    alignment: '配置: 下詰め, 標準',
    font: 'ＭＳ Ｐゴシック (本文) 11, テキスト 1',
    border: '罫線なし',
    fill: '網かけなし',
    protection: 'ロック',
  },
  preview: { fill: '#c6efce', color: '#006100', bold: true },
  onFormat,
});

const q = <T extends HTMLElement>(sel: string): T => {
  const el = document.querySelector<T>(sel);
  if (!el) throw new Error(`missing ${sel}`);
  return el;
};

describe('Mac Style dialog', () => {
  afterEach(() => {
    document.body.replaceChildren();
  });

  it('shows the name, a preview and the property checklist with its format summary', async () => {
    const pending = showCellStyleDialog({
      title: 'ignored',
      strings: ja,
      initialName: '良い',
      mac: mac(),
    });
    expect(q('.fc-fmtdlg__header').textContent).toBe('セル スタイルの変更');
    expect(q('.fc-cellstyledlg__namerow').textContent).toContain('スタイル名:');
    const preview = q<HTMLElement>('.fc-cellstyledlg__preview');
    expect(preview.textContent).toBe('良い');
    expect(preview.style.background).toBe('#c6efce');
    expect(preview.style.fontWeight).toBe('700');
    expect(q('.fc-cellstyledlg__caption').textContent).toContain('選択されたスタイル');
    expect(
      Array.from(document.querySelectorAll('.fc-cellstyledlg__head span'), (el) => el.textContent),
    ).toEqual(['スタイル プロパティ', '書式']);
    const rows = Array.from(document.querySelectorAll('.fc-cellstyledlg__row'));
    expect(
      rows.map((row) => [
        row.querySelector('label')?.textContent,
        row.querySelector<HTMLInputElement>('input')?.checked,
        row.querySelector('.fc-cellstyledlg__summary')?.textContent,
      ]),
    ).toEqual([
      ['数値', true, 'G/標準'],
      ['配置', true, '配置: 下詰め, 標準'],
      ['フォント', true, 'ＭＳ Ｐゴシック (本文) 11, テキスト 1'],
      ['枠線', true, '罫線なし'],
      ['フィル', true, '網かけなし'],
      ['保護', true, 'ロック'],
    ]);
    document.querySelector<HTMLButtonElement>('.fc-fmtdlg__footer button:nth-child(2)')?.click();
    await pending;
  });

  it('puts Format on the left and returns the unchecked properties', async () => {
    const onFormat = vi.fn();
    const pending = showCellStyleDialog({ title: 'x', strings: ja, mac: mac(onFormat) });
    const footer = Array.from(document.querySelectorAll('.fc-fmtdlg__footer button'));
    expect(footer.map((b) => b.textContent)).toEqual(['書式...', 'キャンセル', 'OK']);
    (footer[0] as HTMLButtonElement).click();
    expect(onFormat).toHaveBeenCalledTimes(1);
    document.querySelector<HTMLInputElement>('.fc-cellstyledlg__row input')?.click();
    (footer[2] as HTMLButtonElement).click();
    const result = await pending;
    expect(result?.include).toEqual({
      number: false,
      alignment: true,
      font: true,
      border: true,
      fill: true,
      protection: true,
    });
  });
});

describe('describeCellStyleFormat', () => {
  const t = {
    general: 'G/標準',
    alignment: '配置: {vertical}, {horizontal}',
    vertical: { top: '上詰め', middle: '中央', bottom: '下詰め' },
    horizontal: { general: '標準', left: '左詰め', center: '中央揃え', right: '右詰め' },
    noBorder: '罫線なし',
    border: '罫線あり',
    noFill: '網かけなし',
    locked: 'ロック',
    unlocked: 'ロックなし',
  };

  it('reads the plain Normal style as the measured summaries', () => {
    expect(describeCellStyleFormat(undefined, t)).toMatchObject({
      number: 'G/標準',
      alignment: '配置: 下詰め, 標準',
      border: '罫線なし',
      fill: '網かけなし',
      protection: 'ロック',
    });
  });

  it('reports the format a style sets', () => {
    expect(
      describeCellStyleFormat(
        {
          align: 'center',
          vAlign: 'top',
          fill: '#ffc7ce',
          fontFamily: 'Arial',
          fontSize: 12,
          locked: false,
        },
        t,
      ),
    ).toMatchObject({
      alignment: '配置: 上詰め, 中央揃え',
      fill: '#ffc7ce',
      font: 'Arial 12',
      protection: 'ロックなし',
    });
  });
});
