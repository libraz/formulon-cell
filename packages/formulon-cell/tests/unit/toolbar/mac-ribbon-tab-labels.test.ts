import { describe, expect, it } from 'vitest';
import { buildRibbonModel } from '../../../src/toolbar/ribbon-model.js';

/** The app reports these zero-width spaces between the words of some names. */
const ZWSP = String.fromCodePoint(0x200b);

// Control names the Japanese Mac app reports on each tab (AX), in order.
const labelsOf = (tabId: string): string[] =>
  (buildRibbonModel('ja', { profile: 'excel365Mac' }).find((tab) => tab.id === tabId)?.groups ?? [])
    .flatMap((g) => g.commands)
    .filter((c) => c.kind !== 'break')
    .map((c) => c.label || c.title);

describe('Mac ribbon control names', () => {
  it('Data tab', () => {
    expect(labelsOf('data')).toEqual([
      'データ ファイル指定 (Power Query)',
      '画像から',
      'すべて更新',
      'クエリ  接続',
      'プロパティ',
      'ブックのリンク',
      'データの種類',
      '昇順',
      '降順',
      '並べ替え',
      'フィルター',
      'クリア',
      '再適用',
      '詳細設定',
      `区切り位置${ZWSP}${ZWSP}`,
      'データのクリーニング',
      'データの分析',
      'フラッシュ フィル',
      '重複を削除',
      '入力規則',
      '統合',
      'What-If 分析',
      'グループ化',
      'グループ解除',
      '小計',
      '詳細データの表示',
      '詳細を表示しない',
      '分析ツール',
    ]);
  });

  it('Formulas tab', () => {
    const names = labelsOf('formulas');
    for (const expected of [
      'オート SUM',
      'ネーム マネージャー',
      '参照元のトレース',
      '参照先のトレース',
      'トレース矢印の削除',
      'ウォッチ ウィンドウ',
      '再計算実行',
      'エラー チェック',
    ]) {
      expect(names).toContain(expected);
    }
  });

  it('Review tab', () => {
    expect(labelsOf('review')).toEqual([
      'スペル チェック',
      '類義語辞典',
      'ブックの統計情報',
      'アクセシビリティ チェック',
      '翻訳',
      '変更内容を表示',
      '新しいコメント',
      '削除',
      `前の${ZWSP}${ZWSP}コメント`,
      `次の${ZWSP}${ZWSP}コメント`,
      'コメントの表示',
      'メモ',
      'シートの保護',
      'ブックの保護',
      '常に読み取り専用で開く',
      'インクを非表示にする',
    ]);
  });

  it('View tab window and layout names', () => {
    const names = labelsOf('view');
    for (const expected of [
      '改ページ プレビュー',
      'ページ レイアウト',
      '新しいウインドウを開く',
      'ウインドウの整列',
      'ウインドウ枠の固定',
      '表示しない',
      'ウインドウを元の位置に戻す',
      'ウインドウの切り替え',
    ]) {
      expect(names).toContain(expected);
    }
  });

  it('Page Layout and Insert names', () => {
    expect(labelsOf('pageLayout')).toEqual(expect.arrayContaining(['幅:', '高さ:']));
    expect(labelsOf('insert')).toContain('記号/文字');
  });

  it('View tab shows its options as separate toggles', () => {
    const view = buildRibbonModel('ja', { profile: 'excel365Mac' }).find(
      (tab) => tab.id === 'view',
    );
    const show = view?.groups.find((g) => g.commands.some((c) => c.id === 'mac.view.gridlines'));
    expect((show?.commands ?? []).map((c) => c.label)).toEqual([
      'ナビゲーション',
      'ルーラー',
      '枠線',
      '数式バー',
      '見出し',
      'データ型アイコン',
      'セルにフォーカス',
    ]);
    expect(show?.commands.find((c) => c.id === 'mac.view.gridlines')?.className).toContain(
      'mac-checkbox',
    );
  });
});
