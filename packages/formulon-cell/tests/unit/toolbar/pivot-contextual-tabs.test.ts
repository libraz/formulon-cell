import { describe, expect, it } from 'vitest';
import {
  buildRibbonModel,
  PIVOT_CONTEXTUAL_RIBBON_TABS,
  type RibbonCommand,
} from '../../../src/toolbar/ribbon-model.js';

const mac = (lang: 'ja' | 'en' = 'ja') =>
  buildRibbonModel(lang, { profile: 'excel365Mac', contextualTabs: PIVOT_CONTEXTUAL_RIBBON_TABS });
const tabOf = (id: string, lang: 'ja' | 'en' = 'ja') => mac(lang).find((tab) => tab.id === id);
const commands = (id: string): RibbonCommand[] =>
  (tabOf(id)?.groups ?? []).flatMap((g) => g.commands).filter((c) => c.kind !== 'break');

// Excel for Mac (ja): the two tabs a PivotTable adds, after Automate.
describe('PivotTable contextual tabs', () => {
  it('adds PivotTable Analyze and Design after the standard tabs', () => {
    const tabs = mac();
    expect(tabs.slice(-2).map((tab) => [tab.id, tab.label])).toEqual([
      ['pivotAnalyze', 'ピボットテーブル分析'],
      ['pivotDesign', 'デザイン'],
    ]);
    expect(tabs).toHaveLength(11);
  });

  it('is absent from the ribbon unless asked for', () => {
    expect(buildRibbonModel('ja', { profile: 'excel365Mac' })).toHaveLength(9);
  });

  it('lays the Analyze tab out as the desktop groups', () => {
    const groups = (tabOf('pivotAnalyze')?.groups ?? []).map((g) =>
      g.commands.filter((c) => c.kind !== 'break').map((c) => c.label),
    );
    expect(groups).toEqual([
      ['ピボットテーブル名:', 'オプション'],
      ['アクティブなフィールド:', 'フィールドの設定', 'フィールドの展開', 'フィールドの折りたたみ'],
      ['グループの選択'],
      ['スライサーの挿入', 'タイムラインの挿入', 'フィルターの接続'],
      ['更新', 'データ ソースの変更'],
      ['クリア', '選択', 'ピボットテーブルの移動'],
      ['フィールド/項目/セット'],
      ['ピボットグラフ'],
      ['フィールド リスト', '+/- ボタン', 'フィールドのヘッダー'],
    ]);
  });

  it('enables the field list and leaves what the host cannot do disabled', () => {
    const byId = new Map(commands('pivotAnalyze').map((c) => [c.id, c]));
    expect(byId.get('pivotFieldList')?.disabled).toBeFalsy();
    expect(byId.get('pivotMove')?.disabled).toBeFalsy();
    expect(byId.get('pivotInsertSlicer')?.disabled).toBeFalsy();
    for (const id of ['pivotInsertTimeline', 'pivotChart']) {
      expect(byId.get(id)?.disabled, id).toBe(true);
    }
  });

  it('binds Options to the PivotTable Options command', () => {
    const options = commands('pivotAnalyze').find((c) => c.id === 'pivotOptions');
    expect(options?.disabled).toBeFalsy();
  });
});
