import { describe, expect, it } from 'vitest';
import { MAC_OBJECT_TAB_DISABLED_COMMAND_IDS } from '../../../src/toolbar/ribbon/mac/object-tab-ids.js';
import {
  buildRibbonModel,
  CHART_CONTEXTUAL_RIBBON_TABS,
  CONTEXTUAL_RIBBON_TABS,
  type RibbonCommand,
} from '../../../src/toolbar/ribbon-model.js';

const macTab = (id: string) =>
  buildRibbonModel('ja', { profile: 'excel365Mac', contextualTabs: CONTEXTUAL_RIBBON_TABS }).find(
    (tab) => tab.id === id,
  );
const labels = (id: string): string[][] =>
  (macTab(id)?.groups ?? []).map((group) => group.commands.map((c: RibbonCommand) => c.label));

// Excel for Mac (ja): the table tab and the two chart tabs.
describe('Mac table and chart tabs', () => {
  it('names the table tab テーブル and groups it as Excel does', () => {
    expect(macTab('tableDesign')?.label).toBe('テーブル');
    expect(labels('tableDesign')).toEqual([
      ['テーブル名:'],
      ['ピボットテーブルで集計', '重複を削除', '範囲に変換', 'スライサーの挿入'],
      ['更新'],
      [
        '見出し行',
        '集計行',
        '縞模様 (行)',
        '最初の列',
        '最後の列',
        '縞模様 (列)',
        'フィルター ボタン',
      ],
      ['クイック スタイル'],
    ]);
  });

  it('names the PivotTable design tab デザイン and groups it as Excel does', () => {
    expect(macTab('pivotDesign')?.label).toBe('デザイン');
    expect(labels('pivotDesign').map((group) => group.length)).toEqual([4, 4, 1]);
  });

  it('lists the chart design tab', () => {
    expect(labels('chartDesign')).toEqual([
      ['グラフ要素を追加', 'クイック レイアウト'],
      ['色の変更', 'クイック スタイル'],
      ['行/列の切り替え', 'グラフ データの選択'],
      ['グラフの種類の変更'],
      ['グラフの移動'],
    ]);
  });

  it('appends the two chart tabs after Automate for a selected chart', () => {
    const strip = buildRibbonModel('ja', {
      profile: 'excel365Mac',
      contextualTabs: CHART_CONTEXTUAL_RIBBON_TABS,
    }).map((tab) => tab.label);
    expect(strip).toEqual([
      'ホーム',
      '挿入',
      '描画',
      'ページ レイアウト',
      '数式',
      'データ',
      '校閲',
      '表示',
      '自動化',
      'グラフのデザイン',
      '書式',
    ]);
  });

  it('keeps the commands without a host feature disabled and runs the rest', () => {
    const commands = ['tableDesign', 'chartDesign', 'chartFormat'].flatMap(
      (id) => macTab(id)?.groups.flatMap((group) => group.commands) ?? [],
    );
    for (const command of commands) {
      expect(command.disabled === true, command.id).toBe(
        MAC_OBJECT_TAB_DISABLED_COMMAND_IDS.includes(command.id),
      );
    }
    const ids = commands.map((c) => c.id);
    expect(ids).toContain('tableHeaderRow');
    expect(ids).toContain('chartAddElement');
    expect(ids).toContain('chartFormatSelection');
  });

  it('keeps the shared tabs on the default profile', () => {
    const shared = buildRibbonModel('ja', { contextualTabs: CONTEXTUAL_RIBBON_TABS }).find(
      (tab) => tab.id === 'tableDesign',
    );
    expect(shared?.label).toBe('テーブル デザイン');
  });
});
