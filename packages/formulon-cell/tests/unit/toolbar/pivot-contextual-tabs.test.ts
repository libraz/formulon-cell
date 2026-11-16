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
      ['ピボット フィールド名:', 'フィールドの設定', 'フィールドの展開', 'フィールドの折りたたみ'],
      ['グループの選択'],
      ['スライサーの挿入', 'タイムラインの挿入', 'フィルターの接続'],
      ['更新', 'データ ソースの変更'],
      ['クリア', '選択', 'ピボットテーブルの移動'],
      ['フィールド/項目/セット'],
      ['ピボットグラフ'],
      ['フィールド リスト', '+/- ボタン', 'フィールドのヘッダー'],
    ]);
  });

  it('words both tabs from the desktop bundle in every locale and profile', () => {
    const label = (lang: 'en' | 'de' | 'fr', id: string, profile?: 'excel365Mac') =>
      buildRibbonModel(lang, { profile, contextualTabs: PIVOT_CONTEXTUAL_RIBBON_TABS })
        .flatMap((tab) => tab.groups)
        .flatMap((group) => group.commands)
        .find((command) => command.id === id)?.label;
    expect(label('en', 'pivotActiveField', 'excel365Mac')).toBe('PivotField Name:');
    expect(label('en', 'pivotFieldsItemsSets')).toBe('Fields, Items, & Sets');
    expect(label('de', 'pivotFieldSettings', 'excel365Mac')).toBe('Feldeinstellungen');
    expect(label('de', 'pivotSubtotals')).toBe('Teilergebnisse');
    expect(label('fr', 'pivotName', 'excel365Mac')).toBe('Nom du tableau croisé dynamique :');
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
