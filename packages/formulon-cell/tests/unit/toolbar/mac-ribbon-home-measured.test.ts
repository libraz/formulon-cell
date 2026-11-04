import { describe, expect, it } from 'vitest';
import { MAC_RIBBON_DISABLED_COMMAND_IDS } from '../../../src/toolbar/ribbon/mac/model.js';
import { buildRibbonModel } from '../../../src/toolbar/ribbon-model.js';

const homeCommands = () =>
  buildRibbonModel('ja', { profile: 'excel365Mac' })
    .find((tab) => tab.id === 'home')
    ?.groups.flatMap((g) => g.commands) ?? [];
const byId = (id: string) => homeCommands().find((c) => c.id === id);

// Excel for Mac (ja) Home tab: control names and AX help strings.
describe('Mac Home ribbon labels', () => {
  it.each([
    ['textOrientation', '文字の向き'],
    ['indentDecrease', 'インデント解除'],
    ['indentIncrease', 'インデント'],
    ['borders', '線種/網かけ変更'],
    ['fillColor', '網かけ'],
    ['decUp', '小数点表示桁上げ'],
    ['decDown', '小数点表示桁下げ'],
    ['insertRows', '挿入'],
    ['deleteRows', '削除'],
  ])('names %s %s', (id, name) => {
    expect(byId(id)?.label).toBe(name);
    expect(byId(id)?.title).toBe(name);
  });
});

describe('Mac Home ribbon tooltips', () => {
  it('carries no shortcut text in any tooltip', () => {
    for (const command of homeCommands()) {
      expect(command.title, command.id).not.toMatch(/\((?:Cmd|Ctrl|⌘)[^)]*\)|\(Σ\)/);
    }
  });

  it('uses the control name, except the two commands with their own help text', () => {
    expect(byId('bold')?.title).toBe('太字');
    expect(byId('findHome')?.title).toBe('検索と選択');
    expect(byId('autosum')?.title).toBe('オート SUM');
    expect(byId('formatPainter')?.title).toBe('書式を別の場所にコピーして適用');
    expect(byId('alignC')?.title).toBe('文字列中央揃え');
  });
});

describe('Mac Home add-in button', () => {
  it('is enabled', () => {
    expect(byId('mac.home.addins')?.disabled).toBeFalsy();
    expect(MAC_RIBBON_DISABLED_COMMAND_IDS.has('mac.home.addins')).toBe(false);
  });

  describe('Mac Home dropdown lists', () => {
    it('lists font sizes from 6 as the desktop list does', () => {
      expect(byId('fontSize')?.options?.map((o) => o.value)).toEqual([
        '6',
        '8',
        '9',
        '10',
        '11',
        '12',
        '14',
        '16',
        '18',
        '20',
        '22',
        '24',
        '26',
        '28',
        '36',
        '48',
        '72',
      ]);
    });

    it('names the percent number format as the desktop does', () => {
      const percent = byId('numberFormat')?.options?.find((o) => o.value === 'percent');
      expect(percent?.label).toBe('パーセンテージ');
    });
  });
});
