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
});
