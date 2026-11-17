import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { dictionaries } from '../../../src/i18n/strings.js';
import { MAC_AUTOMATION_GALLERY } from '../../../src/toolbar/ribbon/mac/automation-gallery.js';
import { MAC_LOCALES } from '../../../src/toolbar/ribbon/mac/labels-data.js';
import {
  MAC_RIBBON_MENU_ITEMS,
  type MacRibbonMenuItem,
} from '../../../src/toolbar/ribbon/mac/menu-manifest.js';
import { createMacRibbonMenuFactory } from '../../../src/toolbar/ribbon/mac/menus.js';
import { MAC_SHAPE_GALLERY } from '../../../src/toolbar/ribbon/mac/shape-gallery.js';
import { buildRibbonModel, type ToolbarLang } from '../../../src/toolbar/ribbon-model.js';

const REASON_KEYS = new Set(Object.keys(dictionaries.en.mac.ribbonReasons));

const rows = (items: readonly MacRibbonMenuItem[]): MacRibbonMenuItem[] =>
  items.flatMap((item) => [item, ...rows(item.submenu ?? [])]);

const macCommand = (lang: ToolbarLang, id: string) =>
  buildRibbonModel(lang, { profile: 'excel365Mac' })
    .flatMap((tab) => tab.groups.flatMap((group) => group.commands))
    .find((command) => command.id === id);

describe('Mac ribbon rows carry keys, not English text', () => {
  it('menu and shape gallery rows carry no label and only reason keys', () => {
    for (const item of [
      ...Object.values(MAC_RIBBON_MENU_ITEMS).flatMap(rows),
      ...MAC_SHAPE_GALLERY,
    ]) {
      expect('label' in item, item.id).toBe(false);
      if (item.disabledReason !== undefined) {
        expect(REASON_KEYS.has(item.disabledReason), item.id).toBe(true);
      }
    }
  });

  it('automation scripts carry their labels per locale only', () => {
    for (const script of MAC_AUTOMATION_GALLERY) {
      expect('label' in script, script.id).toBe(false);
      expect('detail' in script, script.id).toBe(false);
    }
  });

  it('select options carry no written text', () => {
    const source = readFileSync(join(process.cwd(), 'src/toolbar/ribbon/mac/model.ts'), 'utf8');
    expect(source).not.toMatch(/\{\s*value:[^}]*\blabel:\s*'[^']*\p{L}/u);
  });
});

describe('Mac ribbon rendered text', () => {
  it.each(MAC_LOCALES)('draws the Draw button in %s', (lang) => {
    expect(macCommand(lang, 'mac.draw.toggle')?.label).toBe(
      { en: 'Draw', ja: '描画', de: 'Zeichnen', fr: 'Dessin', ko: '그리기', th: 'วาด', zh: '绘图' }[
        lang
      ],
    );
  });

  it.each([
    [
      'en',
      ['Automatic', '1 page', '2 pages', '3 pages', 'Custom...'],
      ['Default', 'Temporary View'],
    ],
    ['ja', ['自動', '1 ページ', '2 ページ', '3 ページ', 'ユーザー設定...'], ['既定', '一時ビュー']],
    [
      'de',
      ['Automatisch', '1 Seite', '2 Seiten', '3 Seiten', 'Benutzerdefiniert...'],
      ['Standard', 'Temporäre Ansicht'],
    ],
  ] as const)('labels the scale and sheet view options in %s', (lang, scale, views) => {
    expect(macCommand(lang, 'scaleWidth')?.options?.map((o) => o.label)).toEqual(scale);
    expect(macCommand(lang, 'scaleHeight')?.options?.map((o) => o.label)).toEqual(scale);
    expect(macCommand(lang, 'sheetViewSelect')?.options?.map((o) => o.label)).toEqual(views);
  });

  it.each([
    ['mac.data.powerQuery', 'notImplemented'],
    ['mac.insert.forms', 'officeRequired'],
    ['mac.review.hideInk', 'embeddedUnsupported'],
  ] as const)('explains %s in every locale', (id, reason) => {
    for (const lang of MAC_LOCALES) {
      expect(macCommand(lang, id)?.disabledReason, lang).toBe(
        dictionaries[lang].mac.ribbonReasons[reason],
      );
    }
  });

  it('explains a greyed-out menu row in German', () => {
    const row = createMacRibbonMenuFactory('de')('mac.data.whatIf')?.querySelector<HTMLElement>(
      '[data-ribbon-command="mac.data.scenarioManager"]',
    );
    expect(row?.title).toContain(dictionaries.de.mac.ribbonReasons.notImplemented);
  });
});
