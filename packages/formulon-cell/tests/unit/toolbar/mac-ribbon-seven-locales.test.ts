import { describe, expect, it } from 'vitest';
import { createSpreadsheetStore } from '../../../src/store/store.js';
import { MAC_AUTOMATION_GALLERY } from '../../../src/toolbar/ribbon/mac/automation-gallery.js';
import {
  GROUP_LABELS,
  labels,
  macRibbonLabelForCommand,
} from '../../../src/toolbar/ribbon/mac/labels.js';
import { MAC_LOCALES, type MacLocale } from '../../../src/toolbar/ribbon/mac/labels-data.js';
import {
  createMacRibbonMenuFactory,
  projectMacFunctionCategoryMenus,
} from '../../../src/toolbar/ribbon/mac/menus.js';
import { MAC_RIBBON_MENU_ITEMS } from '../../../src/toolbar/ribbon/mac/model.js';
import { buildRibbonModel } from '../../../src/toolbar/ribbon-model.js';

const chartColumnLabels: Readonly<Record<MacLocale, string>> = {
  ja: '縦棒',
  en: 'Column',
  de: 'Säule',
  fr: 'Histogramme',
  zh: '柱形图',
  ko: '세로 막대형',
  th: 'คอลัมน์',
};

const pageSizeLabels: Readonly<Record<MacLocale, readonly string[]>> = {
  ja: ['A4', 'A3', 'レター', 'リーガル'],
  en: ['A4', 'A3', 'Letter', 'Legal'],
  de: ['A4', 'A3', 'Letter', 'Legal'],
  fr: ['A4', 'A3', 'Lettre', 'Juridique'],
  zh: ['A4', 'A3', '信纸', '法律'],
  ko: ['A4', 'A3', 'Letter', 'Legal'],
  th: ['A4', 'A3', 'Letter', 'Legal'],
};

describe('Mac ribbon seven-language localization', () => {
  it('matches the platform-specific tabs captured from Excel 16.113.4', () => {
    const french = buildRibbonModel('fr', { profile: 'excel365Mac' });
    const thai = buildRibbonModel('th', { profile: 'excel365Mac' });
    expect(french.find((tab) => tab.id === 'insert')?.label).toBe('Insérer');
    expect(french.find((tab) => tab.id === 'automate')?.label).toBe('Automatisation');
    expect(thai.find((tab) => tab.id === 'automate')?.label).toBe('อัตโนมัติ');
    expect(buildRibbonModel('fr').find((tab) => tab.id === 'insert')?.label).toBe('Insertion');
  });

  it('provides complete localized labels for the model and command dictionaries', () => {
    for (const locale of MAC_LOCALES) {
      const model = buildRibbonModel(locale, { profile: 'excel365Mac' });
      expect(model).toHaveLength(9);
      expect(macRibbonLabelForCommand('mac.insert.chartColumn', locale)).toBe(
        chartColumnLabels[locale],
      );

      for (const tab of model) {
        expect(tab.label.length, `${locale}:${tab.id}`).toBeGreaterThan(0);
        for (const group of tab.groups) {
          expect(group.title.length, `${locale}:${tab.id}:${group.title}`).toBeGreaterThan(0);
          for (const command of group.commands) {
            if (command.kind !== 'break') {
              expect(command.title.length, `${locale}:${command.id}:title`).toBeGreaterThan(0);
            }
            for (const option of command.options ?? []) {
              expect(
                option.label.length,
                `${locale}:${command.id}:${option.value}`,
              ).toBeGreaterThan(0);
            }
          }
        }
      }
      const automate = model.find((tab) => tab.id === 'automate');
      for (const script of MAC_AUTOMATION_GALLERY) {
        const command = automate?.groups
          .flatMap((group) => group.commands)
          .find((candidate) => candidate.id === script.commandId);
        expect(command?.label, `${locale}:${script.commandId}`).toBe(script.labels[locale]);
        expect(command?.title, `${locale}:${script.commandId}:title`).toBe(script.labels[locale]);
      }
    }

    for (const [key, label] of Object.entries(labels)) {
      expect(Object.keys(label).sort(), key).toEqual([...MAC_LOCALES].sort());
      for (const locale of MAC_LOCALES) expect(label[locale], `${key}:${locale}`).toBeTruthy();
    }
    for (const [key, label] of Object.entries(GROUP_LABELS)) {
      expect(Object.keys(label).sort(), key).toEqual([...MAC_LOCALES].sort());
    }
  });

  it('localizes inline Mac menu leaves instead of retaining manifest English', () => {
    const staticMenuItems = Object.values(MAC_RIBBON_MENU_ITEMS)
      .flat()
      .filter((item) => !item.id.startsWith('mac.function.'));
    for (const locale of MAC_LOCALES) {
      for (const item of staticMenuItems) {
        expect(macRibbonLabelForCommand(item.id, locale), `${item.id}:${locale}`).not.toBe(item.id);
      }
    }
    for (const locale of MAC_LOCALES) {
      const menu = createMacRibbonMenuFactory(locale)('mac.page.size');
      if (!menu) throw new Error('Page size menu was not created.');
      expect(
        Array.from(menu.querySelectorAll<HTMLElement>('[data-ribbon-command]')).map(
          (button) => button.textContent,
        ),
        locale,
      ).toEqual(pageSizeLabels[locale]);
    }
    expect(
      createMacRibbonMenuFactory('de')('mac.formulas.autoSum')?.querySelector(
        '[data-ribbon-command="mac.autosum.SUM"]',
      )?.textContent,
    ).toBe('SUMME');
    expect(
      createMacRibbonMenuFactory('fr')('mac.formulas.autoSum')?.querySelector(
        '[data-ribbon-command="mac.autosum.SUM"]',
      )?.textContent,
    ).toBe('SOMME');
  });

  it('uses German and French native function metadata while retaining canonical command ids', () => {
    const store = createSpreadsheetStore();
    const calls: Array<{ name: string; locale: number }> = [];
    const reader = {
      functionNames: () => ['SUM'] as const,
      functionMetadata: (name: string, locale: 0 | 1 | 2 | 3) => {
        calls.push({ name, locale });
        return {
          name,
          minArity: 1,
          maxArity: 1,
          localizedName: locale === 2 ? 'SUMME' : locale === 3 ? 'SOMME' : name,
          availability: 0,
        };
      },
      canonicalizeFunctionName: (name: string) =>
        name === 'SUMME' || name === 'SOMME' ? 'SUM' : name,
    };

    for (const [locale, displayName, ordinal] of [
      ['de', 'SUMME', 2],
      ['fr', 'SOMME', 3],
    ] as const) {
      const menu = createMacRibbonMenuFactory(locale)('mac.formulas.math');
      if (!menu) throw new Error(`Math menu missing for ${locale}.`);
      const host = document.createElement('div');
      host.append(menu);
      projectMacFunctionCategoryMenus(
        host,
        store,
        locale,
        new Set([locale === 'de' ? 'SUMME' : 'SOMME']),
        { reader },
      );
      const leaf = menu.querySelector<HTMLButtonElement>(
        '[data-ribbon-command="mac.function.SUM"]',
      );
      expect(leaf?.textContent, locale).toBe(displayName);
      expect(leaf?.dataset.ribbonCommand).toBe('mac.function.SUM');
      expect(calls.some((call) => call.name === 'SUM' && call.locale === ordinal)).toBe(true);
    }
  });

  it('uses tagged provider metadata for every function tile without widening native ordinals', () => {
    const store = createSpreadsheetStore();
    const localeTags: Readonly<Record<MacLocale, string>> = {
      en: 'en-US',
      ja: 'ja-JP',
      de: 'de-DE',
      fr: 'fr-FR',
      zh: 'zh-CN',
      ko: 'ko-KR',
      th: 'th-TH',
    };
    const providerCalls: Array<{ name: string; locale: string }> = [];
    const nativeCalls: Array<{ name: string; locale: 0 | 1 | 2 | 3 }> = [];
    const canonicalCalls: Array<{ name: string; locale: 0 | 1 | 2 | 3 }> = [];
    const reader = {
      functionNames: () => ['SUM'] as const,
      functionMetadata: (name: string, locale: 0 | 1 | 2 | 3) => {
        nativeCalls.push({ name, locale });
        return {
          name,
          minArity: 1,
          maxArity: 1,
          localizedName: name,
          availability: 0,
        };
      },
      functionMetadataForLocale: (name: string, locale: string) => {
        providerCalls.push({ name, locale });
        return {
          name,
          minArity: 1,
          maxArity: 1,
          localizedName: `${name} (${locale})`,
          description: `sum-${locale}`,
          availability: 0,
        };
      },
      canonicalizeFunctionName: (name: string, locale: 0 | 1 | 2 | 3) => {
        canonicalCalls.push({ name, locale });
        return name;
      },
    };

    for (const locale of MAC_LOCALES) {
      const menu = createMacRibbonMenuFactory(locale)('mac.formulas.math');
      if (!menu) throw new Error(`Math menu missing for ${locale}.`);
      const host = document.createElement('div');
      host.append(menu);
      projectMacFunctionCategoryMenus(host, store, locale, new Set(['SUM']), { reader });
      const leaf = menu.querySelector<HTMLButtonElement>(
        '[data-ribbon-command="mac.function.SUM"]',
      );
      expect(leaf?.textContent, locale).toBe(`SUM (${localeTags[locale]})`);
      expect(leaf?.disabled, locale).toBe(false);
      expect(
        providerCalls.filter((call) => call.locale === localeTags[locale]).length,
      ).toBeGreaterThan(0);
    }

    expect(nativeCalls).toHaveLength(0);
    expect(canonicalCalls.map((call) => call.locale)).toEqual(
      MAC_LOCALES.map((locale) =>
        locale === 'ja' ? 1 : locale === 'de' ? 2 : locale === 'fr' ? 3 : 0,
      ),
    );
  });

  it('ships visible automation copy for every locale', () => {
    for (const script of MAC_AUTOMATION_GALLERY) {
      for (const locale of MAC_LOCALES) {
        expect(script.labels[locale], `${script.id}:label:${locale}`).toBeTruthy();
        expect(script.details[locale], `${script.id}:detail:${locale}`).toBeTruthy();
      }
    }
  });
});
