import { describe, expect, it } from 'vitest';
import { dictionaries } from '../../../src/i18n/strings.js';
import {
  fontFamilyCatalog,
  fontFamilyLabel,
  fontFamilyOptions,
  fontFamilyValueFromLabel,
} from '../../../src/toolbar/ribbon/font-family-options.js';
import { createSelectColorRibbon } from '../../../src/toolbar/ribbon/select-color.js';
import { buildRibbonModel, FONT_FAMILIES } from '../../../src/toolbar/ribbon-model.js';

describe('toolbar/ribbon/font-family-options', () => {
  it('keeps the Japanese theme body value separate from its display label', () => {
    const option = fontFamilyCatalog('ja').find((candidate) => candidate.value === '游ゴシック');

    expect(option).toMatchObject({
      value: '游ゴシック',
      label: '游ゴシック Regular',
      section: 'theme',
      role: 'body',
    });
    expect(option?.probeNames).toContain('Yu Gothic');
    expect(new Set(fontFamilyCatalog('ja').map((candidate) => candidate.value)).size).toBe(
      fontFamilyCatalog('ja').length,
    );
  });

  it('retains an unavailable current family exactly and injects it only once', () => {
    const current = 'Imported Japanese Font';
    const options = fontFamilyOptions('ja', current);
    const matching = options.filter((candidate) => candidate.value === current);

    expect(matching).toHaveLength(1);
    expect(matching[0]).toMatchObject({ value: current, label: current, section: 'all' });
  });

  it('keeps the legacy Japanese spelling untouched until the user chooses the canonical row', () => {
    const options = fontFamilyOptions('ja', '游ゴシック Regular');
    const duplicateLabels = options.filter((candidate) => candidate.label === '游ゴシック Regular');

    expect(options.map((candidate) => candidate.value)).toContain('游ゴシック');
    expect(options.map((candidate) => candidate.value)).toContain('游ゴシック Regular');
    expect(duplicateLabels).toEqual([
      expect.objectContaining({ value: '游ゴシック Regular', label: '游ゴシック Regular' }),
    ]);
    expect(fontFamilyLabel('ja', '游ゴシック Regular')).toBe('游ゴシック Regular');
    expect(fontFamilyValueFromLabel('ja', '游ゴシック Regular')).toBe('游ゴシック Regular');
  });

  it('uses canonical full-width MS Japanese values', () => {
    const values = fontFamilyCatalog('ja').map((candidate) => candidate.value);
    expect(values).toContain('ＭＳ Ｐゴシック');
    expect(values).toContain('ＭＳ Ｐ明朝');
    expect(values).not.toContain('MS Pゴシック');
    expect(values).not.toContain('MS P明朝');
  });

  it('sorts the all section with the active locale collator', () => {
    const options = fontFamilyOptions('ja');
    const all = options.filter((candidate) => candidate.section === 'all');
    const expected = [...all].sort((left, right) =>
      new Intl.Collator('ja', { numeric: true, sensitivity: 'base' }).compare(
        left.label,
        right.label,
      ),
    );
    expect(all.map((candidate) => candidate.value)).toEqual(
      expected.map((candidate) => candidate.value),
    );
  });

  it('does not add Japanese-only catalog entries to other locales', () => {
    const english = fontFamilyOptions('en');
    expect(english.some((candidate) => /[\u3000-\u9fff]/.test(candidate.value))).toBe(false);

    const imported = fontFamilyOptions('en', '游ゴシック');
    expect(imported.map((candidate) => candidate.value)).toContain('游ゴシック');
  });

  it('uses canonical values and localized role text in the ribbon picker', () => {
    const command = buildRibbonModel('ja')
      .find((tab) => tab.id === 'home')
      ?.groups.find((group) => group.variant === 'font')
      ?.commands.find((candidate) => candidate.id === 'fontFamily');
    if (!command) throw new Error('font family ribbon command missing');
    const applied: string[] = [];
    const ribbon = createSelectColorRibbon({
      ribbonLang: 'ja',
      ribbonText: dictionaries.ja.ribbon,
      pageScaleText: { automatic: '自動', page: 'ページ', pages: 'ページ' },
      getInst: () => null,
      applyRibbonControl: (_id, value) => applied.push(value),
      currentRibbonControlValue: () => '游ゴシック',
      shouldShowFontOption: () => true,
      createRibbonIcon: () => null,
    });
    const select = ribbon.createRibbonSelect(command);
    document.body.appendChild(select);
    try {
      expect(select.querySelector('.fc-tb__rb-dd__value')?.textContent).toContain(
        '游ゴシック Regular',
      );
      select.querySelector<HTMLButtonElement>('.fc-tb__rb-dd__btn')?.click();
      const bodyOption = document.querySelector<HTMLButtonElement>(
        '.fc-tb__rb-dd__opt[data-value="游ゴシック"]',
      );
      expect(bodyOption?.textContent).toContain('(本文)');
      expect(bodyOption?.title).toBe('游ゴシック Regular (本文)');
      expect(bodyOption?.getAttribute('aria-label')).toBe('游ゴシック Regular (本文)');
      expect(bodyOption?.querySelector('.fc-tb__rb-dd__submenu')).toBeNull();
      bodyOption?.click();
      expect(applied).toContain('游ゴシック');
    } finally {
      select.remove();
      document.querySelector('.fc-tb__rb-dd__list')?.remove();
    }
  });

  it('passes a legacy ribbon row value explicitly instead of remapping its label', () => {
    const command = buildRibbonModel('ja')
      .find((tab) => tab.id === 'home')
      ?.groups.find((group) => group.variant === 'font')
      ?.commands.find((candidate) => candidate.id === 'fontFamily');
    if (!command) throw new Error('font family ribbon command missing');
    const applied: string[] = [];
    const ribbon = createSelectColorRibbon({
      ribbonLang: 'ja',
      ribbonText: dictionaries.ja.ribbon,
      pageScaleText: { automatic: '自動', page: 'ページ', pages: 'ページ' },
      getInst: () => null,
      applyRibbonControl: (_id, value) => applied.push(value),
      currentRibbonControlValue: () => '游ゴシック Regular',
      shouldShowFontOption: () => true,
      createRibbonIcon: () => null,
    });
    const select = ribbon.createRibbonSelect(command);
    document.body.appendChild(select);
    try {
      select.querySelector<HTMLButtonElement>('.fc-tb__rb-dd__btn')?.click();
      const legacy = document.querySelector<HTMLButtonElement>(
        '.fc-tb__rb-dd__opt[data-value="游ゴシック Regular"]',
      );
      const canonical = document.querySelector<HTMLButtonElement>(
        '.fc-tb__rb-dd__opt[data-value="游ゴシック"]',
      );
      expect(legacy?.textContent).not.toContain('(本文)');
      expect(canonical?.textContent).toContain('(本文)');
      legacy?.click();
      expect(applied).toContain('游ゴシック Regular');
    } finally {
      select.remove();
      document.querySelector('.fc-tb__rb-dd__list')?.remove();
    }
  });

  it('preserves custom command candidates instead of replacing them with dictionary entries', () => {
    const custom = structuredClone(dictionaries.ja);
    custom.ribbon.fontFamilies = ['Custom Corporate Font'];
    const command = buildRibbonModel(custom)
      .find((tab) => tab.id === 'home')
      ?.groups.flatMap((group) => group.commands)
      .find((candidate) => candidate.id === 'fontFamily');
    if (!command) throw new Error('custom font family ribbon command missing');
    expect(command.options?.map((option) => option.value)).toEqual(['Custom Corporate Font']);

    const ribbon = createSelectColorRibbon({
      ribbonLang: 'ja',
      ribbonText: custom.ribbon,
      pageScaleText: { automatic: '自動', page: 'ページ', pages: 'ページ' },
      getInst: () => null,
      applyRibbonControl: () => {},
      currentRibbonControlValue: () => '游ゴシック',
      shouldShowFontOption: () => true,
      createRibbonIcon: () => null,
    });
    const select = ribbon.createRibbonSelect(command);
    document.body.appendChild(select);
    try {
      select.querySelector<HTMLButtonElement>('.fc-tb__rb-dd__btn')?.click();
      expect(
        document.querySelector('.fc-tb__rb-dd__opt[data-value="Custom Corporate Font"]'),
      ).not.toBeNull();
    } finally {
      select.remove();
      document.querySelector('.fc-tb__rb-dd__list')?.remove();
    }
  });

  it('keeps the public compatibility tuple unchanged', () => {
    expect(FONT_FAMILIES).toEqual([
      'Aptos',
      'Aptos Display',
      'Aptos Narrow',
      'Calibri',
      'Arial',
      'Segoe UI',
      '游ゴシック Light',
      '游ゴシック Regular',
      'Yu Gothic UI',
      'BIZ UDGothic',
      'BIZ UDMincho',
      'Times New Roman',
      'Consolas',
      'Georgia',
      'Tahoma',
      'Verdana',
    ]);
  });
});
