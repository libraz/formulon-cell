import { describe, expect, it } from 'vitest';
import { createExcelRibbonSvg } from '../../../src/toolbar/excel-ribbon-icons.js';
import { RIBBON_GLYPHS } from '../../../src/toolbar/icon-kit/glyphs/index.js';
import { MAC_ICON_KEYS, MAC_ICONS } from '../../../src/toolbar/icon-kit/glyphs/mac.js';
import { PALETTE } from '../../../src/toolbar/icon-kit/tokens.js';
import type { IconSegment } from '../../../src/toolbar/icon-kit/types.js';
import { ribbonDisplayText, toolbarMenuText } from '../../../src/toolbar/menu-text.js';
import { ribbonActivationForCommand } from '../../../src/toolbar/ribbon/activation.js';
import { createRenderRibbon } from '../../../src/toolbar/ribbon/render-ribbon.js';
import {
  buildRibbonModel,
  EXCEL365_MAC_RIBBON_TABS,
  toolbarText,
} from '../../../src/toolbar/ribbon-model.js';

const functionFamilyCommands = [
  ['mac.formulas.recent', 'functionRecent', '最近使ったもの'],
  ['mac.formulas.financial', 'functionFinancial', '財務'],
  ['mac.formulas.logical', 'functionLogical', '論理'],
  ['mac.formulas.text', 'functionText', '文字列操作'],
  ['mac.formulas.dateTime', 'functionDateTime', '日付/時刻'],
  ['mac.formulas.lookup', 'functionLookup', '検索/行列'],
  ['mac.formulas.math', 'functionMath', '数学/三角'],
  ['mac.formulas.more', 'functionMore', 'その他の関数'],
] as const;

const segmentGeometry = (name: string): string =>
  ((RIBBON_GLYPHS[name as keyof typeof RIBBON_GLYPHS] ?? []) as readonly IconSegment[])
    .map((segment) => `${segment.d}|${segment.transform ?? ''}`)
    .join('||');

describe('Mac ribbon icon coverage', () => {
  it('registers every Mac-only icon in the shared ribbon registry', () => {
    expect(MAC_ICON_KEYS).toHaveLength(36);
    for (const key of MAC_ICON_KEYS) {
      expect(MAC_ICONS[key], key).toBeDefined();
      expect(RIBBON_GLYPHS[key], key).toBe(MAC_ICONS[key]);
    }
    expect(RIBBON_GLYPHS.chart).toBeDefined();
    expect(RIBBON_GLYPHS.table).toBeDefined();
    expect(RIBBON_GLYPHS.function).toBeDefined();
  });

  it('uses non-empty vector segments with the closed icon-kit palette', () => {
    for (const key of MAC_ICON_KEYS) {
      const segments = MAC_ICONS[key];
      expect(segments.length, key).toBeGreaterThan(0);
      for (const segment of segments) {
        expect(segment.d, key).toMatch(/[A-Za-z]/);
        expect(segment.d, key).not.toContain('<');
        if (segment.fill && segment.fill !== 'none') {
          expect(Object.values(PALETTE), key).toContain(segment.fill);
        }
        if (segment.stroke) expect(Object.values(PALETTE), key).toContain(segment.stroke);
      }
    }
  });

  it('resolves every icon declared by the Excel 365 Mac ribbon model', () => {
    const commands = buildRibbonModel('en', { profile: 'excel365Mac' })
      .flatMap((tab) => tab.groups)
      .flatMap((group) => group.commands);
    const iconNames = new Set(
      commands
        .map((command) => command.icon)
        .filter((icon): icon is string => typeof icon === 'string' && icon.length > 0),
    );

    expect(iconNames.size).toBeGreaterThan(MAC_ICON_KEYS.length);
    for (const icon of iconNames) {
      expect(RIBBON_GLYPHS[icon as keyof typeof RIBBON_GLYPHS], icon).toBeDefined();
      expect(RIBBON_GLYPHS[icon as keyof typeof RIBBON_GLYPHS]?.length, icon).toBeGreaterThan(0);
    }
  });

  it('maps the eight Function Library families to native icons and Japanese labels', () => {
    const formulas = buildRibbonModel('ja', { profile: 'excel365Mac' }).find(
      (tab) => tab.id === 'formulas',
    );
    expect(formulas).toBeDefined();
    const commands = new Map(
      formulas?.groups.flatMap((group) => group.commands).map((c) => [c.id, c]),
    );
    for (const [id, icon, label] of functionFamilyCommands) {
      const command = commands.get(id);
      expect(command, id).toMatchObject({ title: label, label, icon });
      expect(ribbonActivationForCommand(id), id).toEqual({
        kind: 'dropdown',
        menuId: expect.stringMatching(/^menu-mac-/),
      });
    }
  });

  it('keeps family glyphs distinct through geometry and the built-in SVG route', () => {
    const iconNames = functionFamilyCommands.map(([, icon]) => icon);
    expect(new Set(iconNames).size).toBe(8);
    const geometries = iconNames.map(segmentGeometry);
    expect(new Set(geometries).size).toBe(8);
    for (const icon of iconNames) {
      const definition = RIBBON_GLYPHS[icon as keyof typeof RIBBON_GLYPHS];
      expect(definition, icon).toBeDefined();
      expect(definition?.length, icon).toBeGreaterThan(0);
      const svg = createExcelRibbonSvg(icon, 'test-icon');
      expect(svg, icon).not.toBeNull();
      expect(svg?.getAttribute('viewBox'), icon).toBe('0 0 24 24');
      expect(svg?.getAttribute('focusable'), icon).toBe('false');
      expect(svg?.getAttribute('aria-hidden'), icon).toBe('true');
      expect(svg?.querySelectorAll('path').length, icon).toBeGreaterThan(0);
    }
  });

  it('renders localized family names, decorative SVGs, and the function-library group', () => {
    const ribbonRoot = document.createElement('div');
    document.body.append(ribbonRoot);
    const makeSvg = (viewBox: string, pathData: string, className: string): SVGSVGElement => {
      const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
      svg.setAttribute('viewBox', viewBox);
      svg.classList.add(className);
      const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
      path.setAttribute('d', pathData);
      svg.append(path);
      return svg;
    };
    const renderer = createRenderRibbon({
      getInst: () => null,
      ribbonLang: 'ja',
      ribbonText: toolbarText('ja'),
      ribbonMenuText: toolbarMenuText('ja'),
      ribbonDisplayOptionsText: ribbonDisplayText('ja'),
      ribbonTabs: EXCEL365_MAC_RIBBON_TABS,
      profile: 'excel365Mac',
      ribbonRoot,
      state: {
        getActiveTab: () => 'formulas',
        getCollapsed: () => false,
        getDisplayMode: () => 'full',
        getAutoHidePeek: () => false,
        getBackstageOpen: () => false,
        getDisplayMenuOpen: () => false,
        getFormulaBarVisible: () => true,
      },
      helpers: {
        createSelect: () => document.createElement('div'),
        createColor: () => document.createElement('div'),
        createIcon: (name) => createExcelRibbonSvg(name, 'fc-tb__rb-icon'),
        makeSvg,
        chevronPath: 'M0 0L6 6L12 0',
      },
      createBackstageView: () => document.createElement('div'),
      projectFormatToolbar: () => undefined,
    });
    renderer.renderRibbon();
    const panel = ribbonRoot.querySelector<HTMLElement>('[data-ribbon-panel="formulas"]');
    expect(panel).not.toBeNull();
    for (const [id, , label] of functionFamilyCommands) {
      const button = panel?.querySelector<HTMLButtonElement>(`[data-ribbon-command="${id}"]`);
      expect(button, id).not.toBeNull();
      expect(button?.getAttribute('aria-label'), id).toBe(label);
      expect(button?.getAttribute('aria-haspopup'), id).toBe('menu');
      expect(button?.querySelector('svg[aria-hidden="true"]'), id).not.toBeNull();
    }
    const group = panel?.querySelector<HTMLElement>('.fc-tb__ribbon-group--function-library');
    expect(group).not.toBeNull();
    expect(group?.classList.contains('fc-tb__ribbon-group--compact')).toBe(false);
    expect(group?.classList.contains('fc-tb__ribbon-group--gallery')).toBe(false);
    const tools = group?.querySelector<HTMLElement>('.fc-tb__ribbon-tools');
    expect(
      Array.from(tools?.children ?? [])
        .filter((child): child is HTMLElement => child instanceof HTMLElement)
        .filter((child) => child.dataset.ribbonCommand !== undefined)
        .map((child) => child.dataset.ribbonCommand),
    ).toEqual([
      'mac.formulas.insertFunction',
      'mac.formulas.autoSum',
      ...functionFamilyCommands.map(([id]) => id),
    ]);
    ribbonRoot.remove();
  });
});
