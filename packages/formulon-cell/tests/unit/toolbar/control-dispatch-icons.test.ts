import { describe, expect, it, vi } from 'vitest';
import { EXCEL_RIBBON_ICON_PATHS } from '../../../src/toolbar/excel-ribbon-icons.js';
import { pageScaleMenuText } from '../../../src/toolbar/menu-text.js';
import { createControlDispatch } from '../../../src/toolbar/ribbon/control-dispatch.js';
import { toolbarText } from '../../../src/toolbar/ribbon-model.js';

const createIcon = () =>
  createControlDispatch({
    getInst: () => null,
    ribbonLang: 'ja',
    ribbonText: toolbarText('ja'),
    pageScaleText: pageScaleMenuText('ja'),
    sheetEl: document.createElement('div'),
    focusSheet: vi.fn(),
    refreshWorkbookCells: vi.fn(),
    projectFormatToolbar: vi.fn(),
  }).createRibbonIcon;

describe('toolbar/ribbon control dispatch icons', () => {
  it('prefers the multi-color Excel ribbon SVG over the Fluent fallback', () => {
    const svg = createIcon()('fillColor');

    expect(svg?.classList.contains('fc-tb__rb-icon')).toBe(true);
    expect(svg?.getAttribute('viewBox')).toBe('0 0 24 24');
    expect(svg?.getAttribute('fill')).toBeNull();

    const paths = Array.from(svg?.querySelectorAll('path') ?? []);
    expect(paths.length).toBeGreaterThan(1);
    expect(paths.some((path) => path.hasAttribute('stroke'))).toBe(true);
    expect(paths.some((path) => path.getAttribute('fill') !== 'none')).toBe(true);
  });

  it('renders every Excel ribbon icon as a self-contained 24x24 SVG', () => {
    for (const name of Object.keys(EXCEL_RIBBON_ICON_PATHS)) {
      const svg = createIcon()(name);

      expect(svg, name).toBeTruthy();
      expect(svg?.getAttribute('viewBox'), name).toBe('0 0 24 24');
      expect(svg?.getAttribute('fill'), name).toBeNull();
      expect(svg?.getAttribute('aria-hidden'), name).toBe('true');

      const paths = Array.from(svg?.querySelectorAll('path') ?? []);
      expect(paths.length, name).toBeGreaterThan(0);
      expect(
        paths.every((path) => (path.getAttribute('d')?.length ?? 0) > 0),
        name,
      ).toBe(true);
    }
  });

  it('falls back to currentColor Fluent SVGs for non-overridden icons', () => {
    const svg = createIcon()('add');

    expect(svg?.getAttribute('fill')).toBe('currentColor');
    expect(svg?.querySelectorAll('path')).toHaveLength(1);
  });

  it('returns null for unknown icon names', () => {
    expect(createIcon()('not-a-ribbon-icon')).toBeNull();
  });
});
