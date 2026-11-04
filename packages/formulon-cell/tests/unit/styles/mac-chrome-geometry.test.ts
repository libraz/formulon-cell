import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');
const css = readFileSync(join(root, 'src/styles/platform-mac.css'), 'utf8');
const block = css.slice(css.indexOf('/* Mac chrome geometry */'));

/** Declarations of the first rule in the Mac chrome block whose selector ends with `suffix`. */
const decls = (suffix: string): Record<string, string> => {
  const re = /([^{}]+)\{([^{}]*)\}/g;
  for (let m = re.exec(block); m; m = re.exec(block)) {
    const selectors = (m[1] ?? '').split(',').map((x) => x.trim());
    if (selectors.some((x) => x.endsWith(suffix) && !x.includes(`${suffix}:`))) {
      return Object.fromEntries(
        (m[2] ?? '')
          .split(';')
          .map((d) => d.split(/:(.+)/).map((x) => x.trim()))
          .filter((d) => d[0] && d[1]),
      );
    }
  }
  throw new Error(`no Mac chrome rule for ${suffix}`);
};

// Excel for Mac, ja, 100%: pt values from the AX frames.
describe('Mac formula bar geometry', () => {
  it('uses 26 pt action buttons and a 40 pt fx button', () => {
    expect(block).toContain('--fc-formulabar-action-size: 26px');
    expect(block).toContain('--fc-formulabar-fx-width: 40px');
    expect(decls('.fc-host__formulabar-action').height).toBe('26px');
    expect(decls('.fc-host__formulabar-fx').height).toBe('26px');
  });
});

describe('Mac sheet tab geometry', () => {
  it('sizes the tab to its name, 22 pt tall, with rounded bottom corners', () => {
    const tab = decls('.fc-host__sheetbar-tab');
    expect(tab.width).toBe('auto');
    expect(tab['min-width']).toBe('0');
    expect(tab['max-width']).toBe('none');
    expect(tab.height).toBe('22px');
    expect(tab['border-radius']).toBe('0 0 6px 6px');
  });

  it('draws the add button as a bare glyph followed by a hairline', () => {
    const add = decls('.fc-host__sheetbar-add');
    expect(add.border).toBe('0');
    expect(add['border-radius']).toBe('0');
    expect(add['border-right']).toContain('solid');
  });
});

describe('Mac status bar geometry', () => {
  it('hides the status dot and pads each aggregate item', () => {
    expect(decls('.fc-host__statusbar-dot').display).toBe('none');
    expect(decls('.fc-host__statusbar-agg').padding).toBe('0 8px');
  });

  it('sizes the view buttons and the selected fill', () => {
    expect(decls('.fc-host__statusbar-view').width).toBe('36px');
    expect(decls('.fc-host__statusbar-view--active').background).toBe('#b8b8b8');
  });

  it('draws the zoom controls as 14 pt glyphs around a 102 pt slider and a 54 pt label', () => {
    const btn = decls('.fc-host__statusbar-zoom-btn');
    expect(btn.width).toBe('14px');
    expect(btn.border).toBe('0');
    expect(decls('.fc-host__statusbar-zoom-slider').width).toBe('102px');
    expect(decls('.fc-host__statusbar-zoom-label')['min-width']).toBe('54px');
    expect(decls('.fc-host__statusbar-zoom').gap).toBe('3px');
  });
});
