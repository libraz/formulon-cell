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

  it('draws the zoom controls as 14 pt glyphs around a 102 pt slider', () => {
    const btn = decls('.fc-host__statusbar-zoom-btn');
    expect(btn.width).toBe('14px');
    expect(btn.border).toBe('0');
    expect(decls('.fc-host__statusbar-zoom-slider').width).toBe('102px');
    expect(decls('.fc-host__statusbar-zoom').gap).toBe('3px');
  });

  // Each zoom button keeps its 14 px glyph box in place and widens its target to 24 px into the
  // free space beside it; the label gives up its 6 px default left padding to make room.
  it('gives each zoom button a 24 px target without moving what is drawn', () => {
    const px = (value: string | undefined): number => Number.parseFloat(value ?? '0');
    const gap = 3;
    const glyph = 14;
    const buttonPadding = 6;
    const out = decls('.fc-host__statusbar-zoom-btn[data-fc-zoom="out"]');
    const zoomIn = decls('.fc-host__statusbar-zoom-btn[data-fc-zoom="in"]');
    const label = decls('.fc-host__statusbar-zoom-label');

    expect(px(out.width)).toBeGreaterThanOrEqual(24);
    expect(px(out['margin-left']) + px(out['padding-left'])).toBe(0);
    expect(px(out.width) - px(out['padding-left'])).toBe(glyph);
    // The out button grows leftward only, into the gap after the view buttons.
    expect(px(out['margin-left'])).toBeLessThan(0);
    // Fixed lengths only, so the target is the same at every hairline (0.5 px on a 2x screen).
    for (const rule of [out, zoomIn, label]) {
      expect(Object.values(rule).join(' ')).not.toContain('--fc-hairline');
    }

    const inStart = px(zoomIn['margin-left']);
    const inEnd = inStart + px(zoomIn.width);
    const inContent = px(zoomIn.width) - px(zoomIn['padding-right']);
    expect(px(zoomIn.width)).toBeGreaterThanOrEqual(24);
    expect(inStart).toBeGreaterThan(-gap);
    expect(inStart + inContent / 2).toBe(glyph / 2);

    // Positions are relative to where the zoom-in box started before it widened.
    const labelStart = inEnd + gap + px(label['margin-left']);
    const originalLabelStart = glyph + gap;
    expect(labelStart).toBeGreaterThan(inEnd);
    expect(labelStart + px(label['min-width'])).toBe(originalLabelStart + 54);
    expect(labelStart + px(label['padding-left'])).toBe(originalLabelStart + buttonPadding);
  });

  // The drawn field is the 18.2 px line plus the padding, which grows with the hairline (0.5 px on
  // a 2x screen, 2 px in the contrast theme); the transparent bands shrink by the same amount.
  it('grows the formula bar input box to the 24 px target around an unchanged field', () => {
    const input = decls('.fc-host__formulabar-input');
    expect(input.border).toBe('solid transparent');
    expect(input['--fc-formulabar-input-band']).toBe('max(0px, calc(2px - var(--fc-hairline)))');
    expect(input['border-width']).toBe('var(--fc-formulabar-input-band) 0');
    expect(input['border-radius']).toBe('2px / calc(2px + var(--fc-formulabar-input-band))');
    expect(input.padding).toBe('calc(var(--fc-hairline) + 1px) calc(var(--fc-hairline) + 6px)');
    expect(input['background-clip']).toBe('padding-box');
    expect(input['box-shadow']).toBe(
      'inset 0 0 0 var(--fc-hairline) var(--fc-formulabar-input-frame)',
    );
    for (const hairline of [0.5, 1, 2]) {
      const band = Math.max(0, 2 - hairline);
      expect(18.2 + 2 * (hairline + 1) + 2 * band).toBeGreaterThanOrEqual(24);
    }
  });
});
