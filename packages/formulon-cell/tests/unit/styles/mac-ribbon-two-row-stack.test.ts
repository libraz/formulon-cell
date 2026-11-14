import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');
const css = readFileSync(join(root, 'src/styles/platform-mac.css'), 'utf8').replace(
  /\/\*[\s\S]*?\*\//g,
  '',
);

/** Declarations of the rules whose selector list holds `selector`, later rules winning. */
const decls = (selector: string): Record<string, string> => {
  const out: Record<string, string> = {};
  for (const [, selectors = '', body = ''] of css.matchAll(/([^{}]+)\{([^{}]*)\}/g)) {
    if (!selectors.split(',').some((s) => s.replace(/\s+/g, ' ').trim() === selector)) continue;
    for (const d of body.split(';')) {
      const [name, value] = d.split(/:(.+)/).map((x) => x.trim());
      if (name && value) out[name] = value;
    }
  }
  return out;
};

const GROUP =
  '[data-fc-platform="mac"] > .fc-tb__ribbon-shell .fc-tb__ribbon-group--calculation.fc-tb__ribbon-group--mixed';

// Excel for Mac, ja, 100%: the AX frames put Calculate Now at y107 and Calculate Sheet at
// y139, each 26 pt tall, beside the Calculation Options tile at y104 (66 pt).
describe('Mac Formulas Calculation group', () => {
  it('stacks Calculate Now and Calculate Sheet in two 26 pt rows on a 32 pt pitch', () => {
    expect(decls(`${GROUP} .fc-tb__ribbon-tools`)).toMatchObject({
      'grid-template-rows': 'repeat(2, 32px)',
      'row-gap': '0',
    });
    expect(decls(`${GROUP} .fc-tb__rb--stacked`)).toMatchObject({
      height: '26px',
      'margin-top': '3px',
    });
    expect(decls(`${GROUP} .fc-tb__rb--wide:not(.fc-tb__rb--stacked)`)['grid-row']).toBe(
      '1 / span 2',
    );
  });
});

const CHARTS =
  '[data-fc-platform="mac"] > .fc-tb__ribbon-shell .fc-tb__ribbon-group--charts.fc-tb__ribbon-group--mixed';

// Excel for Mac, ja, 100%: the gallery cells are 36 pt wide in three rows; Waterfall (y107) and
// Combo (y135) are 38x26 pt beside them, the tiles at y104.
describe('Mac Insert Charts group', () => {
  it('draws 36 pt icon cells and Waterfall and Combo as a two-row stack', () => {
    expect(decls(`${CHARTS} .fc-tb__ribbon-tools`)).toMatchObject({
      'grid-template-rows': 'repeat(6, 12px)',
      'row-gap': '0',
    });
    expect(decls(`${CHARTS} .fc-tb__rb--mac-icon-only`)).toMatchObject({
      'grid-row': 'span 2',
      width: '36px',
      'min-width': '36px',
    });
    for (const id of ['mac.insert.chartWaterfall', 'mac.insert.chartCombo']) {
      expect(decls(`${CHARTS} [data-ribbon-command="${id}"]`)).toMatchObject({
        'grid-row': 'span 3',
        width: '38px',
        height: '26px',
      });
    }
    // Rows start 3 pt below the tile top, 28 pt apart: 3 and 36 - 5 = 31.
    expect(decls(`${CHARTS} [data-ribbon-command="mac.insert.chartWaterfall"]`)['margin-top']).toBe(
      '3px',
    );
    expect(decls(`${CHARTS} [data-ribbon-command="mac.insert.chartCombo"]`)['margin-top']).toBe(
      '-5px',
    );
    expect(decls(`${CHARTS} .fc-tb__rb--wide:not(.fc-tb__rb--stacked)`)['grid-row']).toBe(
      '1 / span 6',
    );
  });
});
