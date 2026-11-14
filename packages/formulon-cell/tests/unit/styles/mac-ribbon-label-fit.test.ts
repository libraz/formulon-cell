import { readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { ribbonActivationForCommand } from '../../../src/toolbar/ribbon/activation.js';
import {
  buildRibbonModel,
  DRAW_MIXED_LAYOUT_GROUP_VARIANTS,
  FORMULAS_MIXED_LAYOUT_GROUP_VARIANTS,
  INSERT_MIXED_LAYOUT_GROUP_VARIANTS,
  PAGE_LAYOUT_MIXED_LAYOUT_GROUP_VARIANTS,
  type RibbonCommand,
} from '../../../src/toolbar/ribbon-model.js';

const here = dirname(fileURLToPath(import.meta.url));
const css = readFileSync(
  resolve(here, '../../../src/styles/toolbar/ribbon/groups.css'),
  'utf8',
).replace(/\/\*[\s\S]*?\*\//g, '');

/** Rules with each selector of a list on its own, whitespace-normalised. */
const rules = [...css.matchAll(/([^{}]+)\{([^{}]*)\}/g)].flatMap(([, selectors = '', body = '']) =>
  selectors
    .split(',')
    .map((selector) => ({ selector: selector.replace(/\s+/g, ' ').trim(), body })),
);
const decl = (selector: string, name: string): string | undefined => {
  let value: string | undefined;
  for (const rule of rules) {
    if (rule.selector !== selector) continue;
    const hit = rule.body.match(new RegExp(`(?:^|[;\\s])${name}:\\s*([^;]+);`));
    if (hit?.[1]) value = hit[1].trim();
  }
  return value;
};
/** The first declaration of `name` in the rule for `selector`. */
const declBase = (selector: string, name: string): string | undefined => {
  for (const rule of rules) {
    if (rule.selector !== selector) continue;
    const hit = rule.body.match(new RegExp(`(?:^|[;\\s])${name}:\\s*([^;]+);`));
    if (hit?.[1]) return hit[1].trim();
  }
  return undefined;
};
const px = (value: string | undefined): number => Number.parseFloat(value ?? 'NaN');

const LABEL = 'span:not(.fc-tb__rb-icon):not(.fc-tb__rb-split-chevron)';
const TILES_LABEL = `.fc-tb__ribbon-group--tiles .fc-tb__rb > ${LABEL}`;
const MIXED_TILE_LABEL = `.fc-tb__ribbon-group--mixed .fc-tb__rb--wide:not(.fc-tb__rb--stacked) > ${LABEL}`;
const CONDENSED_LABEL = `.fc-tb__ribbon .fc-tb__ribbon-group--condensed.fc-tb__ribbon-group--condensed .fc-tb__ribbon-tools > .fc-tb__rb > ${LABEL}`;
const ROW = '.fc-tb__ribbon-group--mixed .fc-tb__rb--stacked';

const TILE_LABEL_MAX = px(declBase(TILES_LABEL, 'max-width'));
const MIXED_TILE_LABEL_MAX = px(declBase(MIXED_TILE_LABEL, 'max-width'));
/** A condensed row sizes to its label; a finite value here would be the tile cap it kept. */
const CONDENSED_LABEL_MAX = decl(CONDENSED_LABEL, 'max-width');
/** `none` reads as NaN: a row as wide as its label never clips. */
const ROW_MAX = px(decl(ROW, 'max-width'));
const ROW_GAP = px(decl(ROW, 'gap'));
const ROW_PAD_X = px(decl(ROW, 'padding')?.split(/\s+/)[1]);
const ROW_ICON = px(decl(`${ROW} .fc-tb__rb-icon`, 'width'));
const ROW_CHEVRON = px(decl(`${ROW} .fc-tb__rb-split-chevron`, 'width'));
const FONT = px(decl(ROW, 'font-size'));

/**
 * Text width per character class, in em of the 10.5px -apple-system label font. Latin factors are
 * calibrated so "Recommended Charts" measures 112px as in the browser: lowercase and accented
 * letters 0.60, capitals 0.70, digits 0.60, space 0.30, punctuation 0.35. CJK, Hangul and
 * full-width forms are 1em; Thai base glyphs 0.55em and combining marks 0.
 */
const charEm = (ch: string): number => {
  if (/[　-鿿가-힯＀-￯]/.test(ch)) return 1;
  if (/[ัิ-ฺ็-๎]/.test(ch)) return 0;
  if (/[฀-๿]/.test(ch)) return 0.55;
  if (/\s/.test(ch)) return 0.3;
  if (/[A-Z]/.test(ch)) return 0.7;
  if (/[\p{L}\d]/u.test(ch)) return 0.6;
  return 0.35;
};
const textWidth = (text: string): number => [...text].reduce((w, ch) => w + charEm(ch) * FONT, 0);

/**
 * Slack before an estimate counts as a clip. The calibration label measures 112px against an
 * estimate of 112.3, so 1px covers that error with room for other glyph mixes.
 */
const MARGIN_PX = 1;

const MIXED_VARIANTS = new Set<string>([
  ...FORMULAS_MIXED_LAYOUT_GROUP_VARIANTS,
  ...INSERT_MIXED_LAYOUT_GROUP_VARIANTS,
  ...DRAW_MIXED_LAYOUT_GROUP_VARIANTS,
  ...PAGE_LAYOUT_MIXED_LAYOUT_GROUP_VARIANTS,
]);

/** Whether the renderer draws a label span for the command at all. */
const hasLabel = (c: RibbonCommand): boolean =>
  c.kind === 'wide' || c.kind === 'large' || (!c.icon && c.kind === 'button');

const TABS = ['formulas', 'insert', 'draw', 'pageLayout', 'review', 'automate'] as const;
const LOCALES = ['en', 'ja', 'de', 'fr', 'ko', 'th', 'zh'] as const;

interface Violation {
  command: string;
  label: string;
  stage: 'condensed row' | 'stacked row';
  need: string;
  have: string;
}

const violations = (tab: (typeof TABS)[number], locale: (typeof LOCALES)[number]): Violation[] => {
  const model = buildRibbonModel(locale, { profile: 'excel365Mac' }).find((t) => t.id === tab);
  const out: Violation[] = [];
  for (const group of model?.groups ?? []) {
    const tiles = group.variant === 'tiles';
    if (!tiles && !(group.variant && MIXED_VARIANTS.has(group.variant))) continue;
    for (const c of group.commands) {
      if (!hasLabel(c)) continue;
      const width = textWidth(c.label);
      if (c.layout === 'stacked') {
        const menu = ribbonActivationForCommand(c.id).menuId !== undefined;
        const space =
          ROW_MAX -
          2 * ROW_PAD_X -
          (c.icon ? ROW_ICON + ROW_GAP : 0) -
          (menu ? ROW_GAP + ROW_CHEVRON : 0);
        if (width > space + MARGIN_PX) {
          out.push({
            command: c.id,
            label: c.label,
            stage: 'stacked row',
            need: `${width.toFixed(1)}px`,
            have: `${space}px`,
          });
        }
        continue;
      }
      const cap = tiles ? TILE_LABEL_MAX : MIXED_TILE_LABEL_MAX;
      // A tile label that needs a third line is widened to two by fit-ribbon (fitLabelToTwoLines,
      // for every tiles or mixed group), so only its demotion to a row can clip it.
      const rowCap = CONDENSED_LABEL_MAX === undefined ? cap : px(CONDENSED_LABEL_MAX);
      if (Number.isFinite(rowCap) && width > rowCap + MARGIN_PX) {
        out.push({
          command: c.id,
          label: c.label,
          stage: 'condensed row',
          need: `${width.toFixed(1)}px`,
          have: `${rowCap}px`,
        });
      }
    }
  }
  return out;
};

describe('Mac ribbon label fit', () => {
  it('reads the label caps from the ribbon CSS', () => {
    for (const value of [
      TILE_LABEL_MAX,
      MIXED_TILE_LABEL_MAX,
      ROW_GAP,
      ROW_PAD_X,
      ROW_ICON,
      ROW_CHEVRON,
      FONT,
    ]) {
      expect(Number.isFinite(value)).toBe(true);
    }
  });

  it('leaves rows uncapped and condensed rows free of the tile cap', () => {
    expect(decl(ROW, 'max-width')).toBe('none');
    expect(CONDENSED_LABEL_MAX).toBe('none');
  });

  it('estimates "Recommended Charts" at the measured 112px', () => {
    expect(Math.abs(textWidth('Recommended Charts') - 112)).toBeLessThanOrEqual(1);
  });

  for (const tab of TABS) {
    for (const locale of LOCALES) {
      it(`never clips a visible ${tab} label in ${locale}`, () => {
        expect(violations(tab, locale)).toEqual([]);
      });
    }
  }
});
