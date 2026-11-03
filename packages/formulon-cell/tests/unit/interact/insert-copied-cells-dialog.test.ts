import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');
const css = readFileSync(
  join(root, 'src/styles/core/app/dialog-modules/insert-copied-cells.css'),
  'utf8',
).replace(/\/\*[\s\S]*?\*\//g, '');
const mac = '.fc-overlay-portal[data-fc-platform="mac"]';

/** The declarations of every rule whose selector list names `selector`. */
const ruleFor = (selector: string): string => {
  let declarations = '';
  for (const match of css.matchAll(/([^{}]+)\{([^{}]*)\}/g)) {
    const selectors = (match[1] ?? '').split(',').map((s) => s.replace(/\s+/g, ' ').trim());
    if (selectors.includes(selector)) declarations += match[2] ?? '';
  }
  return declarations;
};

describe('Insert Copied Cells, Mac geometry', () => {
  // Excel for Mac: 'コピーしたセルを挿入' is 221 x 146 pt — the Insert dialog's
  // frame with two radios, OK / Cancel 86 x 26 at y 101 (15 pt under the radios).
  it('shares the Insert / Delete cells frame', () => {
    expect(ruleFor(`${mac} .fc-insertcopied__panel`)).toContain('width: 221px');
    expect(ruleFor(`${mac} .fc-insertcopied__title`)).toContain('height: 32px');
    expect(ruleFor(`${mac} .fc-insertcopied__choices`)).toContain('padding: 18px 19px 0 20px');
    expect(ruleFor(`${mac} .fc-insertcopied__radio`)).toContain('width: 16px');
    expect(ruleFor(`${mac} .fc-insertcopied__button`)).toContain('height: 26px');
  });

  it('sits the buttons 15 pt under the two radios', () => {
    expect(ruleFor(`${mac} .fc-insertcopied__footer`)).toContain('padding-top: 15px');
  });
});
