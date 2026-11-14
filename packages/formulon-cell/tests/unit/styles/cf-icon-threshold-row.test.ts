import { readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';

const styles = resolve(dirname(fileURLToPath(import.meta.url)), '../../../src/styles/core/app');
const read = (path: string): string => readFileSync(resolve(styles, path), 'utf8');

/** Every rule of `css` as its whitespace-normalised selector and declarations. */
const rules = (css: string): { selector: string; body: Record<string, string> }[] =>
  [...css.matchAll(/([^{}]+)\{([^{}]*)\}/g)].map(([, selector = '', body = '']) => ({
    selector: selector.replace(/\s+/g, ' ').trim(),
    body: Object.fromEntries(
      body
        .split(';')
        .map((decl) => decl.split(':').map((part) => part.trim()))
        .filter((pair): pair is [string, string] => pair.length >= 2 && pair[0] !== '')
        .map(([name, ...value]) => [name, value.join(':')]),
    ),
  }));

const px = (value: string | undefined): number => Number.parseFloat(value ?? 'NaN');

describe('icon-set threshold rows of New/Edit Formatting Rule', () => {
  const conditional = rules(read('dialog-modules/conditional-and-names.css'));
  const select = rules(read('dialog-modules/custom-select.css'));
  const row = conditional.find((r) =>
    r.selector.endsWith('.fc-fmtdlg__row.fc-conddlg__icon-threshold-row'),
  );
  // Label, Operator, Type, Value: the Operator is the second track.
  const operatorTrack = px(row?.body['grid-template-columns']?.split(/\s+(?![^(]*\))/)[1]);
  const buttonMinWidth = (): number => {
    const scoped = conditional.find(
      (r) =>
        r.selector.includes('icon-threshold-row') &&
        r.selector.includes('.fc-select__button') &&
        r.body['min-width'] !== undefined,
    );
    const base = select.find((r) => r.selector === '.fc-select__button');
    return px(scoped?.body['min-width'] ?? base?.body['min-width']);
  };

  it('fits the Operator pop-up button inside its grid track', () => {
    expect(operatorTrack).toBeGreaterThan(0);
    expect(buttonMinWidth()).toBeLessThanOrEqual(operatorTrack);
  });

  it('leaves the Operator track room for >= and the pop-up arrow', () => {
    const base = select.find((r) => r.selector === '.fc-select__button')?.body ?? {};
    const [top, right = top, , left = right] = (base.padding ?? '0').split(/\s+/).map(px);
    const arrow = 8;
    const border = 2;
    const text = 2 * 8;
    expect(operatorTrack).toBeGreaterThanOrEqual(
      (left ?? 0) + text + px(base.gap) + arrow + (right ?? 0) + border,
    );
  });
});
