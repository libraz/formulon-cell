import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import {
  macFormatDialogDe,
  macFormatDialogEn,
  macFormatDialogFr,
  macFormatDialogKo,
  macFormatDialogTh,
  macFormatDialogZh,
} from '../../../src/i18n/strings/mac/format-dialog.js';

const css = readFileSync(
  join(__dirname, '../../../src/styles/core/app/format-dialog/mac-layout.css'),
  'utf8',
);
const TAB = '.fc-fmtdlg--mac .fc-fmtdlg__panel-tab[data-fc-tab="align"] ';

/** The declarations of the align tab rule for `selector`. */
const declarations = (selector: string): Record<string, string> => {
  const start = css.indexOf(`${TAB}${selector} {`);
  if (start < 0) throw new Error(`no rule for ${selector}`);
  const body = css.slice(css.indexOf('{', start) + 1, css.indexOf('}', start));
  return Object.fromEntries(
    body
      .split(';')
      .map((line) => line.trim().split(/:\s*/))
      .filter((pair): pair is [string, string] => pair.length === 2),
  );
};
const px = (value: string | undefined): number => Number.parseFloat(value ?? 'NaN');

/** Every other align tab frame the generated rules place, as [left, top, right, bottom]. */
const alignFrames = (except: string): [number, number, number, number][] =>
  [...css.matchAll(/([^{}]+)\{([^}]*)\}/g)]
    .filter(([, selector = '']) => {
      const flat = selector.replace(/\s+/g, ' ').trim();
      return flat.startsWith(TAB) && !flat.includes(except);
    })
    .map(([, , body]) => {
      const value = (name: string): number =>
        px(body?.match(new RegExp(`(?:^|\\s)${name}:\\s*([\\d.]+)px`))?.[1]);
      const left = value('left');
      const top = value('top');
      return [left, top, left + value('width'), top + value('height')] as [
        number,
        number,
        number,
        number,
      ];
    })
    .filter((frame) => frame.every(Number.isFinite));

/** Where a label's text lands for a run `width` px wide, as the rule places it. */
const textSpan = (rule: Record<string, string>, width: number): [number, number] => {
  const left = px(rule.left);
  const fixed = rule.width !== 'max-content';
  if (fixed) return [left, left + width];
  return rule.transform === 'translateX(-100%)' ? [left - width, left] : [left, left + width];
};

/** A generous width for a 12 px system-font label: 8 px a character. */
const runWidth = (text: string): number => [...text].length * 8;

describe('Mac Format Cells alignment degrees label', () => {
  const labels: Record<string, string | undefined> = {
    ja: '度:',
    en: macFormatDialogEn.rotation,
    de: macFormatDialogDe.rotation,
    fr: macFormatDialogFr.rotation,
    ko: macFormatDialogKo.rotation,
    th: macFormatDialogTh.rotation,
    zh: macFormatDialogZh.rotation,
  };

  it.each(Object.entries(labels))('keeps the %s label clear of the degrees field', (_, text) => {
    const label = declarations('.fc-fmtdlg__align-degree > span');
    const input = declarations('.fc-fmtdlg__align-degree input');
    const [start, end] = textSpan(label, runWidth(text ?? ''));

    const top = px(label.top);
    const bottom = top + px(label.height);

    expect(end).toBeLessThanOrEqual(px(input.left));
    expect(start).toBeGreaterThanOrEqual(0);
    const hits = alignFrames('.fc-fmtdlg__align-degree > span').filter(
      ([left, frameTop, right, frameBottom]) =>
        left < end && right > start && frameTop < bottom && frameBottom > top,
    );
    expect(hits).toEqual([]);
  });
});
