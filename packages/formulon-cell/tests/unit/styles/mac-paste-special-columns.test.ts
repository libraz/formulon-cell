import { readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { afterEach, describe, expect, it } from 'vitest';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import type { Strings } from '../../../src/i18n/strings.js';
import { de, en, fr, ja, ko, th, zh } from '../../../src/i18n/strings.js';
import { attachPasteSpecial } from '../../../src/interact/paste-special.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

const css = readFileSync(
  resolve(
    dirname(fileURLToPath(import.meta.url)),
    '../../../src/styles/core/app/dialog-modules/paste-special.css',
  ),
  'utf8',
).replace(/\/\*[\s\S]*?\*\//g, '');
const MAC = '.fc-overlay-portal[data-fc-platform="mac"]';

/** The Mac rules, selectors whitespace-normalised, each selector of a list on its own. */
const macRules = [...css.matchAll(/([^{}]+)\{([^{}]*)\}/g)].flatMap(
  ([, selectors = '', body = '']) =>
    selectors
      .split(',')
      .map((selector) => selector.replace(/\s+/g, ' ').trim())
      .filter((selector) => selector.startsWith(MAC))
      .map((selector) => ({ selector: selector.slice(MAC.length).trim(), body })),
);
const decl = (selector: string, name: string): string | undefined => {
  let value: string | undefined;
  for (const rule of macRules) {
    if (rule.selector !== selector) continue;
    const hit = rule.body.match(new RegExp(`(?:^|[;\\s])${name}:\\s*([^;]+);`));
    if (hit?.[1]) value = hit[1].trim();
  }
  return value;
};
const px = (value: string | undefined): number => Number.parseFloat(value ?? 'NaN');

/** A label's drawn width: the 23 px radio or checkbox and its gap, then ~6.5 px a Latin
 *  character and 13 px a CJK or Hangul one. */
const itemWidth = (text: string): number =>
  23 +
  [...text].reduce(
    (w, ch) => w + (/[\u3000-\u9fff\uac00-\ud7af\uff00-\uffef]/.test(ch) ? 13 : 6.5),
    0,
  );

const SECTIONS = [
  {
    list: '.fc-pastesp__group:first-child .fc-pastesp__list',
    base: '.fc-pastesp__list',
    item: '.fc-fmtdlg__radio',
    rows: 6,
  },
  {
    list: '.fc-pastesp__group + .fc-pastesp__group .fc-pastesp__list',
    base: '.fc-pastesp__list',
    item: '.fc-fmtdlg__radio',
    rows: 3,
  },
  {
    list: '.fc-pastesp__bottomrow',
    base: '.fc-pastesp__bottomrow',
    item: '.fc-fmtdlg__check',
    rows: 1,
  },
] as const;

/** A section's declaration from its own rule, else from the rule every list of its kind shares. */
const sectionDecl = (section: (typeof SECTIONS)[number], name: string): string | undefined =>
  decl(section.list, name) ?? decl(section.base, name);

/** Where a section's columns start for its labels: from the grid rule when the section is a grid,
 *  otherwise from the absolute left of its first and last items. */
const columns = (section: (typeof SECTIONS)[number], col1: number[]): [number, number] => {
  const tracks = sectionDecl(section, 'grid-template-columns');
  if (tracks) {
    const floor = px(tracks.match(/minmax\(([\d.]+)px/)?.[1]);
    const left = px(sectionDecl(section, 'margin')?.split(/\s+/)[3]);
    return [left, left + Math.max(floor, ...col1) + px(sectionDecl(section, 'column-gap'))];
  }
  const nth = (n: number) => px(decl(`${section.list} > ${section.item}:nth-child(${n})`, 'left'));
  return [nth(1), nth(section.rows + 1)];
};

const LOCALES: [string, Strings][] = [
  ['ja', ja],
  ['en', en],
  ['de', de],
  ['fr', fr],
  ['ko', ko],
  ['th', th],
  ['zh', zh],
];

afterEach(() => {
  document.body.replaceChildren();
});

describe('Mac Paste Special columns', () => {
  for (const [locale, strings] of LOCALES) {
    it(`keeps every ${locale} label clear of the next column and inside the panel`, async () => {
      const host = document.createElement('div');
      host.dataset.fcPlatform = 'mac';
      document.body.appendChild(host);
      const handle = attachPasteSpecial({
        host,
        store: createSpreadsheetStore(),
        wb: await WorkbookHandle.createDefault({ preferStub: true }),
        getSnapshot: () => null,
        strings,
        onAfterCommit: () => {},
      });
      const lists = [
        ...document.querySelectorAll('.fc-pastesp__list'),
        document.querySelector('.fc-pastesp__bottomrow'),
      ];
      let right = 0;
      const col2Starts: number[] = [];
      SECTIONS.forEach((section, index) => {
        const widths = [...(lists[index]?.children ?? [])].map((el) =>
          itemWidth(el.textContent ?? ''),
        );
        const col1 = widths.slice(0, section.rows);
        const col2 = widths.slice(section.rows);
        const [x1, x2] = columns(section, col1);
        col2Starts.push(x2);
        for (const w of col1) expect(x1 + w, `${locale} first column`).toBeLessThanOrEqual(x2 - 4);
        right = Math.max(right, ...col2.map((w) => x2 + w));
      });
      const panel = '.fc-pastesp__panel';
      const width =
        decl(panel, 'width') === 'max-content'
          ? Math.max(
              px(decl(panel, 'min-width')),
              right + px(decl(panel, 'padding')?.split(/\s+/)[1]),
            )
          : px(decl(panel, 'width'));
      expect(right, `${locale} second column`).toBeLessThanOrEqual(width);
      // The measured ja layout stays where it was measured.
      if (locale === 'ja') expect([...col2Starts, width]).toEqual([188, 187, 187, 461]);
      handle.detach();
    });
  }
});
