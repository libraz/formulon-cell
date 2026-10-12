import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { dictionaries } from '../../../src/i18n/strings.js';
import { createFormatDialogView } from '../../../src/interact/format-dialog-view.js';
import { FONT_SIZES } from '../../../src/toolbar/ribbon-model.js';

const en = dictionaries.en;
const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');

describe('interact/format-dialog-view', () => {
  let host: HTMLElement;

  beforeEach(() => {
    host = document.createElement('div');
    document.body.appendChild(host);
  });

  afterEach(() => {
    host.remove();
  });

  it('renders the overlay hidden + with role="dialog" aria-modal="true"', () => {
    const view = createFormatDialogView({ host, strings: en, t: en.formatDialog });
    expect(view.overlay.hidden).toBe(true);
    expect(view.overlay.getAttribute('role')).toBe('dialog');
    expect(view.overlay.getAttribute('aria-modal')).toBe('true');
    expect(view.overlay.getAttribute('aria-label')).toBe(en.formatDialog.title);
    expect(view.closeBtn.getAttribute('aria-label')).toBe(en.formatDialog.cancel);
  });

  it('creates one button + one panel per tab id', () => {
    const view = createFormatDialogView({ host, strings: en, t: en.formatDialog });
    const expectedTabs = ['number', 'align', 'font', 'border', 'fill', 'protection', 'more'];
    expect(view.tabsStrip.getAttribute('role')).toBe('tablist');
    expect(view.tabsStrip.getAttribute('aria-label')).toBe(en.formatDialog.title);
    for (const id of expectedTabs) {
      expect(view.tabButtons.has(id as never), `button missing for tab ${id}`).toBe(true);
      expect(view.tabPanels.has(id as never), `panel missing for tab ${id}`).toBe(true);
    }
  });

  it('starts every tab panel hidden — the controller toggles based on state', () => {
    const view = createFormatDialogView({ host, strings: en, t: en.formatDialog });
    for (const panel of view.tabPanels.values()) {
      expect(panel.hidden).toBe(true);
    }
  });

  it('renders the 12 number-category buttons with role="option"', () => {
    const view = createFormatDialogView({ host, strings: en, t: en.formatDialog });
    expect(view.catButtons.size).toBe(12);
    for (const btn of view.catButtons.values()) {
      expect(btn.getAttribute('role')).toBe('option');
    }
  });

  it('renders the Excel-style horizontal and vertical alignment choices', () => {
    const view = createFormatDialogView({ host, strings: en, t: en.formatDialog });
    expect(view.hAlignRadios.size).toBe(8); // default + left/center/right + fill/justify/center-across/distributed
    expect(view.vAlignRadios.size).toBe(6); // default + top/middle/bottom + justify/distributed
  });

  it('wires data-fc-check on the font + alignment checkboxes', () => {
    const view = createFormatDialogView({ host, strings: en, t: en.formatDialog });
    expect(view.boldCk.input.dataset.fcCheck).toBe('bold');
    expect(view.italicCk.input.dataset.fcCheck).toBe('italic');
    expect(view.underlineSelect.dataset.fcInput).toBe('underline');
    expect(view.underlineSelect.options).toHaveLength(5);
    expect(view.strikeCk.input.dataset.fcCheck).toBe('strike');
    expect(view.superscriptCk.input.dataset.fcCheck).toBe('superscript');
    expect(view.subscriptCk.input.dataset.fcCheck).toBe('subscript');
    expect(view.wrapCk.input.dataset.fcCheck).toBe('wrap');
    expect(view.justifyLastLineCk.input.dataset.fcCheck).toBe('justifyLastLine');
    expect(view.lockedCk.input.dataset.fcCheck).toBe('locked');
  });

  it('renders font family datalist options through the shared dialog option helper', () => {
    const view = createFormatDialogView({ host, strings: en, t: en.formatDialog });
    const datalist = document.getElementById(
      view.familyInput.getAttribute('list') ?? '',
    ) as HTMLDataListElement | null;
    expect(datalist).toBeTruthy();
    expect(Array.from(datalist?.options ?? [], (option) => option.value)).toContain('Arial');
  });

  it('uses the curated Japanese values and keeps the Font tab list/datalist in sync', () => {
    const view = createFormatDialogView({
      host,
      strings: dictionaries.ja,
      t: dictionaries.ja.formatDialog,
      fontLocale: 'ja',
    });
    const datalist = document.getElementById(
      view.familyInput.getAttribute('list') ?? '',
    ) as HTMLDataListElement;
    const datalistValues = Array.from(datalist.options, (option) => option.value);
    const listValues = Array.from(
      view.panel.querySelectorAll<HTMLButtonElement>('button[data-fc-font-family]'),
      (button) => button.dataset.fcFontFamily,
    );
    expect(datalistValues).toEqual(listValues);
    expect(new Set(datalistValues).size).toBe(datalistValues.length);
    expect(datalistValues).toContain('游ゴシック');
    expect(datalistValues).not.toContain('游ゴシック Regular');
    expect(
      view.panel.querySelector<HTMLButtonElement>('button[data-fc-font-family="游ゴシック"]')
        ?.textContent,
    ).toContain('游ゴシック Regular');

    view.syncFontFamilyOptions('Imported Font 14');
    const retainedValues = Array.from(
      view.panel.querySelectorAll<HTMLButtonElement>('button[data-fc-font-family]'),
      (button) => button.dataset.fcFontFamily,
    );
    expect(retainedValues).toContain('Imported Font 14');
    expect(Array.from(datalist.options, (option) => option.value)).toEqual(retainedValues);
    expect(
      Array.from(
        view.panel.querySelectorAll<HTMLButtonElement>('button[data-fc-font-size]'),
        (button) => Number(button.dataset.fcFontSize),
      ),
    ).toEqual([...FONT_SIZES]);
  });

  it('defers selected family and size reveal until the dialog and Font tab are visible', async () => {
    const scrolled: Element[] = [];
    const originalDescriptor = Object.getOwnPropertyDescriptor(
      HTMLElement.prototype,
      'scrollIntoView',
    );
    Object.defineProperty(HTMLElement.prototype, 'scrollIntoView', {
      configurable: true,
      value(this: Element) {
        scrolled.push(this);
      },
    });

    const view = createFormatDialogView({ host, strings: en, t: en.formatDialog });
    try {
      view.syncFontFamilyOptions('Imported Font 14');
      view.syncFontSizeOptions(36);
      expect(scrolled).toHaveLength(0);

      const fontPanel = view.tabPanels.get('font');
      if (!fontPanel) throw new Error('font panel missing');
      view.overlay.hidden = false;
      fontPanel.hidden = false;
      await new Promise<void>((resolve) => requestAnimationFrame(() => resolve()));

      expect(scrolled).toEqual(
        expect.arrayContaining([
          view.panel.querySelector('[data-fc-font-family="Imported Font 14"]'),
          view.panel.querySelector('[data-fc-font-size="36"]'),
        ]),
      );
    } finally {
      view.shell.dispose();
      if (originalDescriptor) {
        Object.defineProperty(HTMLElement.prototype, 'scrollIntoView', originalDescriptor);
      } else {
        Reflect.deleteProperty(HTMLElement.prototype, 'scrollIntoView');
      }
    }
  });

  it('uses the caller ribbon font source for both Format Cells pickers', () => {
    const custom = structuredClone(dictionaries.ja);
    custom.ribbon.fontFamilies = ['Custom Corporate Font'];
    const view = createFormatDialogView({
      host,
      strings: custom,
      t: custom.formatDialog,
      fontLocale: 'ja',
    });
    try {
      const datalist = document.getElementById(
        view.familyInput.getAttribute('list') ?? '',
      ) as HTMLDataListElement;
      const datalistValues = Array.from(datalist.options, (option) => option.value);
      const listValues = Array.from(
        view.panel.querySelectorAll<HTMLButtonElement>('button[data-fc-font-family]'),
        (button) => button.dataset.fcFontFamily,
      );
      expect(datalistValues).toEqual(['Custom Corporate Font']);
      expect(listValues).toEqual(datalistValues);
      expect(listValues).not.toContain('游ゴシック');
    } finally {
      view.shell.dispose();
    }
  });

  it('appends the overlay to document.body so it escapes `.fc-host` contain:strict / isolation:isolate', () => {
    const view = createFormatDialogView({ host, strings: en, t: en.formatDialog });
    expect(document.body.contains(view.overlay)).toBe(true);
    expect(host.contains(view.overlay)).toBe(false);
  });

  it('renders all 13 border-style buttons with the right data attribute', () => {
    const view = createFormatDialogView({ host, strings: en, t: en.formatDialog });
    expect(view.borderStyleButtons.size).toBe(13);
    const ids = Array.from(view.borderStyleButtons.keys()).sort();
    expect(ids).toEqual([
      'dashDot',
      'dashDotDot',
      'dashed',
      'dotted',
      'double',
      'hair',
      'medium',
      'mediumDashDot',
      'mediumDashDotDot',
      'mediumDashed',
      'slantDashDot',
      'thick',
      'thin',
    ]);
    for (const [id, btn] of view.borderStyleButtons.entries()) {
      expect(btn.dataset.borderStyle).toBe(id);
      expect(btn.getAttribute('aria-pressed')).toBe('false');
    }
  });

  it('renders the validation kind/op selectors with the expected option counts', () => {
    const view = createFormatDialogView({ host, strings: en, t: en.formatDialog });
    expect(view.validationKindSelect.options.length).toBe(8); // none + 7 kinds
    expect(view.validationOpSelect.options.length).toBe(8); // between/notBetween/=/<>/</<=/>/>=
    expect(view.validationKindSelect.getAttribute('aria-label')).toBe(
      en.formatDialog.validationKind,
    );
    expect(Array.from(view.validationKindSelect.options, (option) => option.value)).toEqual([
      'none',
      'list',
      'whole',
      'decimal',
      'date',
      'time',
      'textLength',
      'custom',
    ]);
    expect(Array.from(view.validationErrorStyleSelect.options, (option) => option.value)).toEqual([
      'stop',
      'warning',
      'information',
    ]);
  });

  it('renders number tab selects through the shared dialog select contract', () => {
    const view = createFormatDialogView({ host, strings: en, t: en.formatDialog });
    expect(view.symbolSelect.getAttribute('aria-label')).toBe(en.formatDialog.symbol);
    expect(Array.from(view.symbolSelect.options, (option) => option.value)).toContain('€');
    expect(view.patternPresetSelect.dataset.fcSelect).toBe('patternPreset');
    expect(view.patternPresetSelect.getAttribute('aria-label')).toBe(en.formatDialog.patternType);
    expect(
      Array.from(view.localeSelect.options, (option) => [option.value, option.textContent]),
    ).toEqual([
      ['ja', '日本語'],
      ['en', 'English'],
      ['de', 'Deutsch'],
      ['fr', 'Français'],
      ['zh', '简体中文'],
      ['ko', '한국어'],
      ['th', 'ไทย'],
    ]);
    expect(view.localeSelect.disabled).toBe(true);
    expect(view.calendarSelect.dataset.fcSelect).toBe('calendarType');
    expect(Array.from(view.calendarSelect.options, (option) => option.value)).toEqual([
      'gregorian',
      'japanese',
    ]);
  });

  it('keeps font list and alignment dial buttons on shared dialog primitives', () => {
    const fontSource = readFileSync(join(root, 'src/interact/format-dialog-tabs/font.ts'), 'utf8');
    const alignSource = readFileSync(
      join(root, 'src/interact/format-dialog-tabs/align.ts'),
      'utf8',
    );
    expect(fontSource).toContain('appendDialogOptionButton(familyList');
    expect(fontSource).toContain('appendDialogOptionButton(fontStyleList');
    expect(fontSource).toContain('appendDialogOptionButton(sizeList');
    expect(fontSource).not.toContain("document.createElement('button')");
    expect(alignSource).toContain('createDialogToggleButton({');
    expect(alignSource).not.toContain("document.createElement('button')");
  });
});
