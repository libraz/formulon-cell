import { describe, expect, it } from 'vitest';
import { computeAutofitRowHeight } from '../../../src/commands/autofit-measurement.js';
import { paintCellText } from '../../../src/render/painters/text.js';
import {
  availableTextWidth,
  CELL_TEXT_HORIZONTAL_PADDING_PX,
  cellTextHorizontalMetrics,
  wrapText,
} from '../../../src/render/text-wrap.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';
import type { ResolvedTheme } from '../../../src/theme/resolve.js';

const theme: ResolvedTheme = {
  bg: '#fff',
  bgRail: '#eee',
  bgElev: '#fff',
  bgHeader: '#eee',
  fg: '#000',
  fgMute: '#666',
  fgFaint: '#aaa',
  fgStrong: '#000',
  rule: '#ccc',
  ruleStrong: '#999',
  freezeRule: '#9b9b9b',
  accent: '#0078d4',
  accentFg: '#fff',
  accentSoft: 'rgba(0,120,212,0.1)',
  cellErrorFg: '#c00',
  cellFormulaFg: '#333',
  cellNumFg: '#000',
  cellBoolFg: '#060',
  hoverStripe: 'rgba(0,0,0,0.04)',
  headerBg: '#eee',
  headerRule: '#ccc',
  headerFg: '#666',
  headerFgActive: '#000',
  pageBackdrop: '#999',
  pagePaper: '#fff',
  pageEdge: '#666',
  pageBandFg: '#888',
  pageBreakAuto: '#06c',
  pageBreakManual: '#036',
  pageNumberFg: '#999',
  pageOutside: '#ddd',
  textCell: 13,
  textHeader: 11.5,
  fontUi: 'sans-serif',
  fontMono: 'monospace',
};

const textContext = (): {
  ctx: CanvasRenderingContext2D;
  fills: string[];
} => {
  const fills: string[] = [];
  let font = '';
  const ctx = {
    get font(): string {
      return font;
    },
    set font(value: string) {
      font = value;
    },
    fillStyle: '',
    direction: 'inherit',
    textBaseline: 'alphabetic',
    textAlign: 'left',
    save(): void {},
    restore(): void {},
    beginPath(): void {},
    rect(): void {},
    clip(): void {},
    moveTo(): void {},
    lineTo(): void {},
    stroke(): void {},
    measureText(value: string): TextMetrics {
      return { width: [...value].length * 10 } as TextMetrics;
    },
    fillText(value: string): void {
      fills.push(value);
    },
  } as unknown as CanvasRenderingContext2D;
  return { ctx, fills };
};

describe('shared wrapped text behavior', () => {
  it('uses one available-width and indent contract for renderer and AutoFit', () => {
    const { ctx, fills } = textContext();
    const text = '東京都千代田区丸の内';
    const format = { indent: 1, wrap: true } as const;
    const bounds = { x: 0, y: 0, w: 76, h: 100 };

    paintCellText({
      ctx,
      bounds,
      theme,
      value: { kind: 'text', value: text },
      formula: null,
      isActive: false,
      isInRange: false,
      format,
      locale: 'ja-JP',
    });

    const store = createSpreadsheetStore();
    store.setState((state) => ({
      ...state,
      data: {
        ...state.data,
        cells: new Map([['0:0:0', { value: { kind: 'text', value: text }, formula: null }]]),
      },
    }));
    mutators.setColWidth(store, 0, bounds.w);
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, format);

    const autoHeight = computeAutofitRowHeight(store.getState(), 0, ctx, { locale: 'ja-JP' });
    expect(autoHeight).toBe(fills.length * Math.round(theme.textCell * 1.28) + 8);
    expect(availableTextWidth(bounds.w, format.indent * 8)).toBe(
      bounds.w - CELL_TEXT_HORIZONTAL_PADDING_PX - format.indent * 8,
    );
  });

  it('wraps Japanese text without whitespace into multiple painted lines', () => {
    const { ctx, fills } = textContext();
    const text = '東京都千代田区丸の内';

    paintCellText({
      ctx,
      bounds: { x: 0, y: 0, w: 36, h: 80 },
      theme,
      value: { kind: 'text', value: text },
      formula: null,
      isActive: false,
      isInRange: false,
      format: { wrap: true },
      locale: 'ja-JP',
    });

    expect(fills.length).toBeGreaterThan(1);
    expect(fills.join('')).toBe(text);
  });

  it.each([
    ['ja-JP', '東京都千代田区丸の内'],
    ['zh-CN', '北京市朝阳区建国门外大街'],
    ['th-TH', 'กรุงเทพมหานครประเทศไทย'],
    ['ko-KR', '서울특별시중구세종대로'],
    ['en-US', 'https://example.test/a/very/long/path'],
    ['de-DE', 'Donaudampfschifffahrtsgesellschaft'],
    ['fr-FR', 'anticonstitutionnellementlongue'],
  ] as const)('wraps a no-space %s string', (locale, text) => {
    const lines = wrapText(text, {
      maxWidth: 30,
      measureText: (value) => [...value].length * 10,
      locale,
    });

    expect(lines.length).toBeGreaterThan(1);
    expect(lines.join('')).toBe(text);
  });

  it('preserves CRLF blank lines as empty painted lines', () => {
    const { ctx, fills } = textContext();

    paintCellText({
      ctx,
      bounds: { x: 0, y: 0, w: 120, h: 80 },
      theme,
      value: { kind: 'text', value: 'A\r\n\r\nB' },
      formula: null,
      isActive: false,
      isInRange: false,
      format: { wrap: true },
      locale: 'en-US',
    });

    expect(fills).toEqual(['A', '', 'B']);
  });

  it('uses the same no-whitespace wrapping when AutoFit measures row height', () => {
    const store = createSpreadsheetStore();
    const text = '東京都千代田区丸の内';
    store.setState((state) => ({
      ...state,
      data: {
        ...state.data,
        cells: new Map([['0:0:0', { value: { kind: 'text', value: text }, formula: null }]]),
      },
    }));
    mutators.setColWidth(store, 0, 36);
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { wrap: true });

    expect(computeAutofitRowHeight(store.getState(), 0, null, { locale: 'ja-JP' })).toBeGreaterThan(
      store.getState().layout.defaultRowHeight + 5,
    );
  });

  it('splits long Latin tokens and keeps combining and ZWJ graphemes intact', () => {
    const Segmenter = (
      globalThis.Intl as unknown as {
        Segmenter: new (
          locales?: string,
          options?: { granularity?: 'grapheme' },
        ) => { segment(value: string): Iterable<{ segment: string }> };
      }
    ).Segmenter;
    const segmenter = new Segmenter('en', { granularity: 'grapheme' });
    const measure = (value: string): number => Array.from(segmenter.segment(value)).length * 10;
    const text = 'a\u0301👩‍💻abcdef';

    const lines = wrapText(text, { maxWidth: 10, measureText: measure, locale: 'en-US' });

    expect(lines).toEqual(['a\u0301', '👩‍💻', 'a', 'b', 'c', 'd', 'e', 'f']);
  });

  it('preserves blank paragraphs and falls back when Intl.Segmenter is unavailable', () => {
    const intl = globalThis.Intl as unknown as { Segmenter?: unknown };
    const segmenter = intl.Segmenter;
    try {
      intl.Segmenter = undefined;
      expect(
        wrapText('A\r\n\r\nB', {
          maxWidth: 100,
          measureText: (value) => value.length * 10,
          locale: 'ja-JP',
        }),
      ).toEqual(['A', '', 'B']);
      expect(
        wrapText('a\u0301👩‍💻b', {
          maxWidth: 10,
          measureText: (value) => [...value].length * 10,
          locale: 'en-US',
        }),
      ).toEqual(['a\u0301', '👩‍💻', 'b']);
    } finally {
      intl.Segmenter = segmenter;
    }
  });

  it('keeps authored paragraph-edge whitespace while trimming automatic separators', () => {
    const measureText = (value: string): number => value.length * 10;
    expect(wrapText('  alpha', { maxWidth: 100, measureText })).toEqual(['  alpha']);
    expect(wrapText('alpha  ', { maxWidth: 100, measureText })).toEqual(['alpha  ']);
    expect(wrapText('   ', { maxWidth: 100, measureText })).toEqual(['   ']);
    expect(wrapText('alpha   beta', { maxWidth: 50, measureText })).toEqual(['alpha', 'beta']);
  });

  it('keeps CJK closing punctuation with the preceding grapheme when it fits', () => {
    expect(
      wrapText('あいう、えお。かき', {
        maxWidth: 20,
        measureText: (value) => [...value].length * 10,
        locale: 'ja-JP',
      }),
    ).toEqual(['あい', 'う、', 'え', 'お。', 'かき']);
    const bracketLines = wrapText('「あいう」', {
      maxWidth: 20,
      measureText: (value) => [...value].length * 10,
      locale: 'ja-JP',
    });
    expect(bracketLines.some((line) => line.endsWith('「') || line.startsWith('」'))).toBe(false);
    expect(
      wrapText('「あ', {
        maxWidth: 10,
        measureText: (value) => [...value].length * 10,
        locale: 'ja-JP',
      }).join(''),
    ).toBe('「あ');
  });

  it('relaxes an impossible CJK punctuation pair at one-grapheme width', () => {
    expect(
      wrapText('あいう、', {
        maxWidth: 10,
        measureText: (value) => [...value].length * 10,
        locale: 'ja-JP',
      }),
    ).toEqual(['あ', 'い', 'う', '、']);
    expect(
      wrapText('「あいう」', {
        maxWidth: 10,
        measureText: (value) => [...value].length * 10,
        locale: 'ja-JP',
      }),
    ).toEqual(['「', 'あ', 'い', 'う', '」']);
  });

  it('does not construct an Intl.Segmenter for every grapheme in a long token', () => {
    const intl = globalThis.Intl as unknown as {
      Segmenter?: typeof Intl.Segmenter;
    };
    const NativeSegmenter = intl.Segmenter;
    if (!NativeSegmenter) throw new Error('Intl.Segmenter is required for this regression.');
    const SegmenterCtor: typeof Intl.Segmenter = NativeSegmenter;
    let constructions = 0;
    const CountingSegmenter = class {
      private readonly delegate: InstanceType<typeof SegmenterCtor>;

      constructor(...args: ConstructorParameters<typeof SegmenterCtor>) {
        constructions += 1;
        this.delegate = new SegmenterCtor(...args);
      }

      segment(input: string): ReturnType<InstanceType<typeof SegmenterCtor>['segment']> {
        return this.delegate.segment(input);
      }
    };
    intl.Segmenter = CountingSegmenter as unknown as typeof Intl.Segmenter;
    try {
      const text = 'a'.repeat(32_767);
      const lines = wrapText(text, {
        maxWidth: 10,
        measureText: (value) => value.length * 10,
        locale: 'en-US',
      });
      expect(lines.join('')).toBe(text);
      expect(constructions).toBeLessThanOrEqual(4);
    } finally {
      intl.Segmenter = NativeSegmenter;
    }
  });

  it('keeps French nonbreaking connectors intact when an ordinary space may break', () => {
    const measureText = (value: string): number => [...value].length * 10;
    const options = {
      maxWidth: 70,
      measureText,
      locale: 'fr-FR',
    };
    expect(wrapText('bonjour !', options)).toEqual(['bonjour', '!']);
    expect(wrapText('bonjour\u00a0!', options)).toEqual(['bonjour\u00a0!']);
    expect(wrapText('bonjour\u202f!', options)).toEqual(['bonjour\u202f!']);
    expect(wrapText('bonjour\u2060!', options)).toEqual(['bonjour\u2060!']);
    expect(wrapText('bonjour\u00a0! salut', options)).toEqual(['bonjour\u00a0!', 'salut']);
  });

  it('uses scaled horizontal metrics as one canvas contract', () => {
    expect(cellTextHorizontalMetrics(100, 2, 0.5)).toEqual({
      padX: 1.5,
      indentPx: 8,
      availableWidth: 89,
    });
    expect(cellTextHorizontalMetrics(100, 2, 1)).toEqual({
      padX: 3,
      indentPx: 16,
      availableWidth: 78,
    });
    expect(cellTextHorizontalMetrics(100, 2, 2)).toEqual({
      padX: 6,
      indentPx: 32,
      availableWidth: 56,
    });
  });

  it('keeps wrapped lines and origins stable when zoom scales the cell and glyphs', () => {
    const renderedAt = (zoom: number): { lines: string[]; x: number; lineGap: number } => {
      const fills: Array<{ text: string; x: number; y: number }> = [];
      let font = '';
      const ctx = {
        get font(): string {
          return font;
        },
        set font(value: string) {
          font = value;
        },
        fillStyle: '',
        direction: 'inherit',
        textBaseline: 'alphabetic',
        textAlign: 'left',
        save(): void {},
        restore(): void {},
        beginPath(): void {},
        rect(): void {},
        clip(): void {},
        moveTo(): void {},
        lineTo(): void {},
        stroke(): void {},
        measureText(value: string): TextMetrics {
          const size = Number.parseFloat(font.match(/\d+(?:\.\d+)?px/u)?.[0] ?? '13');
          return { width: [...value].length * (10 * (size / 13)) } as TextMetrics;
        },
        fillText(text: string, x: number, y: number): void {
          fills.push({ text, x, y });
        },
      } as unknown as CanvasRenderingContext2D;
      paintCellText({
        ctx,
        bounds: { x: 0, y: 0, w: 40 * zoom, h: 100 * zoom },
        theme,
        value: { kind: 'text', value: 'abcd' },
        formula: null,
        isActive: false,
        isInRange: false,
        format: { indent: 1, wrap: true, vAlign: 'top' },
        zoom,
      });
      return {
        lines: fills.map(({ text }) => text),
        x: fills[0]?.x ?? Number.NaN,
        lineGap: (fills[1]?.y ?? 0) - (fills[0]?.y ?? 0),
      };
    };

    expect(renderedAt(0.5)).toEqual({ lines: ['ab', 'cd'], x: 5.5, lineGap: 8.5 });
    expect(renderedAt(1)).toEqual({ lines: ['ab', 'cd'], x: 11, lineGap: 17 });
    expect(renderedAt(2)).toEqual({ lines: ['ab', 'cd'], x: 22, lineGap: 34 });
  });

  it('wraps with logical font metrics and restores the scaled paint font', () => {
    const renderedAt = (zoom: number): { lines: string[]; fonts: string[] } => {
      const fills: Array<{ text: string; font: string }> = [];
      const fonts: string[] = [];
      let font = '';
      const ctx = {
        get font(): string {
          return font;
        },
        set font(value: string) {
          font = value;
          fonts.push(value);
        },
        fillStyle: '',
        direction: 'inherit',
        textBaseline: 'alphabetic',
        textAlign: 'left',
        save(): void {},
        restore(): void {},
        beginPath(): void {},
        rect(): void {},
        clip(): void {},
        moveTo(): void {},
        lineTo(): void {},
        stroke(): void {},
        measureText(value: string): TextMetrics {
          const size = Number.parseFloat(font.match(/\d+(?:\.\d+)?px/u)?.[0] ?? '13');
          const widthPerGrapheme = size < 10 ? 16 : 10;
          return { width: [...value].length * widthPerGrapheme } as TextMetrics;
        },
        fillText(text: string): void {
          fills.push({ text, font });
        },
      } as unknown as CanvasRenderingContext2D;
      paintCellText({
        ctx,
        bounds: { x: 0, y: 0, w: 40 * zoom, h: 100 * zoom },
        theme,
        value: { kind: 'text', value: 'abcd' },
        formula: null,
        isActive: false,
        isInRange: false,
        format: { wrap: true, vAlign: 'top' },
        zoom,
      });
      return { lines: fills.map(({ text }) => text), fonts: fills.map(({ font }) => font) };
    };

    expect(renderedAt(0.5)).toEqual({
      lines: ['abc', 'd'],
      fonts: ['400 6.5px sans-serif', '400 6.5px sans-serif'],
    });
    expect(renderedAt(1)).toEqual({
      lines: ['abc', 'd'],
      fonts: ['400 13px sans-serif', '400 13px sans-serif'],
    });
    expect(renderedAt(2)).toEqual({
      lines: ['abc', 'd'],
      fonts: ['400 26px sans-serif', '400 26px sans-serif'],
    });
  });

  it('keeps AutoFit on the same logical line count and unzoomed row height', () => {
    const store = createSpreadsheetStore();
    store.setState((state) => ({
      ...state,
      data: {
        ...state.data,
        cells: new Map([['0:0:0', { value: { kind: 'text', value: 'abcd' }, formula: null }]]),
      },
    }));
    mutators.setColWidth(store, 0, 40);
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { indent: 1, wrap: true });

    const { ctx } = textContext();
    expect(computeAutofitRowHeight(store.getState(), 0, ctx, { locale: 'en-US' })).toBe(42);
  });
});
