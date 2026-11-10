import { expect, type Page } from '@playwright/test';
import type { E2eProbes } from '../../../apps/demo-shared/e2e-probes.js';
import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/mount/types.js';
import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

export async function runWrappedText(page: Page): Promise<void> {
  const spreadsheet = new SpreadsheetPage(page);
  const texts = {
    en: 'averylongunbrokenLatinwordforwrapping',
    ja: '東京都千代田区丸の内一丁目、二丁目。',
    de: 'Donaudampfschifffahrtsgesellschaft',
    fr: '  bonjour\u00a0! encore du texte',
    zh: '北京市海淀区中关村科技园区。',
    ko: '대한민국서울특별시종로구대한민국',
    th: 'ประเทศไทยกรุงเทพมหานครประเทศไทย',
  };
  for (const [locale, text] of Object.entries(texts)) {
    await spreadsheet.mount({ locale });
    await spreadsheet.expectNoStub();
    const result = await page.evaluate(
      ({ locale, text }) => {
        const instance = (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst;
        const probes = (window as unknown as { __fcProbes: E2eProbes }).__fcProbes;
        const { paintCellText } = probes['render/painters/text'];
        const { computeAutofitRowHeight } = probes['commands/autofit-measurement'];
        const { mutators } = probes['store/store'];
        const { resolveTheme } = probes['theme/resolve'];
        const font = instance.workbook.workbookDisplayFont;
        if (!font) throw new Error('Missing workbook display font');
        const format = {
          wrap: true,
          fontSize: font.size * (4 / 3),
          fontFamily: font.name,
          indent: 1,
        };
        instance.applyChanges([{ addr: { sheet: 0, row: 0, col: 0 }, input: text }]);
        mutators.setCellFormat(instance.store, { sheet: 0, row: 0, col: 0 }, format);
        mutators.setColWidth(instance.store, 0, 80);
        const theme = resolveTheme(instance.host);
        const painted: string[][] = [];
        for (const zoom of [0.5, 1, 2]) {
          const ctx = document.createElement('canvas').getContext('2d');
          if (!ctx) throw new Error('Missing real canvas context');
          const lines: string[] = [];
          ctx.fillText = (line: string) => {
            lines.push(line);
          };
          paintCellText({
            ctx,
            theme,
            bounds: { x: 0, y: 0, w: 80 * zoom, h: 400 * zoom },
            value: { kind: 'text', value: text },
            formula: null,
            format,
            locale,
            zoom,
            isActive: false,
            isInRange: false,
          });
          painted.push(lines);
        }
        const ctx = document.createElement('canvas').getContext('2d');
        if (!ctx) throw new Error('Missing real canvas context');
        const height = computeAutofitRowHeight(instance.store.getState(), 0, ctx, {
          locale,
          theme,
          workbookDisplayFont: font,
        });
        // The cell uses the Normal font, so the row padding is the default
        // height minus one Normal line pitch.
        const defaultHeight = instance.store.getState().layout.defaultRowHeight;
        const pitch = Math.round(format.fontSize * 1.28);
        return {
          painted,
          height,
          expectedHeight: Math.max(
            defaultHeight,
            painted[1].length * pitch + Math.max(0, defaultHeight - pitch),
          ),
        };
      },
      { locale, text },
    );
    expect(result.painted[1].length, `${locale} wraps in the real browser font`).toBeGreaterThan(1);
    expect(result.painted[0], `${locale} 50%`).toEqual(result.painted[1]);
    expect(result.painted[2], `${locale} 200%`).toEqual(result.painted[1]);
    expect(result.height, `${locale} AutoFit matches painted line count`).toBe(
      result.expectedHeight,
    );
  }
}
