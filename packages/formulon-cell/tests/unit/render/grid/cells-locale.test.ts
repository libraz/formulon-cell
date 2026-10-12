import { describe, expect, it } from 'vitest';

import { formatA1FormulaAsR1C1 } from '../../../../src/commands/refs.js';
import type { WorkbookHandle as WorkbookHandleType } from '../../../../src/engine/workbook-handle.js';
import { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { buildColLayout, buildRowLayout, layoutForView } from '../../../../src/render/geometry.js';
import { type CellDisplayResolver, paintCells } from '../../../../src/render/grid/cells.js';
import { createSpreadsheetStore, mutators } from '../../../../src/store/store.js';
import type { ResolvedTheme } from '../../../../src/theme/resolve.js';

const theme = {
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
  accent: '#0078d4',
  accentFg: '#fff',
  accentSoft: 'rgba(0,120,212,0.1)',
  cellErrorFg: '#c00',
  cellFormulaFg: '#333',
  cellBoolFg: '#060',
  cellNumFg: '#000',
  hoverStripe: 'rgba(0,0,0,0.04)',
  headerFg: '#666',
  headerFgActive: '#000',
  pageBackdrop: '#aaa',
  pagePaper: '#fff',
  pageEdge: '#999',
  pageBandFg: '#666',
  pageBreakAuto: '#2f6fd0',
  pageBreakManual: '#1a4fa0',
  pageNumberFg: '#999',
  pageOutside: '#ddd',
  fontUi: 'sans-serif',
  fontMono: 'monospace',
  textCell: 13,
  textHeader: 11.5,
} as ResolvedTheme;

interface CanvasTextSpy {
  ctx: CanvasRenderingContext2D;
  texts: string[];
}

function canvasTextSpy(): CanvasTextSpy {
  const texts: string[] = [];
  const target: Record<string, unknown> = {
    fillStyle: '',
    font: '',
    textAlign: 'left',
    textBaseline: 'alphabetic',
    direction: 'inherit',
    globalAlpha: 1,
    measureText: () =>
      ({
        width: 7,
        actualBoundingBoxAscent: 9,
        actualBoundingBoxDescent: 2,
      }) as TextMetrics,
    fillText: (text: string) => texts.push(text),
  };
  const ctx = new Proxy(target, {
    get(object, property: string | symbol) {
      if (property in object) return object[property as string];
      return () => undefined;
    },
    set(object, property: string | symbol, value: unknown) {
      object[property as string] = value;
      return true;
    },
  }) as unknown as CanvasRenderingContext2D;
  return { ctx, texts };
}

function paintOne({
  formula,
  value = { kind: 'number', value: 3 },
  wb = null,
  showFormulas = true,
  r1c1 = false,
  getDisplay,
}: {
  formula: string;
  value?: { kind: 'number'; value: number };
  wb?: WorkbookHandleType | null;
  showFormulas?: boolean;
  r1c1?: boolean;
  getDisplay?: CellDisplayResolver;
}): CanvasTextSpy {
  const store = createSpreadsheetStore();
  const addr = { sheet: 0, row: 2, col: 2 } as const;
  mutators.setCell(store, addr, value, formula);
  store.setState((state) => ({
    ...state,
    ui: { ...state.ui, showFormulas, r1c1 },
    viewport: { ...state.viewport, widthPx: 640, rowCount: 8, colCount: 8 },
  }));
  const state = store.getState();
  const view = { ...state, layout: layoutForView(state) };
  const spy = canvasTextSpy();
  paintCells(
    {
      ctx: spy.ctx,
      wb,
      locale: 'en-US',
      getDisplay,
      sheetBackgroundImage: null,
    },
    view,
    theme,
    buildColLayout(view.layout, view.viewport),
    buildRowLayout(view.layout, view.viewport),
  );
  return spy;
}

describe('grid formula painting follows the active language profile', () => {
  it.each([
    ['de', 'windows-de_DE', '=SUMME(A1;B1)'],
    ['fr', 'windows-fr_FR', '=SOMME(A1;B1)'],
  ] as const)(
    'draws localized %s function names on the real canvas path',
    async (_locale, profile, expected) => {
      const wb = await WorkbookHandle.createDefault();
      try {
        expect(wb.setSpreadsheetProfileId(profile)).toBe(true);
        expect(paintOne({ formula: '=SUM(A1,B1)', wb }).texts).toContain(expected);
      } finally {
        wb.dispose();
      }
    },
  );

  it('keeps A1 references in localized formula view and resolves native R1C1 references once', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      expect(wb.setSpreadsheetProfileId('windows-de_DE')).toBe(true);
      const addr = { sheet: 0, row: 2, col: 2 } as const;
      wb.setNumber({ sheet: 0, row: 0, col: 0 }, 1);
      wb.setNumber({ sheet: 0, row: 1, col: 1 }, 2);
      wb.setFormula(addr, '=SUM(A1,$B$2)');

      const a1 = paintOne({ formula: '=SUM(A1,$B$2)', wb }).texts;
      expect(a1).toContain('=SUMME(A1;$B$2)');

      const r1c1 = paintOne({ formula: '=SUM(A1,$B$2)', wb, r1c1: true }).texts;
      expect(r1c1).toContain('=SUMME(R[-2]C[-2];R2C2)');
    } finally {
      wb.dispose();
    }
  });

  it('preserves quoted sheet names and structured references while localizing functions', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      expect(wb.setSpreadsheetProfileId('windows-de_DE')).toBe(true);
      const formula = "=SUM('O''SUMME(A)'!B1,Table1[SUMME(A)])";
      expect(paintOne({ formula, wb }).texts).toContain(
        "=SUMME('O''SUMME(A)'!B1;Table1[SUMME(A)])",
      );
    } finally {
      wb.dispose();
    }
  });

  it('keeps the callback input canonical and displays the calculated result when formulas are off', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      expect(wb.setSpreadsheetProfileId('windows-de_DE')).toBe(true);
      const seen: string[] = [];
      const callbackTexts = paintOne({
        formula: '=SUM(A1,B1)',
        wb,
        showFormulas: false,
        getDisplay: (_addr, _value, formula) => {
          if (formula) seen.push(formula);
          return null;
        },
      }).texts;
      expect(seen).toEqual(['=SUM(A1,B1)']);
      expect(callbackTexts).toContain('3');
    } finally {
      wb.dispose();
    }
  });

  it('falls back to the standalone A1-to-R1C1 conversion without a workbook', () => {
    const formula = '=SUM(A1,$B$2)';
    expect(paintOne({ formula, wb: null, r1c1: true }).texts).toContain(
      formatA1FormulaAsR1C1(formula, { row: 2, col: 2 }),
    );
  });

  it('falls back safely for a partial workbook object', () => {
    const partialWb = { capabilities: {} } as WorkbookHandleType;
    const formula = '=SUM(A1,$B$2)';
    expect(paintOne({ formula, wb: partialWb, r1c1: true }).texts).toContain(
      formatA1FormulaAsR1C1(formula, { row: 2, col: 2 }),
    );
  });
});
