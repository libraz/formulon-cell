import { describe, expect, it } from 'vitest';

import {
  effectiveCellFormatFromEngine,
  hydrateCellFormatsFromEngine,
  syncCellFormatsToEngine,
} from '../../../src/engine/cell-format-sync.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { buildColLayout, buildRowLayout, layoutForView } from '../../../src/render/geometry.js';
import { paintCells } from '../../../src/render/grid/cells.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';
import type { ResolvedTheme } from '../../../src/theme/resolve.js';

const INHERITED_STYLES_XLSX =
  'UEsDBBQAAAAIAC1bR122+9qcCQEAAK4CAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbK1SO0/DMBDe+RWW1yp2yoAQatqhwAgM5QccziWx4pd8bkn+PU4KHVB5DJ1O9vfU6VabwRp2wEjau4ovRckZOuVr7dqKv+4ei1vOKIGrwXiHFR+R+GZ9tdqNAYllsaOKdymFOylJdWiBhA/oMtL4aCHlZ2xlANVDi/K6LG+k8i6hS0WaPHg2u8cG9iaxhyH/H5tENMTZ9sicwioOIRitIGVcHlz9Lab4jBBZOXOo04EWmcDl+YgJ+jnhS/iclxN1jewFYnoCm2lyMPLdx/7N+1787nKmp28arbD2am+zRFCICDV1iMkaMU9hQbvFPwrMbJLzWF64ycn/ryKURoN06T3MpqdoOZ/b+gNQSwMEFAAAAAgALVtHXX5vwIWxAAAAKgEAAAsAAABfcmVscy8ucmVsc43POw7CMAwG4J1TRN5pWgaEUEMXhNQVlQOE1H2oSRwlAdrbkxEqBkbL/j/bZTUbzZ7ow0hWQJHlwNAqakfbC7g1l+0BWIjStlKTRQELBqhOm/KKWsaUCcPoAkuIDQKGGN2R86AGNDJk5NCmTkfeyJhK33Mn1SR75Ls833P/acAKZXUrwNdtAaxZHP6DU9eNCs+kHgZt/LFjNZFk6XuMAmbNX+SnO9GUJRR4OoZ/vXh6A1BLAwQUAAAACAAtW0ddi2DA2M0AAABEAQAADwAAAHhsL3dvcmtib29rLnhtbI2QTU/DMAyG7/sVke8sLQeEqra7TEg7w7iHxl2jNXZlZx/8exLGEEdOfi3Ljx673VzjbM4oGpg6qNcVGKSBfaBDB/u3l4dnMJoceTczYQefqLDpV+2F5fjBfDR5n7SDKaWlsVaHCaPTNS9IeTKyRJdyKweri6DzOiGmONvHqnqy0QWCG6GR/zB4HMOAWx5OESndIIKzS9lep7AoZLWi9R7wov2vZWlt39o/o2+RezXkYj7uteQ6H1zqzud/gJEm5CA7X0Mh/KytWntn919QSwMEFAAAAAgALVtHXR+qsIPGAAAAqwEAABoAAAB4bC9fcmVscy93b3JrYm9vay54bWwucmVsc62QzaoCMQyF9/cpSvZOZlyIiNWNCG5FH6B0Mj8405Ym/szbWxQGFS/cxV2Fk5DvHM5yfes7daHIrXcaiiwHRc76snW1huNhO5mDYjGuNJ13pGEghvXqZ7mnzkj64aYNrBLEsYZGJCwQ2TbUG858IJculY+9kSRjjcHYk6kJp3k+w/jKgA+o2pUa4q4sQB2GQH+B+6pqLW28Pffk5IsHXn08cUMkCWpiTaJhXDE+RpElKuAvaab/mYZl6FKdY5SnHv3xrePVHVBLAwQUAAAACAAtW0ddvCswdooBAACGAwAADQAAAHhsL3N0eWxlcy54bWydk8Fu2zAMhu97CkH3VXYwFEVhu4cCBnbZpR2wqxzJjgCKMiSlSPb0Iy27SYAWw+aDRdLkx1+i3DydPIg3G5ML2Mr6rpLC4j4Yh1Mrf772Xx+kSFmj0RDQtvJsk3zqvjQpn8G+HKzNggiYWnnIeX5UKu0P1ut0F2aL9GUM0etMbpxUmqPVJnGRB7WrqnvltUNJODz63uck9uGIuZWVVF0zBrxEdrIEuib9Fm8aSGvNSai9Lf6zBjdEx0FVMst7UP9SsyyJFI0O4LY7Bbpm1jnbiD05YrVfzzMdDNLxFM6S95fsFMAZhk7PAUIUcRpa2dNTLVsf1rBDY0/WtPL+28K+4r13WhZWPIRoaJCb5prwJdQ1YMdM9dFNB15zmLlJyDl4MozTU0AN3GGrWA3m7i3AC4/713gDP42ijO274YkJPrnNJEmrWTjF4QbXtBV+xd39L/ezovpSVN8WCT3PcO5D2c7qUSp7m9CLxkXyzfbfo4LvUyt/8E0HKU7jpvDoIDv8YOsMVZc/qPsDUEsDBBQAAAAIAC1bR11HOVWXJQEAAAQCAAAYAAAAeGwvd29ya3NoZWV0cy9zaGVldDEueG1sjVFNb4MwDL3zKyLf11A2bRMKVFWrSrvv45yBKVFJgpIU2n8/B7YO7bRDHDt+7/kjYnPRHRvQeWVNAetVCgxNZWtljgW8vR7unoH5IE0tO2uwgCt62JSJGK07+RYxMBIwvoA2hD7n3FctaulXtkdDmcY6LQOF7sh971DWE0l3PEvTR66lMjAr5O4/GrZpVIV7W501mjCLOOxkoPZ9q3oP1FutKBkHYg6bArbrfHcPvBRT6XeFo1/4LE7yae0pBi91AWmE8gU2EZXtiEKWaUVbyoBpeZnuUdWhpb2RW519sPrj+yFu7dph9KLcpJDMVfcyyFI4OzI3AWfiYZpyZkZL9WJ+SzG9GoqH8iETfIhqdIh/E8l+0LtsgX76A+a/1RPBbx9YfgFQSwECFAMUAAAACAAtW0ddtvvanAkBAACuAgAAEwAAAAAAAAAAAAAAgAEAAAAAW0NvbnRlbnRfVHlwZXNdLnhtbFBLAQIUAxQAAAAIAC1bR11+b8CFsQAAACoBAAALAAAAAAAAAAAAAACAAToBAABfcmVscy8ucmVsc1BLAQIUAxQAAAAIAC1bR12LYMDYzQAAAEQBAAAPAAAAAAAAAAAAAACAARQCAAB4bC93b3JrYm9vay54bWxQSwECFAMUAAAACAAtW0ddH6qwg8YAAACrAQAAGgAAAAAAAAAAAAAAgAEOAwAAeGwvX3JlbHMvd29ya2Jvb2sueG1sLnJlbHNQSwECFAMUAAAACAAtW0ddvCswdooBAACGAwAADQAAAAAAAAAAAAAAgAEMBAAAeGwvc3R5bGVzLnhtbFBLAQIUAxQAAAAIAC1bR11HOVWXJQEAAAQCAAAYAAAAAAAAAAAAAACAAcEFAAB4bC93b3Jrc2hlZXRzL3NoZWV0MS54bWxQSwUGAAAAAAYABgCAAQAAHAcAAAAA';

const decodeFixture = (): Uint8Array =>
  Uint8Array.from(atob(INHERITED_STYLES_XLSX), (char) => char.charCodeAt(0));

describe('formulon 0.13 style projections and workbook themes', () => {
  it('uses native display formatting and resolves an effective cell XF', async () => {
    const wb = await WorkbookHandle.createDefault();
    let reloaded: WorkbookHandle | null = null;
    try {
      expect(wb.capabilities.displayText).toBe(true);
      expect(wb.capabilities.effectiveStyle).toBe(true);
      wb.setNumber({ sheet: 0, row: 0, col: 0 }, 1234.5);
      expect(wb.getDisplayText(0, 0, 0)).toBe('1234.5');
      expect(wb.formatValueForDisplay({ kind: 'number', value: 1234.5 }, '#,##0.00')).toBe(
        '1,234.50',
      );

      const normal = wb.getCellStyleXf(0);
      const font = wb.getFontRecord(0);
      expect(normal).not.toBeNull();
      expect(font).not.toBeNull();
      if (!normal || !font) return;
      const boldFont = wb.addFontRecord({ ...font, bold: true, hasBold: true });
      const xf = wb.addXfRecord({ ...normal, fontIndex: boldFont });
      expect(wb.setCellXfIndex(0, 0, 0, xf)).toBe(true);
      expect(wb.getEffectiveStyle(0, 0, 0)?.source).toBe(0);
      expect(effectiveCellFormatFromEngine(wb, 0, 0, 0)?.bold).toBe(true);

      reloaded = await WorkbookHandle.loadBytes(wb.save());
      expect(reloaded.getDisplayText(0, 0, 0)).toBe('1234.5');
      expect(reloaded.getEffectiveStyle(0, 0, 0)?.source).toBe(0);
      expect(effectiveCellFormatFromEngine(reloaded, 0, 0, 0)?.bold).toBe(true);
    } finally {
      reloaded?.dispose();
      wb.dispose();
    }
  });

  it('projects row and column XFs without flattening them during format sync', async () => {
    const wb = await WorkbookHandle.loadBytes(decodeFixture());
    let reloaded: WorkbookHandle | null = null;
    try {
      // A1 has a stored value and therefore explicitly owns XF 0; the blank
      // C1 cell exposes the row style that would otherwise be inherited.
      expect(wb.getEffectiveStyle(0, 0, 0)?.source).toBe(0);
      expect(wb.getEffectiveStyle(0, 0, 2)?.source).toBe(1);
      expect(wb.getEffectiveStyle(0, 2, 1)?.source).toBe(2);
      expect(effectiveCellFormatFromEngine(wb, 0, 0, 2)).toMatchObject({
        bold: true,
        fill: '#ffff00',
      });
      expect(effectiveCellFormatFromEngine(wb, 0, 2, 1)).toMatchObject({
        bold: true,
        fill: '#ffff00',
      });

      const renderStore = createSpreadsheetStore();
      renderStore.setState((state) => ({
        ...state,
        viewport: { ...state.viewport, rowCount: 3, colCount: 3, widthPx: 240 },
      }));
      const renderState = renderStore.getState();
      const view = { ...renderState, layout: layoutForView(renderState) };
      const fillStyles: string[] = [];
      let fillStyle = '';
      const ctx = {
        get fillStyle(): string {
          return fillStyle;
        },
        set fillStyle(value: string) {
          fillStyle = value;
        },
        fillRect(): void {
          fillStyles.push(fillStyle);
        },
        save(): void {},
        restore(): void {},
        beginPath(): void {},
        rect(): void {},
        clip(): void {},
        fillText(): void {},
        measureText(): TextMetrics {
          return { width: 0 } as TextMetrics;
        },
      } as unknown as CanvasRenderingContext2D;
      const renderTheme = {
        bg: '#ffffff',
        bgElev: '#f5f5f5',
        accentSoft: 'rgba(0,120,212,0.1)',
        fg: '#1f1f1f',
        fgStrong: '#111111',
        fontUi: 'Arial',
        fontMono: 'monospace',
        textCell: 12,
      } as unknown as ResolvedTheme;
      paintCells(
        {
          ctx,
          wb,
          locale: 'en-US',
          getDisplay: undefined,
          sheetBackgroundImage: null,
        },
        view,
        renderTheme,
        buildColLayout(view.layout, view.viewport),
        buildRowLayout(view.layout, view.viewport),
      );
      expect(fillStyles).toContain('#ffff00');

      expect(wb.getCellXfIndex(0, 0, 0)).toBe(0);
      expect(wb.getCellXfIndex(0, 0, 2)).toBe(0);
      expect(wb.getCellXfIndex(0, 2, 1)).toBe(0);

      const store = createSpreadsheetStore();
      wb.attachStore(store);
      store.setState((state) => ({
        ...state,
        format: {
          ...state.format,
          formats: new Map([['0:2:2', { bold: true }]]),
        },
      }));
      syncCellFormatsToEngine(wb, store, 0);
      expect(wb.getCellXfIndex(0, 0, 0)).toBe(0);
      expect(wb.getCellXfIndex(0, 0, 2)).toBe(0);
      expect(wb.getEffectiveStyle(0, 0, 2)?.source).toBe(1);

      reloaded = await WorkbookHandle.loadBytes(wb.save());
      expect(reloaded.getEffectiveStyle(0, 0, 2)?.source).toBe(1);
      expect(reloaded.getEffectiveStyle(0, 2, 1)?.source).toBe(2);
      expect(reloaded.getCellXfIndex(0, 0, 0)).toBe(0);
      expect(reloaded.getCellXfIndex(0, 0, 2)).toBe(0);

      const hydrated = createSpreadsheetStore();
      hydrateCellFormatsFromEngine(reloaded, hydrated, 0);
      expect(hydrated.getState().format.formats.has('0:0:0')).toBe(false);
      expect(hydrated.getState().format.formats.has('0:0:2')).toBe(false);
      expect(hydrated.getState().format.formats.has('0:2:1')).toBe(false);
    } finally {
      reloaded?.dispose();
      wb.dispose();
    }
  });

  it('edits all twelve theme colours and four theme fonts through save/load', async () => {
    const wb = await WorkbookHandle.createDefault();
    let reloaded: WorkbookHandle | null = null;
    const colors = Array.from({ length: 12 }, (_, index) => 0xff000000 + index * 0x00010101);
    const fonts = {
      majorLatin: 'Aptos Display',
      majorEastAsian: 'Yu Gothic UI',
      minorLatin: 'Aptos',
      minorEastAsian: 'Yu Gothic UI',
    };
    try {
      expect(wb.capabilities.workbookTheme).toBe(true);
      expect(wb.getTheme()?.colors).toHaveLength(12);
      expect(wb.setThemeColors(colors)).toBe(true);
      expect(wb.setThemeFonts(fonts)).toBe(true);
      // A theme-linked minor font keeps its raw scheme selector while the
      // display projection follows the edited workbook face.
      expect(wb.setWorkbookDefaultFont({ scheme: 2 })).toBe(true);
      expect(wb.getFontRecord(0)?.name).toBe('Calibri');
      expect(wb.getResolvedFontRecord(0)?.name).toBe(fonts.minorLatin);
      wb.setText({ sheet: 0, row: 0, col: 0 }, 'themed');
      expect(effectiveCellFormatFromEngine(wb, 0, 0, 0)?.fontFamily).toBe(fonts.minorLatin);
      expect(wb.getTheme()).toMatchObject({ colors, fonts });

      reloaded = await WorkbookHandle.loadBytes(wb.save());
      expect(reloaded.getTheme()).toMatchObject({ colors, fonts });
      expect(reloaded.getFontRecord(0)?.scheme).toBe(2);
      expect(reloaded.getResolvedFontRecord(0)?.name).toBe(fonts.minorLatin);
      expect(effectiveCellFormatFromEngine(reloaded, 0, 0, 0)?.fontFamily).toBe(fonts.minorLatin);
    } finally {
      reloaded?.dispose();
      wb.dispose();
    }
  });
});
