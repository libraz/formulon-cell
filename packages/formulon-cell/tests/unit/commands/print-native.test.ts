import { describe, expect, it } from 'vitest';
import { buildPrintDocument } from '../../../src/commands/print.js';
import { syncPageSetupToEngine } from '../../../src/engine/print-sync.js';
import { addrKey, WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

describe('native pagination print integration', () => {
  it('renders native page placements in order and uses native R1C1 formulas', async () => {
    const wb = await WorkbookHandle.createDefault();
    const store = createSpreadsheetStore();
    try {
      for (let row = 0; row < 140; row += 1) {
        wb.setNumber({ sheet: 0, row, col: 0 }, row + 1);
        store.setState((s) => {
          const cells = new Map(s.data.cells);
          cells.set(addrKey({ sheet: 0, row, col: 0 }), {
            value: { kind: 'number', value: row + 1 },
            formula: null,
          });
          return { ...s, data: { ...s.data, cells } };
        });
      }
      wb.setNumber({ sheet: 0, row: 2, col: 2 }, 3);
      wb.setFormula({ sheet: 0, row: 4, col: 0 }, '=A3');
      store.setState((s) => {
        const cells = new Map(s.data.cells);
        cells.set(addrKey({ sheet: 0, row: 2, col: 2 }), {
          value: { kind: 'number', value: 3 },
          formula: null,
        });
        cells.set(addrKey({ sheet: 0, row: 4, col: 0 }), {
          value: { kind: 'number', value: 3 },
          formula: '=A3',
        });
        return {
          ...s,
          data: { ...s.data, cells },
          ui: { ...s.ui, showFormulas: true, r1c1: true },
        };
      });
      mutators.setPageSetup(store, 0, {
        printArea: 'A1:A140',
        printTitleRows: '1:1',
        manualPageBreakRows: [60],
        scale: 0.5,
      });
      syncPageSetupToEngine(wb, store, 0);

      const native = wb.paginateSheet(0);
      expect(native?.pages.length).toBeGreaterThan(1);
      const document = buildPrintDocument(wb, store, 0);
      expect((document.html.match(/<section class="fc-print__area/g) ?? []).length).toBe(
        native?.pages.length,
      );
      expect(document.html).toContain('fc-print__native-page');
      expect(document.html).toContain(`data-native-origin-x-pt="${native?.pages[0]?.originXPt}"`);
      expect(document.html).toContain(`data-native-width-pt="${native?.pages[0]?.widthPt}"`);
      const titledPage = native?.pages.find(
        (page) => page.originYPt > (native?.printable.y ?? Number.POSITIVE_INFINITY),
      );
      expect(titledPage).toBeDefined();
      expect(document.html).toContain(`data-native-origin-y-pt="${titledPage?.originYPt}"`);
      expect(document.html).toContain(`data-native-height-pt="${titledPage?.heightPt}"`);
      expect(document.html).toContain(
        `@page { size: ${native?.paper.widthPt}pt ${native?.paper.heightPt}pt; margin: ${native?.margins.top}pt`,
      );
      expect(document.html).toContain(`height:${native?.printable.height}pt;overflow:hidden`);
      expect(document.html).toContain('table-layout:fixed');
      expect(document.html).not.toContain('transform:scale(0.5)');
      expect(document.html).toContain(
        `width:${native?.pages[0]?.widthPt}pt;height:${native?.pages[0]?.heightPt}pt;table-layout:fixed`,
      );
      for (const row of [1, 60, 61, 104, 139, 140]) {
        expect(document.html).toContain(`>${row}<`);
      }
      expect(document.html).toContain('fc-print__native-title-spacer');
      expect(document.html).toContain('overflow:hidden;break-inside:avoid');
      expect(document.html).toContain('=R[-2]C');
      expect(document.cssVars['--fc-print-scale']).toBe('0.5');
    } finally {
      wb.dispose();
    }
  });
});
