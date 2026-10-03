import { describe, expect, it } from 'vitest';
import { setSheetRightToLeft } from '../../../../src/commands/view.js';
import { hydrateLayoutFromEngine } from '../../../../src/engine/layout-sync.js';
import { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore } from '../../../../src/store/store.js';

import { canLoadWasm } from './fixtures.js';

describe.skipIf(!canLoadWasm())('real xlsx round-trip', () => {
  it('exposes the 0.12.0 function catalog and recalculates USDOLLAR through the adapter', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      expect(wb.sheetCount).toBe(1);
      expect(wb.functionNames()).toContain('USDOLLAR');
      wb.setFormula({ sheet: 0, row: 0, col: 0 }, '=USDOLLAR(-1234.5)');
      expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
        kind: 'text',
        value: '($1,234.50)',
      });
    } finally {
      wb.dispose();
    }
  });

  it('saves and reloads values and formulas through the real engine', async () => {
    const first = await WorkbookHandle.createDefault();

    try {
      first.setNumber({ sheet: 0, row: 0, col: 0 }, 40);
      first.setNumber({ sheet: 0, row: 0, col: 1 }, 2);
      first.setFormula({ sheet: 0, row: 0, col: 2 }, '=A1+B1');
      first.recalc();

      const bytes = first.save();
      expect(bytes.length).toBeGreaterThan(0);

      const cleared = await WorkbookHandle.loadBytes(bytes);
      try {
        expect(cleared.isStub).toBe(false);
        expect(cleared.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
          kind: 'number',
          value: 40,
        });
        expect(cleared.cellFormula({ sheet: 0, row: 0, col: 2 })).toBe('=A1+B1');
        cleared.recalc();
        expect(cleared.getValue({ sheet: 0, row: 0, col: 2 })).toEqual({
          kind: 'number',
          value: 42,
        });
      } finally {
        cleared.dispose();
      }
    } finally {
      first.dispose();
    }
  });

  it('round-trips sheet display flags through the real engine', async () => {
    const first = await WorkbookHandle.createDefault();

    try {
      expect(first.isStub).toBe(false);
      if (!first.capabilities.sheetViewFlags) return;

      expect(first.setSheetShowGridLines(0, false)).toBe(true);
      expect(first.setSheetShowRowColHeaders(0, false)).toBe(true);
      expect(first.setSheetShowZeros(0, false)).toBe(true);
      expect(first.setSheetRightToLeft(0, true)).toBe(true);

      const reloaded = await WorkbookHandle.loadBytes(first.save());
      try {
        expect(reloaded.getSheetView(0)).toMatchObject({
          showGridLines: false,
          showRowColHeaders: false,
          showZeros: false,
          rightToLeft: true,
        });
      } finally {
        reloaded.dispose();
      }
    } finally {
      first.dispose();
    }
  });

  it('round-trips the sheet direction into the store and back out', async () => {
    const first = await WorkbookHandle.createDefault();

    try {
      if (!first.capabilities.sheetViewFlags) return;
      const store = createSpreadsheetStore();
      expect(store.getState().ui.rightToLeft).toBe(false);

      setSheetRightToLeft(store, true, first);
      expect(store.getState().ui.rightToLeft).toBe(true);

      const reloaded = await WorkbookHandle.loadBytes(first.save());
      try {
        const restored = createSpreadsheetStore();
        hydrateLayoutFromEngine(reloaded, restored, 0);
        expect(restored.getState().ui.rightToLeft).toBe(true);
      } finally {
        reloaded.dispose();
      }
    } finally {
      first.dispose();
    }
  });

  it('saves and reloads workbook metadata supported by the real engine', async () => {
    const first = await WorkbookHandle.createDefault();

    try {
      expect(first.isStub).toBe(false);

      if (first.capabilities.definedNameMutate) {
        first.setNumber({ sheet: 0, row: 0, col: 0 }, 4);
        first.setNumber({ sheet: 0, row: 0, col: 1 }, 7);
        expect(first.setDefinedNameEntry('RoundTripName', 'Sheet1!$A$1')).toBe(true);
        expect(first.setDefinedNameEntry('LocalRoundTripName', 'Sheet1!$B$1', 0)).toBe(true);
        first.setFormula({ sheet: 0, row: 0, col: 3 }, '=RoundTripName');
        first.setFormula({ sheet: 0, row: 0, col: 4 }, '=LocalRoundTripName');
        first.recalc();
        expect(first.getValue({ sheet: 0, row: 0, col: 3 })).toEqual({ kind: 'number', value: 4 });
        expect(first.getValue({ sheet: 0, row: 0, col: 4 })).toEqual({ kind: 'number', value: 7 });
      }
      if (first.capabilities.hyperlinks) {
        expect(
          first.addHyperlink(
            0,
            1,
            1,
            'https://example.com/roundtrip',
            'Round Trip',
            'Open round-trip link',
          ),
        ).toBe(true);
      }
      if (first.capabilities.comments) {
        expect(first.setCommentEntry(0, 2, 2, 'Formulon', 'Round-trip comment')).toBe(true);
        expect(first.setCommentEntry(0, 4, 4, 'Formulon', 'Blank-cell comment')).toBe(true);
      }

      const bytes = first.save();
      expect(bytes.length).toBeGreaterThan(0);

      const reloaded = await WorkbookHandle.loadBytes(bytes);
      try {
        expect(reloaded.isStub).toBe(false);

        if (first.capabilities.definedNameMutate) {
          const names = [...reloaded.definedNames()];
          expect(names).toContainEqual({
            name: 'RoundTripName',
            formula: 'Sheet1!$A$1',
            localSheetId: -1,
          });
          expect(names).toContainEqual({
            name: 'LocalRoundTripName',
            formula: 'Sheet1!$B$1',
            localSheetId: 0,
          });
          reloaded.recalc();
          expect(reloaded.getValue({ sheet: 0, row: 0, col: 3 })).toEqual({
            kind: 'number',
            value: 4,
          });
          expect(reloaded.getValue({ sheet: 0, row: 0, col: 4 })).toEqual({
            kind: 'number',
            value: 7,
          });
        }
        if (first.capabilities.hyperlinks) {
          expect(reloaded.getHyperlinks(0)).toContainEqual({
            row: 1,
            col: 1,
            target: 'https://example.com/roundtrip',
            display: 'Round Trip',
            tooltip: 'Open round-trip link',
          });
        }
        if (first.capabilities.comments) {
          expect(reloaded.getComment(0, 2, 2)).toEqual({
            author: 'Formulon',
            text: 'Round-trip comment',
          });
          if (reloaded.capabilities.commentsEnumerable) {
            expect(reloaded.getComments(0)).toContainEqual({
              row: 4,
              col: 4,
              author: 'Formulon',
              text: 'Blank-cell comment',
            });
          }
        }
      } finally {
        reloaded.dispose();
      }
    } finally {
      first.dispose();
    }
  });

  it('saves and reloads static error values when supported', async () => {
    const first = await WorkbookHandle.createDefault();

    try {
      expect(first.isStub).toBe(false);
      if (!first.capabilities.staticErrorValues) return;

      first.setError({ sheet: 0, row: 0, col: 0 }, 1);
      first.setError({ sheet: 0, row: 1, col: 0 }, 5);

      const bytes = first.save();
      expect(bytes.length).toBeGreaterThan(0);

      const reloaded = await WorkbookHandle.loadBytes(bytes);
      try {
        expect(reloaded.isStub).toBe(false);
        expect(reloaded.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
          kind: 'error',
          code: 1,
          text: '#DIV/0!',
        });
        expect(reloaded.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({
          kind: 'error',
          code: 5,
          text: '#NUM!',
        });
      } finally {
        reloaded.dispose();
      }
    } finally {
      first.dispose();
    }
  });
});
