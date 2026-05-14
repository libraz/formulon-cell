import { describe, expect, it } from 'vitest';

import {
  activeDraftSide,
  computeDialogNumFmt,
  computeDialogValidation,
  explicitDraftBorders,
  hydrateDraftFromFormat,
  makeEmptyDraft,
  restyleDraftBorders,
  setDraftSide,
} from '../../../src/interact/format-dialog-state.js';
import type { CellFormat } from '../../../src/store/store.js';

describe('interact/format-dialog-state', () => {
  describe('makeEmptyDraft', () => {
    it('starts in the "general" number category with locale-aware currency', () => {
      const en = makeEmptyDraft('en');
      expect(en.numberCategory).toBe('general');
      expect(en.currencySymbol).toBe('$');
      expect(en.validationKind).toBe('none');
      expect(en.locked).toBe(true);

      const ja = makeEmptyDraft('ja');
      expect(ja.currencySymbol).toBe('¥');
    });
  });

  describe('hydrateDraftFromFormat', () => {
    it('maps a currency format into the draft fields', () => {
      const draft = makeEmptyDraft('en');
      const fmt: CellFormat = {
        numFmt: { kind: 'currency', decimals: 2, symbol: '€' },
      };
      hydrateDraftFromFormat(draft, fmt, 'en');
      expect(draft.numberCategory).toBe('currency');
      expect(draft.decimals).toBe(2);
      expect(draft.currencySymbol).toBe('€');
    });

    it('maps a date format and preserves the pattern', () => {
      const draft = makeEmptyDraft('en');
      hydrateDraftFromFormat(draft, { numFmt: { kind: 'date', pattern: 'yyyy-mm-dd' } }, 'en');
      expect(draft.numberCategory).toBe('date');
      expect(draft.pattern).toBe('yyyy-mm-dd');
    });

    it('falls back to general when no numFmt is present', () => {
      const draft = makeEmptyDraft('en');
      hydrateDraftFromFormat(draft, {}, 'en');
      expect(draft.numberCategory).toBe('general');
      expect(draft.numFmt).toEqual({ kind: 'general' });
    });

    it('preserves font flags as booleans', () => {
      const draft = makeEmptyDraft('en');
      hydrateDraftFromFormat(
        draft,
        { bold: true, italic: false, underline: true, strike: false },
        'en',
      );
      expect(draft.bold).toBe(true);
      expect(draft.italic).toBe(false);
      expect(draft.underline).toBe(true);
      expect(draft.strike).toBe(false);
    });

    it('inherits borderStyle / borderColor from the first border side', () => {
      const draft = makeEmptyDraft('en');
      hydrateDraftFromFormat(
        draft,
        {
          borders: {
            top: { style: 'thick', color: '#aabbcc' },
            right: { style: 'thin' },
          },
        },
        'en',
      );
      expect(draft.borderStyle).toBe('thick');
      expect(draft.borderColor).toBe('#aabbcc');
    });
  });

  describe('border draft helpers', () => {
    it('setDraftSide toggles a side on/off', () => {
      const draft = makeEmptyDraft('en');
      draft.borderStyle = 'thick';
      draft.borderColor = '#ff0000';

      const on = setDraftSide(draft, 'top', true);
      expect(on.top).toEqual({ style: 'thick', color: '#ff0000' });

      draft.borders = on;
      const off = setDraftSide(draft, 'top', false);
      expect(off.top).toBe(false);
    });

    it('restyleDraftBorders applies the current active style to already-set sides', () => {
      const draft = makeEmptyDraft('en');
      draft.borders = { top: { style: 'thin' } as never, bottom: false };
      draft.borderStyle = 'thick';
      const next = restyleDraftBorders(draft);
      expect(next.top).toEqual(activeDraftSide(draft));
      expect(next.bottom).toBeUndefined();
    });

    it('explicitDraftBorders surfaces all 6 sides with false fallbacks', () => {
      const draft = makeEmptyDraft('en');
      draft.borders = { top: { style: 'thin' } as never };
      const e = explicitDraftBorders(draft);
      expect(e.top).toBeDefined();
      expect(e.bottom).toBe(false);
      expect(e.left).toBe(false);
      expect(e.right).toBe(false);
      expect(e.diagonalDown).toBe(false);
      expect(e.diagonalUp).toBe(false);
    });
  });

  describe('computeDialogNumFmt', () => {
    const fallback = (cat: string): string => `default-${cat}`;
    it('produces a general kind for "general"', () => {
      const draft = makeEmptyDraft('en');
      draft.numberCategory = 'general';
      expect(computeDialogNumFmt(draft, fallback)).toEqual({ kind: 'general' });
    });

    it('produces a currency kind with symbol + decimals', () => {
      const draft = makeEmptyDraft('en');
      draft.numberCategory = 'currency';
      draft.decimals = 0;
      draft.currencySymbol = '€';
      expect(computeDialogNumFmt(draft, fallback)).toEqual({
        kind: 'currency',
        decimals: 0,
        symbol: '€',
      });
    });

    it('falls back to defaultPatternFor when pattern is empty', () => {
      const draft = makeEmptyDraft('en');
      draft.numberCategory = 'date';
      draft.pattern = '';
      const out = computeDialogNumFmt(draft, fallback);
      expect(out).toEqual({ kind: 'date', pattern: 'default-date' });
    });

    it('keeps the user-supplied pattern when present', () => {
      const draft = makeEmptyDraft('en');
      draft.numberCategory = 'custom';
      draft.pattern = '#,##0.00';
      expect(computeDialogNumFmt(draft, fallback)).toEqual({
        kind: 'custom',
        pattern: '#,##0.00',
      });
    });
  });

  describe('computeDialogValidation', () => {
    it('returns undefined when kind is "none"', () => {
      const draft = makeEmptyDraft('en');
      expect(computeDialogValidation(draft, [])).toBeUndefined();
    });

    it('builds a list validation from inline lines', () => {
      const draft = makeEmptyDraft('en');
      draft.validationKind = 'list';
      draft.validationListSourceKind = 'literal';
      expect(computeDialogValidation(draft, ['A', 'B', 'C'])).toEqual({
        kind: 'list',
        source: ['A', 'B', 'C'],
      });
    });

    it('returns undefined for an empty inline list', () => {
      const draft = makeEmptyDraft('en');
      draft.validationKind = 'list';
      draft.validationListSourceKind = 'literal';
      expect(computeDialogValidation(draft, [])).toBeUndefined();
    });

    it('strips leading = from a range source reference', () => {
      const draft = makeEmptyDraft('en');
      draft.validationKind = 'list';
      draft.validationListSourceKind = 'range';
      draft.validationListRange = '=Sheet1!$A$1:$A$10';
      const out = computeDialogValidation(draft, []);
      expect(out).toEqual({ kind: 'list', source: { ref: 'Sheet1!$A$1:$A$10' } });
    });

    it('builds a bounded numeric rule with between/notBetween a..b', () => {
      const draft = makeEmptyDraft('en');
      draft.validationKind = 'whole';
      draft.validationOp = 'between';
      draft.validationA = 1;
      draft.validationB = 10;
      expect(computeDialogValidation(draft, [])).toEqual({
        kind: 'whole',
        op: 'between',
        a: 1,
        b: 10,
      });

      draft.validationOp = '>';
      expect(computeDialogValidation(draft, [])).toEqual({
        kind: 'whole',
        op: '>',
        a: 1,
      });
    });

    it('threads errorStyle + allowBlank meta when non-default', () => {
      const draft = makeEmptyDraft('en');
      draft.validationKind = 'whole';
      draft.validationOp = '=';
      draft.validationA = 5;
      draft.validationAllowBlank = false;
      draft.validationErrorStyle = 'warning';
      expect(computeDialogValidation(draft, [])).toEqual({
        kind: 'whole',
        op: '=',
        a: 5,
        allowBlank: false,
        errorStyle: 'warning',
      });
    });
  });
});
