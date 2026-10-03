import { describe, expect, it } from 'vitest';
import { planSelectionFormat } from '../../../src/commands/format.js';
import {
  buildTouchedDialogPatch,
  computeDialogNumFmt,
  computeDialogValidation,
  explicitDraftBorders,
  hydrateDraftFromFormat,
  makeEmptyDraft,
  setDraftSide,
  summarizeDialogFormats,
} from '../../../src/interact/format-dialog-state.js';
import type { CellFormat, SpreadsheetStore } from '../../../src/store/store.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

describe('interact/format-dialog-state', () => {
  it('summarizes mixed fields and emits only touched draft fields', () => {
    const store: SpreadsheetStore = createSpreadsheetStore();
    store.setState((state) => ({
      ...state,
      selection: {
        ...state.selection,
        active: { sheet: 0, row: 0, col: 0 },
        anchor: { sheet: 0, row: 0, col: 0 },
        range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
        extraRanges: [],
      },
    }));
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      {
        bold: true,
        hyperlink: 'https://active.test',
      },
    );
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 1 },
      {
        bold: false,
        hyperlink: 'https://other.test',
      },
    );
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      {
        validation: { kind: 'list', source: ['A', 'B'] },
      },
    );
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 1 },
      {
        validation: { source: ['A', 'B'], kind: 'list' },
      },
    );
    const plan = planSelectionFormat(store.getState());
    if (!plan) throw new Error('selection plan missing');
    const summary = summarizeDialogFormats(store.getState(), plan);
    expect(summary.activeFormat.bold).toBe(true);
    expect(summary.mixed.has('bold')).toBe(true);
    expect(summary.mixed.has('hyperlink')).toBe(true);
    expect(summary.mixed.has('validation')).toBe(false);

    const draft = makeEmptyDraft('en');
    hydrateDraftFromFormat(draft, summary.activeFormat, 'en');
    draft.align = 'center';
    const touched = new Set(['align'] as const);
    expect(buildTouchedDialogPatch(draft, touched)).toEqual({ align: 'center' });
  });

  describe('makeEmptyDraft', () => {
    it('starts in the "general" number category with locale-aware currency', () => {
      const en = makeEmptyDraft('en');
      expect(en.numberCategory).toBe('general');
      expect(en.currencySymbol).toBe('$');
      expect(en.validationKind).toBe('none');
      expect(en.validationShowDropdown).toBe(true);
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

    it('maps known fraction custom formats into the fraction category', () => {
      const draft = makeEmptyDraft('en');
      hydrateDraftFromFormat(draft, { numFmt: { kind: 'custom', pattern: '# ?/16' } }, 'en');
      expect(draft.numberCategory).toBe('fraction');
      expect(draft.pattern).toBe('# ?/16');
    });

    it('maps combined datetime formats into Custom without changing the pattern', () => {
      const draft = makeEmptyDraft('en');
      hydrateDraftFromFormat(
        draft,
        { numFmt: { kind: 'datetime', pattern: 'yyyy-mm-dd HH:MM' } },
        'en',
      );
      expect(draft.numberCategory).toBe('custom');
      expect(draft.pattern).toBe('yyyy-mm-dd HH:MM');
    });

    it('falls back to general when no numFmt is present', () => {
      const draft = makeEmptyDraft('en');
      hydrateDraftFromFormat(draft, {}, 'en');
      expect(draft.numberCategory).toBe('general');
      expect(draft.numFmt).toEqual({ kind: 'general' });
    });

    it('preserves font flags and normalizes legacy underline to single', () => {
      const draft = makeEmptyDraft('en');
      hydrateDraftFromFormat(
        draft,
        { bold: true, italic: false, underline: true, strike: false },
        'en',
      );
      expect(draft.bold).toBe(true);
      expect(draft.italic).toBe(false);
      expect(draft.underline).toBe('single');
      expect(draft.strike).toBe(false);
    });

    it('preserves superscript and subscript font alignment', () => {
      const draft = makeEmptyDraft('en');
      hydrateDraftFromFormat(draft, { fontVertAlign: 'superscript' }, 'en');
      expect(draft.fontVertAlign).toBe('superscript');
      hydrateDraftFromFormat(draft, { fontVertAlign: 'subscript' }, 'en');
      expect(draft.fontVertAlign).toBe('subscript');
    });

    it('hydrates the distributed-alignment final-line setting', () => {
      const draft = makeEmptyDraft('en');
      hydrateDraftFromFormat(draft, { align: 'distributed', justifyLastLine: true }, 'en');
      expect(draft.justifyLastLine).toBe(true);
    });

    it('preserves the full underline style', () => {
      const draft = makeEmptyDraft('en');
      hydrateDraftFromFormat(draft, { underline: 'doubleAccounting' }, 'en');
      expect(draft.underline).toBe('doubleAccounting');
    });

    it.each([
      'thin',
      'medium',
      'thick',
      'dashed',
      'dotted',
      'double',
      'hair',
      'mediumDashed',
      'dashDot',
      'mediumDashDot',
      'dashDotDot',
      'mediumDashDotDot',
      'slantDashDot',
    ] as const)('inherits %s borderStyle without collapsing it', (style) => {
      const draft = makeEmptyDraft('en');
      hydrateDraftFromFormat(
        draft,
        {
          borders: {
            top: { style, color: '#aabbcc' },
            right: { style: 'thin' },
          },
        },
        'en',
      );
      expect(draft.borderStyle).toBe(style);
      expect(draft.borderColor).toBe('#aabbcc');
    });

    it('hydrates validation input-message and error-alert metadata', () => {
      const draft = makeEmptyDraft('en');
      hydrateDraftFromFormat(
        draft,
        {
          validation: {
            kind: 'whole',
            op: '=',
            a: 5,
            showInputMessage: false,
            promptTitle: 'Number',
            promptMessage: 'Enter five.',
            showErrorMessage: false,
            errorTitle: 'Wrong number',
            errorMessage: 'Only five is allowed.',
          },
        },
        'en',
      );

      expect(draft.validationKind).toBe('whole');
      expect(draft.validationShowInputMessage).toBe(false);
      expect(draft.validationPromptTitle).toBe('Number');
      expect(draft.validationPromptMessage).toBe('Enter five.');
      expect(draft.validationShowErrorMessage).toBe(false);
      expect(draft.validationErrorTitle).toBe('Wrong number');
      expect(draft.validationErrorMessage).toBe('Only five is allowed.');
    });

    it('hydrates hyperlink display and tooltip metadata', () => {
      const draft = makeEmptyDraft('en');
      hydrateDraftFromFormat(
        draft,
        {
          hyperlink: 'https://example.test',
          hyperlinkDisplay: 'Example',
          hyperlinkTooltip: 'Open example',
        },
        'en',
      );

      expect(draft.hyperlink).toBe('https://example.test');
      expect(draft.originalHyperlink).toBe('https://example.test');
      expect(draft.hyperlinkDisplay).toBe('Example');
      expect(draft.hyperlinkTooltip).toBe('Open example');
    });

    it('hydrates hidden list-dropdown metadata', () => {
      const draft = makeEmptyDraft('en');
      hydrateDraftFromFormat(
        draft,
        {
          validation: {
            kind: 'list',
            source: ['A', 'B'],
            showDropdown: false,
          },
        },
        'en',
      );

      expect(draft.validationKind).toBe('list');
      expect(draft.validationShowDropdown).toBe(false);
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

    it('uses one shared style for both diagonal directions', () => {
      const draft = makeEmptyDraft('en');
      draft.borderStyle = 'double';
      draft.borderColor = '#ff0000';
      const both = setDraftSide(draft, 'diagonalDown', true);
      expect(both.diagonalDown).toEqual({ style: 'double', color: '#ff0000' });
      expect(both.diagonalUp).toEqual({ style: 'double', color: '#ff0000' });
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

    it('emits the selected fraction preset as a custom number format', () => {
      const draft = makeEmptyDraft('en');
      draft.numberCategory = 'fraction';
      draft.pattern = '# ??/??';
      expect(computeDialogNumFmt(draft, fallback)).toEqual({
        kind: 'custom',
        pattern: '# ??/??',
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

    it('threads input-message and error-alert metadata', () => {
      const draft = makeEmptyDraft('en');
      draft.validationKind = 'list';
      draft.validationShowInputMessage = false;
      draft.validationPromptTitle = 'Pick a status';
      draft.validationPromptMessage = 'Use one of the approved workflow states.';
      draft.validationShowErrorMessage = false;
      draft.validationErrorTitle = 'Invalid status';
      draft.validationErrorMessage = 'Choose Open, Closed, or Hold.';

      expect(computeDialogValidation(draft, ['Open', 'Closed', 'Hold'])).toEqual({
        kind: 'list',
        source: ['Open', 'Closed', 'Hold'],
        showInputMessage: false,
        promptTitle: 'Pick a status',
        promptMessage: 'Use one of the approved workflow states.',
        showErrorMessage: false,
        errorTitle: 'Invalid status',
        errorMessage: 'Choose Open, Closed, or Hold.',
      });
    });

    it('preserves hidden list-dropdown metadata when rebuilding validation', () => {
      const draft = makeEmptyDraft('en');
      draft.validationKind = 'list';
      draft.validationShowDropdown = false;

      expect(computeDialogValidation(draft, ['A', 'B'])).toEqual({
        kind: 'list',
        source: ['A', 'B'],
        showDropdown: false,
      });
    });
  });
});
