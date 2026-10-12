import { describe, expect, it, vi } from 'vitest';
import { writeInputValidated } from '../../../src/commands/coerce-input.js';
import { listInvalidValidationCells } from '../../../src/commands/validate.js';
import { addrKey } from '../../../src/engine/address.js';
import {
  hydrateValidationsFromEngine,
  syncValidationsToEngine,
} from '../../../src/engine/validation-sync.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

describe('formulon 0.13 native filter and validation surfaces', () => {
  it('uses the upstream comparison ordinals for all strict bounds', async () => {
    const wb = await WorkbookHandle.createDefault();
    const store = createSpreadsheetStore();
    try {
      expect(wb.capabilities.validationEvaluation).toBe(true);
      const rules = [
        { col: 0, op: '>' as const, ordinal: 4, good: 11, bad: 10 },
        { col: 1, op: '<' as const, ordinal: 5, good: 9, bad: 10 },
        { col: 2, op: '>=' as const, ordinal: 6, good: 10, bad: 9 },
        { col: 3, op: '<=' as const, ordinal: 7, good: 10, bad: 11 },
      ];
      store.setState((state) => ({
        ...state,
        format: {
          ...state.format,
          formats: new Map(
            rules.map((rule) => [
              addrKey({ sheet: 0, row: 0, col: rule.col }),
              {
                validation: {
                  kind: 'whole' as const,
                  op: rule.op,
                  a: 10,
                  allowBlank: true,
                  showErrorMessage: true,
                },
              },
            ]),
          ),
        },
      }));
      syncValidationsToEngine(wb, store, 0);
      expect(wb.getValidationsForSheet(0).map((entry) => entry.op)).toEqual([4, 5, 6, 7]);
      for (const rule of rules) {
        expect(
          wb.validateValue(0, 0, rule.col, { kind: 'number', value: rule.good }),
        ).toMatchObject({
          hasRule: true,
          valid: true,
        });
        expect(wb.validateValue(0, 0, rule.col, { kind: 'number', value: rule.bad })).toMatchObject(
          {
            hasRule: true,
            valid: false,
          },
        );
      }

      hydrateValidationsFromEngine(wb, store, 0);
      for (const rule of rules) {
        expect(store.getState().format.formats.get(`0:0:${rule.col}`)?.validation).toMatchObject({
          kind: 'whole',
          op: rule.op,
          a: 10,
        });
      }

      const validateSpy = vi.spyOn(wb, 'validateValue');
      const validation = {
        kind: 'whole' as const,
        op: '>' as const,
        a: 10,
        allowBlank: true,
        showErrorMessage: true,
      };
      expect(
        writeInputValidated(wb, { sheet: 0, row: 0, col: 0 }, '11', validation, store),
      ).toEqual({ ok: true });
      expect(validateSpy).toHaveBeenCalledOnce();

      const restored = await WorkbookHandle.loadBytes(wb.save());
      try {
        expect(restored.getValidationsForSheet(0).map((entry) => entry.op)).toEqual([4, 5, 6, 7]);
      } finally {
        restored.dispose();
      }
    } finally {
      wb.dispose();
    }
  });

  it('round-trips the native user-facing list dropdown visibility', async () => {
    const wb = await WorkbookHandle.createDefault();
    const store = createSpreadsheetStore();
    try {
      expect(wb.capabilities.dataValidation).toBe(true);
      store.setState((state) => ({
        ...state,
        format: {
          ...state.format,
          formats: new Map([
            [
              '0:0:0',
              { validation: { kind: 'list' as const, source: ['Yes', 'No'], showDropdown: false } },
            ],
          ]),
        },
      }));
      syncValidationsToEngine(wb, store, 0);
      expect(wb.getValidationsForSheet(0)[0]?.showDropDown).toBe(false);
      const restored = await WorkbookHandle.loadBytes(wb.save());
      try {
        expect(restored.getValidationsForSheet(0)[0]?.showDropDown).toBe(false);
        const restoredStore = createSpreadsheetStore();
        hydrateValidationsFromEngine(restored, restoredStore, 0);
        expect(restoredStore.getState().format.formats.get('0:0:0')?.validation).toEqual({
          kind: 'list',
          source: ['Yes', 'No'],
          showDropdown: false,
          showInputMessage: true,
        });
      } finally {
        restored.dispose();
      }
    } finally {
      wb.dispose();
    }
  });

  it('keeps custom validation formulas in the raw no-equals API form', async () => {
    const wb = await WorkbookHandle.createDefault();
    const store = createSpreadsheetStore();
    try {
      store.setState((state) => ({
        ...state,
        format: {
          ...state.format,
          formats: new Map([
            ['0:0:0', { validation: { kind: 'custom' as const, formula: '=A1>10' } }],
          ]),
        },
      }));
      syncValidationsToEngine(wb, store, 0);
      expect(wb.getValidationsForSheet(0)[0]?.formula1).toBe('A1>10');
      const restored = await WorkbookHandle.loadBytes(wb.save());
      try {
        expect(restored.getValidationsForSheet(0)[0]?.formula1).toBe('A1>10');
        const restoredStore = createSpreadsheetStore();
        hydrateValidationsFromEngine(restored, restoredStore, 0);
        expect(restoredStore.getState().format.formats.get('0:0:0')?.validation).toEqual({
          kind: 'custom',
          formula: 'A1>10',
          showInputMessage: true,
        });
      } finally {
        restored.dispose();
      }
    } finally {
      wb.dispose();
    }
  });

  it('checks proposed bounded values through validateValue', async () => {
    const wb = await WorkbookHandle.createDefault();
    const store = createSpreadsheetStore();
    try {
      expect(wb.capabilities.validationEvaluation).toBe(true);
      expect(
        wb.addValidationEntry(0, {
          ranges: [{ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }],
          type: 1,
          op: 0,
          errorStyle: 0,
          allowBlank: true,
          showErrorMessage: true,
          formula1: '1',
          formula2: '10',
        }),
      ).toBe(true);
      expect(wb.validateValue(0, 0, 0, { kind: 'number', value: 5 })).toMatchObject({
        hasRule: true,
        valid: true,
        errorStyle: 0,
      });
      expect(wb.validateValue(0, 0, 0, { kind: 'number', value: 15 })).toMatchObject({
        hasRule: true,
        valid: false,
      });

      const validation = {
        kind: 'whole' as const,
        op: 'between' as const,
        a: 1,
        b: 10,
        allowBlank: true,
        showErrorMessage: true,
      };
      expect(writeInputValidated(wb, { sheet: 0, row: 0, col: 0 }, '5', validation, store)).toEqual(
        { ok: true },
      );
      expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'number', value: 5 });
      expect(
        writeInputValidated(wb, { sheet: 0, row: 0, col: 0 }, '15', validation, store),
      ).toMatchObject({ ok: false, severity: 'stop' });
      expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'number', value: 5 });

      // A store rule changed before the native mirror catches up must use the
      // JavaScript path; the stale engine rule allows blanks, while this one does not.
      const changedValidation = { ...validation, allowBlank: false };
      expect(
        writeInputValidated(wb, { sheet: 0, row: 0, col: 0 }, '', changedValidation, store),
      ).toMatchObject({ ok: false, severity: 'stop' });
    } finally {
      wb.dispose();
    }
  });

  it('pages invalid populated cells for the existing circle-invalid command', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      expect(wb.capabilities.validationEvaluation).toBe(true);
      expect(
        wb.addValidationEntry(0, {
          ranges: [{ sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 }],
          type: 1,
          op: 0,
          formula1: '1',
          formula2: '10',
        }),
      ).toBe(true);
      wb.setNumber({ sheet: 0, row: 0, col: 0 }, 15);
      wb.setNumber({ sheet: 0, row: 1, col: 0 }, 5);
      expect(listInvalidValidationCells(wb, 0)).toEqual([{ sheet: 0, row: 0, col: 0 }]);
    } finally {
      wb.dispose();
    }
  });

  it('falls back when an invalid-cell pager makes no progress', () => {
    const malformed = {
      capabilities: { validationEvaluation: true },
      listInvalidCells: () => ({ cells: [], nextCursor: 1 }),
    } as unknown as WorkbookHandle;
    expect(listInvalidValidationCells(malformed, 0)).toBeNull();
  });

  it('uses native proposed-value validation for synchronized range-backed lists', async () => {
    const wb = await WorkbookHandle.createDefault();
    const store = createSpreadsheetStore();
    try {
      expect(wb.capabilities.validationEvaluation).toBe(true);
      wb.setText({ sheet: 0, row: 0, col: 1 }, 'Allowed');
      wb.setText({ sheet: 0, row: 1, col: 1 }, 'No');
      expect(
        wb.addValidationEntry(0, {
          ranges: [{ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }],
          type: 3,
          allowBlank: true,
          showErrorMessage: true,
          // formulon follows the OOXML contract: range-backed list formulas
          // are passed without the optional leading `=`.
          formula1: 'B1:B2',
        }),
      ).toBe(true);
      expect(wb.validateValue(0, 0, 0, { kind: 'text', value: 'Allowed' })).toMatchObject({
        hasRule: true,
        valid: true,
      });
      const validation = {
        kind: 'list' as const,
        source: { ref: 'B1:B2' },
        showErrorMessage: true,
      };
      const outcome = writeInputValidated(
        wb,
        { sheet: 0, row: 0, col: 0 },
        'Allowed',
        validation,
        store,
      );
      expect(outcome).toEqual({ ok: true });
      expect(
        writeInputValidated(wb, { sheet: 0, row: 0, col: 0 }, 'Other', validation, store),
      ).toMatchObject({ ok: false, severity: 'stop' });

      // A stale engine source must never make the native answer authoritative.
      wb.setText({ sheet: 0, row: 2, col: 1 }, 'Other');
      const staleValidation = { ...validation, source: { ref: 'B1:B3' } };
      expect(
        writeInputValidated(wb, { sheet: 0, row: 0, col: 0 }, 'Other', staleValidation, store),
      ).toEqual({ ok: true });

      wb.setNumber({ sheet: 0, row: 0, col: 0 }, 20);
      const custom = {
        kind: 'custom' as const,
        formula: '=A1>10',
        showErrorMessage: true,
      };
      expect(
        writeInputValidated(wb, { sheet: 0, row: 0, col: 0 }, '5', custom, store),
      ).toMatchObject({ ok: false, severity: 'stop' });
    } finally {
      wb.dispose();
    }
  });
});
