import { describe, expect, it } from 'vitest';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';

const canLoadWasm = (): boolean =>
  typeof WebAssembly !== 'undefined' && typeof SharedArrayBuffer !== 'undefined';

describe.skipIf(!canLoadWasm())('real xlsx round-trip', () => {
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

  it('saves and reloads workbook metadata supported by the real engine', async () => {
    const first = await WorkbookHandle.createDefault();

    try {
      expect(first.isStub).toBe(false);

      if (first.capabilities.definedNameMutate) {
        expect(first.setDefinedNameEntry('RoundTripName', 'Sheet1!$A$1')).toBe(true);
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
      }

      const bytes = first.save();
      expect(bytes.length).toBeGreaterThan(0);

      const reloaded = await WorkbookHandle.loadBytes(bytes);
      try {
        expect(reloaded.isStub).toBe(false);

        if (first.capabilities.definedNameMutate) {
          expect([...reloaded.definedNames()]).toContainEqual({
            name: 'RoundTripName',
            formula: 'Sheet1!$A$1',
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
        }
      } finally {
        reloaded.dispose();
      }
    } finally {
      first.dispose();
    }
  });

  it('saves and reloads validation and sheet-protection metadata when supported', async () => {
    const first = await WorkbookHandle.createDefault();

    try {
      expect(first.isStub).toBe(false);

      if (first.capabilities.dataValidation) {
        expect(
          first.addValidationEntry(0, {
            type: 3,
            ranges: [{ sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 }],
            formula1: '"Yes,No"',
            allowBlank: false,
            showErrorMessage: false,
            showDropDown: true,
          }),
        ).toBe(true);
      }
      if (first.capabilities.sheetProtectionRoundtrip) {
        expect(
          first.setSheetProtection(0, {
            enabled: true,
            legacyPassword: 'ABCD',
            sheet: true,
            selectLockedCells: true,
            selectUnlockedCells: true,
            sort: true,
            autoFilter: true,
          }),
        ).toBe(true);
      }

      const bytes = first.save();
      expect(bytes.length).toBeGreaterThan(0);

      const reloaded = await WorkbookHandle.loadBytes(bytes);
      try {
        expect(reloaded.isStub).toBe(false);

        if (first.capabilities.dataValidation) {
          const validations = reloaded.getValidationsForSheet(0);
          expect(validations).toHaveLength(1);
          expect(validations[0]).toMatchObject({
            type: 3,
            formula1: '"Yes,No"',
            allowBlank: false,
            showErrorMessage: false,
            ranges: [{ sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 }],
          });
        }
        if (first.capabilities.sheetProtectionRoundtrip) {
          expect(reloaded.getSheetProtection(0)).toMatchObject({
            enabled: true,
            legacyPassword: 'ABCD',
            sheet: true,
            selectLockedCells: true,
            selectUnlockedCells: true,
            sort: true,
            autoFilter: true,
          });
        }
      } finally {
        reloaded.dispose();
      }
    } finally {
      first.dispose();
    }
  });
});
