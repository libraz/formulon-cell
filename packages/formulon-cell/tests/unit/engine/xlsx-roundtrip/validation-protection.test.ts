import { describe, expect, it } from 'vitest';
import { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';

import { canLoadWasm } from './fixtures.js';

describe.skipIf(!canLoadWasm())('real xlsx round-trip', () => {
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

  it('saves and reloads data-validation hidden dropdown visibility', async () => {
    const first = await WorkbookHandle.createDefault();

    try {
      expect(first.isStub).toBe(false);
      if (!first.capabilities.dataValidation) return;

      expect(
        first.addValidationEntry(0, {
          type: 3,
          ranges: [{ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }],
          formula1: '"Yes,No"',
          showDropDown: true,
        }),
      ).toBe(true);

      const bytes = first.save();
      expect(bytes.length).toBeGreaterThan(0);

      const reloaded = await WorkbookHandle.loadBytes(bytes);
      try {
        expect(reloaded.isStub).toBe(false);
        const validations = reloaded.getValidationsForSheet(0);
        expect(validations).toHaveLength(1);
        const validation = validations[0];
        expect(validation).toBeDefined();
        expect(validation).toMatchObject({
          type: 3,
          formula1: '"Yes,No"',
          ranges: [{ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }],
        });
        expect(validation?.showDropDown).toBe(true);
      } finally {
        reloaded.dispose();
      }
    } finally {
      first.dispose();
    }
  });
});
