import { describe, expect, it } from 'vitest';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';

const canLoadWasm = (): boolean => typeof WebAssembly !== 'undefined';

describe.skipIf(!canLoadWasm())('workbook calc settings', () => {
  it('reads the desktop defaults on a new workbook', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      expect(wb.calcSettings()).toEqual({
        fullPrecision: true,
        calcOnSave: true,
        r1c1: false,
        date1904: false,
      });
    } finally {
      wb.dispose();
    }
  });

  it('round-trips every setting through save and load', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      expect(wb.setCalcSetting('fullPrecision', false)).toBe(true);
      expect(wb.setCalcSetting('calcOnSave', false)).toBe(true);
      expect(wb.setCalcSetting('r1c1', true)).toBe(true);
      expect(wb.setCalcSetting('date1904', true)).toBe(true);
      expect(wb.setCalcMode(2)).toBe(true);
      const restored = await WorkbookHandle.loadBytes(wb.save());
      try {
        expect(restored.calcSettings()).toEqual({
          fullPrecision: false,
          calcOnSave: false,
          r1c1: true,
          date1904: true,
        });
        expect(restored.calcMode()).toBe(2);
      } finally {
        restored.dispose();
      }
    } finally {
      wb.dispose();
    }
  });

  it('switches the date system without converting stored serials', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      wb.setNumber({ sheet: 0, row: 0, col: 0 }, 45000);
      wb.setCalcSetting('date1904', true);
      expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'number', value: 45000 });
      wb.setFormula({ sheet: 0, row: 1, col: 0 }, '=DATE(1904,1,2)');
      expect(wb.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({ kind: 'number', value: 1 });
    } finally {
      wb.dispose();
    }
  });

  it('recalculates a Manual workbook before saving unless told not to', async () => {
    const saveAfterEdit = async (calcOnSave: boolean): Promise<unknown> => {
      const wb = await WorkbookHandle.createDefault();
      try {
        wb.setNumber({ sheet: 0, row: 0, col: 0 }, 1);
        wb.setFormula({ sheet: 0, row: 0, col: 1 }, '=A1*10');
        wb.setCalcMode(1);
        wb.setCalcSetting('calcOnSave', calcOnSave);
        wb.setNumber({ sheet: 0, row: 0, col: 0 }, 2);
        expect(wb.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({ kind: 'number', value: 10 });
        const restored = await WorkbookHandle.loadBytes(wb.save());
        try {
          return restored.getValue({ sheet: 0, row: 0, col: 1 });
        } finally {
          restored.dispose();
        }
      } finally {
        wb.dispose();
      }
    };
    expect(await saveAfterEdit(true)).toEqual({ kind: 'number', value: 20 });
    expect(await saveAfterEdit(false)).toEqual({ kind: 'number', value: 10 });
  });

  it('reports nothing on the stub engine', async () => {
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    try {
      expect(wb.calcSettings()).toBeNull();
      expect(wb.setCalcSetting('r1c1', true)).toBe(false);
    } finally {
      wb.dispose();
    }
  });
});
