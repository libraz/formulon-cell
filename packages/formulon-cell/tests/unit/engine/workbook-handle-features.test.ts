import { describe, expect, it } from 'vitest';

import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';

/**
 * Capability-gated feature methods live on the WorkbookHandle. Under the stub
 * engine those that need first-class engine support no-op (return false / []).
 * We pin that contract here so future refactors can't silently flip a stub
 * call to a thrown error.
 */
describe('engine/workbook-handle-features (stub) — capability gates', () => {
  it('setColumnWidth no-ops on stub', async () => {
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    expect(wb.capabilities.colRowSize).toBe(false);
    expect(wb.setColumnWidth(0, 0, 0, 120)).toBe(false);
  });

  it('setRowHeight no-ops on stub', async () => {
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    expect(wb.setRowHeight(0, 0, 24)).toBe(false);
  });

  it('getColumnLayouts returns [] on stub', async () => {
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    expect(wb.getColumnLayouts(0)).toEqual([]);
  });

  it('setSheetFreeze no-ops on stub', async () => {
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    expect(wb.setSheetFreeze(0, 1, 1)).toBe(false);
  });

  it('setSheetZoom no-ops on stub', async () => {
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    expect(wb.setSheetZoom(0, 150)).toBe(false);
  });

  it('setSheetTabHidden no-ops on stub', async () => {
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    expect(wb.setSheetTabHidden(0, true)).toBe(false);
  });

  it('capabilities snapshot reports the stub flavor', async () => {
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    expect(wb.isStub).toBe(true);
    // The stub publishes a capabilities object so callers can branch.
    expect(wb.capabilities).toBeDefined();
    expect(typeof wb.capabilities).toBe('object');
  });
});
