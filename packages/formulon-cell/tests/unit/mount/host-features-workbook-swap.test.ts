import { afterEach, describe, expect, it, vi } from 'vitest';

import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/mount.js';

describe('host features after setWorkbook', () => {
  let sheet: MountedStubSheet | null = null;
  afterEach(() => {
    sheet?.dispose();
    sheet = null;
  });

  it('pushes Ctrl+wheel zoom to the replacement workbook, not the old one', async () => {
    const original = await WorkbookHandle.createDefault({ preferStub: true });
    const next = await WorkbookHandle.createDefault({ preferStub: true });
    sheet = await mountStubSheet({ workbook: original });
    await sheet.instance.setWorkbook(next);
    const oldZoom = vi.spyOn(original, 'setSheetZoom');
    const nextZoom = vi.spyOn(next, 'setSheetZoom');

    const grid = sheet.host.querySelector<HTMLElement>('.fc-host__grid');
    if (!grid) throw new Error('grid element missing');
    const event = new WheelEvent('wheel', { deltaY: -100, bubbles: true, cancelable: true });
    // happy-dom drops modifier keys from the WheelEvent init bag.
    Object.defineProperty(event, 'ctrlKey', { value: true });
    grid.dispatchEvent(event);

    expect(oldZoom).not.toHaveBeenCalled();
    expect(nextZoom).toHaveBeenCalledTimes(1);
  });
});
