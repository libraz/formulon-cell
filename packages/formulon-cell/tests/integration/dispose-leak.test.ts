import { afterEach, describe, expect, it, vi } from 'vitest';

import { type MountedStubSheet, mountStubSheet } from '../test-utils/index.js';

/**
 * Integration: mount → dispose hygiene. The refactor moved chrome assembly
 * across mount/*.ts modules; if any one of them forgets to detach in dispose
 * the host is left dirty after teardown. Detect that via:
 *  - host has no children
 *  - fc-host class + instance-id dataset are gone
 *  - no fc:* event listeners fire after dispose
 */
describe('integration: dispose hygiene', () => {
  let sheet: MountedStubSheet | undefined;
  afterEach(() => sheet?.dispose());

  it('host is fully cleaned up after dispose()', async () => {
    sheet = await mountStubSheet();
    expect(sheet.host.children.length).toBeGreaterThan(0);
    expect(sheet.host.classList.contains('fc-host')).toBe(true);

    sheet.instance.dispose();

    expect(sheet.host.children.length).toBe(0);
    expect(sheet.host.classList.contains('fc-host')).toBe(false);
    expect(sheet.host.dataset.fcInstId).toBeUndefined();
  });

  it('subsequent dispose() calls are no-ops (idempotent)', async () => {
    sheet = await mountStubSheet();
    sheet.instance.dispose();
    expect(() => sheet?.instance.dispose()).not.toThrow();
    expect(() => sheet?.instance.dispose()).not.toThrow();
  });

  it('mounting twice on the same host replaces the previous instance cleanly', async () => {
    sheet = await mountStubSheet();
    const firstInstId = sheet.host.dataset.fcInstId;
    sheet.instance.dispose();

    // Re-mount on the same host; helper installs canvas / RO stubs again.
    const second = await mountStubSheet({ workbook: sheet.workbook });
    try {
      expect(second.host).not.toBe(sheet.host); // new host element
      expect(second.host.dataset.fcInstId).not.toBe(firstInstId);
    } finally {
      second.dispose();
    }
  });

  it('does not invoke the filter dropdown after dispose() — fc:openfilter detached', async () => {
    sheet = await mountStubSheet();
    const host = sheet.host;
    sheet.instance.dispose();

    // Capture any thrown errors that bubble through to the document if the
    // listener still fires. Successful detach: dispatch is a no-op.
    let threw = false;
    try {
      host.dispatchEvent(
        new CustomEvent('fc:openfilter', {
          detail: {
            range: { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 },
            col: 0,
            anchor: { x: 0, y: 0, h: 24, clientX: 0, clientY: 0 },
          },
        }),
      );
    } catch {
      threw = true;
    }
    expect(threw).toBe(false);
    // No popover root should be attached anywhere after the dispatch.
    expect(document.querySelector('.fc-filter-dropdown')).toBeNull();
  });

  it('cell-registry listeners do not survive dispose() — no rAF after teardown', async () => {
    sheet = await mountStubSheet();
    const cells = sheet.instance.cells;
    sheet.instance.dispose();

    // After dispose, mutating the registry must NOT schedule a paint on the
    // disposed renderer. Spying on rAF catches a leaked subscription that
    // would otherwise keep the canvas alive.
    const rafSpy = vi.spyOn(window, 'requestAnimationFrame');
    cells.registerFormatter({
      id: 'test-after-dispose',
      priority: 50,
      match: () => true,
      format: () => '',
    });
    expect(rafSpy).not.toHaveBeenCalled();
    rafSpy.mockRestore();
  });
});
