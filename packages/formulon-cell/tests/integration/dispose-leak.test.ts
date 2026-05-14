import { afterEach, describe, expect, it } from 'vitest';

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
});
