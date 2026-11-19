import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { type MountedStubSheet, mountStubSheet } from '../../../test-utils/index.js';
import { setupPalette } from './fixtures.js';

const live = vi.hoisted(() => ({ count: 0 }));

// Counts live recent-function listeners (the mounted sheet holds its own), so a teardown
// path that misses the release shows.
vi.mock('../../../../src/commands/function-history.js', async (importOriginal) => {
  const actual =
    await importOriginal<typeof import('../../../../src/commands/function-history.js')>();
  return {
    ...actual,
    subscribeRecentFunctions: (...args: Parameters<typeof actual.subscribeRecentFunctions>) => {
      const unsubscribe = actual.subscribeRecentFunctions(...args);
      live.count += 1;
      let released = false;
      return () => {
        if (!released) {
          released = true;
          live.count -= 1;
        }
        return unsubscribe();
      };
    },
  };
});

describe('interact/mac-formula-palette recent-functions subscription', () => {
  let sheet: MountedStubSheet;

  beforeEach(async () => {
    live.count = 0;
    sheet = await mountStubSheet({ workbook: await WorkbookHandle.createDefault() });
  });

  afterEach(() => sheet.dispose());

  it.each(['close', 'discard', 'detach'] as const)('releases the listener on %s', (exit) => {
    const { palette } = setupPalette(sheet);
    const base = live.count;
    palette.open();
    expect(live.count).toBe(base + 1);
    palette[exit]();
    expect(live.count).toBe(base);
    if (exit !== 'detach') palette.detach();
  });

  it('holds one listener across reopen cycles', () => {
    const { palette } = setupPalette(sheet);
    const base = live.count;
    palette.open();
    palette.close();
    palette.open();
    expect(live.count).toBe(base + 1);
    palette.detach();
    expect(live.count).toBe(base);
  });
});
