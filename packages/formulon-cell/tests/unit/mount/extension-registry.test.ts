import { describe, expect, it, vi } from 'vitest';

import {
  createExtensionRegistry,
  type ExtensionRegistryDeps,
} from '../../../src/mount/extension-registry.js';

const createRegistry = () =>
  createExtensionRegistry({ context: {} as ExtensionRegistryDeps['context'] });

describe('mount/extension-registry', () => {
  it('rolls back partial setExtensions mounts and keeps the setup error', () => {
    const registry = createRegistry();
    const oldDispose = vi.fn();
    const firstDispose = vi.fn(() => {
      throw new Error('rollback cleanup failed');
    });
    const setupError = new Error('second extension setup failed');
    const consoleError = vi.spyOn(console, 'error').mockImplementation(() => {});

    try {
      registry.mountAll([
        {
          id: 'old',
          setup: () => ({ dispose: oldDispose }),
        },
      ]);
      registry.refreshFeaturesView();
      expect(registry.featuresView.old).toBeDefined();

      expect(() =>
        registry.setExtensions([
          {
            id: 'first',
            setup: () => ({ dispose: firstDispose }),
          },
          {
            id: 'second',
            setup: () => {
              throw setupError;
            },
          },
        ]),
      ).toThrow(setupError);

      expect(oldDispose).toHaveBeenCalledTimes(1);
      expect(firstDispose).toHaveBeenCalledTimes(1);
      expect(consoleError).toHaveBeenCalledTimes(1);
      expect(registry.getUserHandle('first')).toBeUndefined();
      expect(registry.featuresView.old).toBeUndefined();
      expect(registry.featuresView.first).toBeUndefined();
      expect(registry.featuresView.second).toBeUndefined();

      const afterDispose = vi.fn();
      registry.setExtensions([
        {
          id: 'after',
          setup: () => ({ dispose: afterDispose }),
        },
      ]);
      expect(registry.featuresView.after).toBeDefined();

      registry.dispose();
      expect(afterDispose).toHaveBeenCalledTimes(1);
      expect(firstDispose).toHaveBeenCalledTimes(1);
    } finally {
      consoleError.mockRestore();
    }
  });
});
