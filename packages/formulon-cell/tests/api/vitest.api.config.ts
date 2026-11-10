import { fileURLToPath } from 'node:url';
import { defineConfig } from 'vitest/config';

import { entryAliases } from '../entry-aliases.js';

const packageDir = fileURLToPath(new URL('../..', import.meta.url));

// Consumer tests import only through the package's entry points.
export default defineConfig({
  root: packageDir,
  resolve: {
    alias: entryAliases,
  },
  test: {
    environment: 'happy-dom',
    setupFiles: ['./tests/setup.ts', './tests/api/setup-dom.ts'],
    include: ['tests/api/**/*.test.ts'],
    globals: false,
    testTimeout: 15_000,
    typecheck: {
      enabled: true,
      include: ['tests/api/**/*.test-d.ts'],
      tsconfig: './tests/api/tsconfig.json',
    },
  },
});
