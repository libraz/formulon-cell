import { configDefaults, defineConfig } from 'vitest/config';

import { entryAliases } from './tests/entry-aliases.js';

export default defineConfig({
  resolve: { alias: entryAliases },
  test: {
    environment: 'happy-dom',
    setupFiles: ['./tests/setup.ts'],
    include: [
      'tests/**/*.test.ts',
      'tests/**/*.test.tsx',
      'tests/**/*.test.mjs',
      '../formulon-cell-react/tests/**/*.test.tsx',
      '../formulon-cell-vue/tests/**/*.test.ts',
      '../../tests/e2e-shared/unit/**/*.test.ts',
    ],
    // Consumer tests run through the entry points under their own config (tests/api/vitest.api.config.ts).
    exclude: [...configDefaults.exclude, 'tests/api/**'],
    globals: false,
    // Full-mount tests take ~1s locally but exceed 5s on contended CI runners
    // under coverage; the timeout only has to catch hangs.
    testTimeout: 15_000,
    coverage: {
      provider: 'v8',
      include: ['src/**/*.ts'],
      exclude: ['src/**/*.d.ts', 'src/**/index.ts', 'src/vite-env.d.ts'],
      reporter: process.env.CI ? ['text', 'lcov'] : ['text', 'html', 'lcov'],
    },
  },
});
