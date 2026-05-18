import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { defineConfig } from 'vitest/config';

const rootDir = fileURLToPath(new URL('.', import.meta.url));

export default defineConfig({
  resolve: {
    alias: {
      '@libraz/formulon-cell': resolve(rootDir, 'src/index.ts'),
    },
  },
  test: {
    environment: 'happy-dom',
    include: [
      'tests/**/*.test.ts',
      'tests/**/*.test.tsx',
      'tests/**/*.test.mjs',
      '../formulon-cell-react/tests/**/*.test.tsx',
      '../formulon-cell-vue/tests/**/*.test.ts',
    ],
    globals: false,
    coverage: {
      provider: 'v8',
      include: ['src/**/*.ts'],
      exclude: ['src/**/*.d.ts', 'src/**/index.ts', 'src/vite-env.d.ts'],
      reporter: ['text', 'html', 'lcov'],
    },
  },
});
