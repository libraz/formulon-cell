import { defineConfig } from 'vitest/config';

export default defineConfig({
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
