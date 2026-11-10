import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { defineConfig } from 'vitest/config';

const packageDir = fileURLToPath(new URL('../..', import.meta.url));
const src = (path: string): string => resolve(packageDir, 'src', path);

// Consumer tests import only through the package's entry points. Each alias is anchored, so a
// subpath never resolves through the root entry.
export default defineConfig({
  root: packageDir,
  resolve: {
    alias: [
      { find: /^@libraz\/formulon-cell\/toolbar$/, replacement: src('toolbar/index.ts') },
      { find: /^@libraz\/formulon-cell\/adapter$/, replacement: src('adapter/index.ts') },
      { find: /^@libraz\/formulon-cell\/extensions$/, replacement: src('extensions/index.ts') },
      { find: /^@libraz\/formulon-cell\/charts$/, replacement: src('charts/index.ts') },
      { find: /^@libraz\/formulon-cell\/i18n\/(.+)$/, replacement: src('i18n/$1.ts') },
      { find: /^@libraz\/formulon-cell$/, replacement: src('index.ts') },
    ],
  },
  test: {
    environment: 'happy-dom',
    setupFiles: ['./tests/setup.ts'],
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
