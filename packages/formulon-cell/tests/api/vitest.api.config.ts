import { rmSync } from 'node:fs';
import { createRequire } from 'node:module';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { defineConfig } from 'vitest/config';

import { entryAliases } from '../entry-aliases.js';
import { scaledTimeout } from '../timeouts.js';

const packageDir = fileURLToPath(new URL('../..', import.meta.url));

// The typecheck pool's incremental state has reported errors the sources no longer have; start clean.
const vitestDir = dirname(createRequire(import.meta.url).resolve('vitest/package.json'));
rmSync(join(vitestDir, 'dist/tsconfig.tmp.tsbuildinfo'), { force: true });

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
    testTimeout: scaledTimeout(15_000),
    hookTimeout: scaledTimeout(10_000),
    typecheck: {
      enabled: true,
      include: ['tests/api/**/*.test-d.ts'],
      tsconfig: './tests/api/tsconfig.json',
    },
  },
});
