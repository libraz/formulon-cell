import { existsSync, readdirSync, readFileSync } from 'node:fs';
import { join, relative } from 'node:path';
import { describe, expect, it } from 'vitest';

const ROOT = join(__dirname, '../../..');
const APPS = join(ROOT, 'apps');

const specFiles = (dir: string): string[] =>
  readdirSync(dir, { withFileTypes: true }).flatMap((entry) => {
    const path = join(dir, entry.name);
    if (entry.isDirectory()) return entry.name === 'node_modules' ? [] : specFiles(path);
    return entry.name.endsWith('.spec.ts') ? [path] : [];
  });

// A value import of `test` (not `import type`) from Playwright itself, under any alias.
const PLAYWRIGHT_TEST = /import\s*\{([^}]*)\}\s*from\s*['"]@playwright\/test['"]/g;
const importsBaseTest = (source: string): boolean =>
  [...source.matchAll(PLAYWRIGHT_TEST)].some(
    ([, names = '']) =>
      !/^\s*type\b/.test(names) &&
      names.split(',').some((name) => /^\s*test(\s+as\s+\w+)?\s*$/.test(name)),
  );

describe('e2e spec imports', () => {
  it('every spec takes `test` from tests/e2e-shared/test.ts, which fails on dialog layout faults', () => {
    const specs = readdirSync(APPS, { withFileTypes: true })
      .filter((entry) => entry.isDirectory())
      .map((entry) => join(APPS, entry.name, 'e2e'))
      .filter((dir) => existsSync(dir))
      .flatMap(specFiles);
    expect(specs.length).toBeGreaterThan(0);
    const offenders = specs
      .filter((path) => importsBaseTest(readFileSync(path, 'utf8')))
      .map((path) => relative(ROOT, path));
    expect(offenders).toEqual([]);
  });

  it("recognises a direct import of Playwright's own `test`", () => {
    expect(importsBaseTest("import { test } from '@playwright/test';")).toBe(true);
    expect(importsBaseTest("import { expect, test as it } from '@playwright/test';")).toBe(true);
    expect(importsBaseTest("import { expect, type Page } from '@playwright/test';")).toBe(false);
    expect(importsBaseTest("import type { test } from '@playwright/test';")).toBe(false);
    expect(importsBaseTest("import { test } from '../../../tests/e2e-shared/test.js';")).toBe(
      false,
    );
  });
});
