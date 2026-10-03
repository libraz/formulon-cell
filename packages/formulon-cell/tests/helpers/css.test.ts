import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { readCssWithImports } from './css.js';

const withTempCss = (run: (root: string) => void): void => {
  const root = mkdtempSync(join(tmpdir(), 'formulon-css-reader-'));
  try {
    run(root);
  } finally {
    rmSync(root, { force: true, recursive: true });
  }
};

describe('readCssWithImports', () => {
  it('expands nested imports in source order', () => {
    withTempCss((root) => {
      writeFileSync(join(root, 'grandchild.css'), 'grandchild rules\n');
      writeFileSync(join(root, 'child.css'), '@import "./grandchild.css";\nchild rules\n');
      writeFileSync(join(root, 'root.css'), '@import "./child.css";\nroot rules\n');

      expect(readCssWithImports(join(root, 'root.css'))).toBe(
        'grandchild rules\nchild rules\nroot rules\n',
      );
    });
  });

  it('reports missing imports', () => {
    withTempCss((root) => {
      writeFileSync(join(root, 'root.css'), '@import "./absent.css";\n');

      expect(() => readCssWithImports(join(root, 'root.css'))).toThrow(
        'CSS import not found: ./absent.css',
      );
    });
  });

  it('reports circular imports', () => {
    withTempCss((root) => {
      writeFileSync(join(root, 'first.css'), '@import "./second.css";\n');
      writeFileSync(join(root, 'second.css'), '@import "./first.css";\n');

      expect(() => readCssWithImports(join(root, 'first.css'))).toThrow('Circular CSS import:');
    });
  });
});
