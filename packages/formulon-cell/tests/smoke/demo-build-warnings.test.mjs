import { spawnSync } from 'node:child_process';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../../..');
const demos = ['@formulon-cell/playground', '@formulon-cell/react-demo', '@formulon-cell/vue-demo'];
const forbiddenWarnings = [/externalized for browser compatibility/];

const extractWarningLines = (output) =>
  output
    .split(/\r?\n/)
    .filter((line) => forbiddenWarnings.some((pattern) => pattern.test(line)))
    .join('\n');

describe('demo production builds', () => {
  it('complete without Node builtin externalization warnings', () => {
    for (const demo of demos) {
      const result = spawnSync('yarn', ['workspace', demo, 'build'], {
        cwd: root,
        encoding: 'utf8',
      });
      const output = `${result.stdout ?? ''}${result.stderr ?? ''}`;

      expect(result.status, `${demo} build failed.\n${output}`).toBe(0);
      expect(extractWarningLines(output), `${demo} emitted forbidden Vite warnings.`).toBe('');
    }
  }, 120_000);
});
