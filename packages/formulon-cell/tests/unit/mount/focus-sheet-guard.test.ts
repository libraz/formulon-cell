import { readdirSync, readFileSync } from 'node:fs';
import { join, relative } from 'node:path';
import { describe, expect, it } from 'vitest';

const SRC = join(process.cwd(), 'src');

/** The only module that may call `.focus` on the mount element directly. */
const FOCUS_HELPER = 'interact/dialog-focus.ts';

/** Modules whose `host` is not the mount element, so a direct `.focus` is legitimate. */
const NON_MOUNT_HOST: Record<string, string> = {
  'interact/pointer.ts': 'attachPointer receives the .fc-host__grid element as `host`',
};

const DIRECT_FOCUS = [/\bhost\)?\??\.focus\(/, /\.host\??\.focus\(/, /\bsheetEl\.focus\(/];

const sourceFiles = (dir: string): string[] =>
  readdirSync(dir, { withFileTypes: true }).flatMap((entry) => {
    const path = join(dir, entry.name);
    if (entry.isDirectory()) return sourceFiles(path);
    return entry.name.endsWith('.ts') ? [path] : [];
  });

describe('sheet focus', () => {
  it('goes through focusSheet instead of focusing the mount element directly', () => {
    const offenders: string[] = [];
    for (const file of sourceFiles(SRC)) {
      const rel = relative(SRC, file).split('\\').join('/');
      if (rel === FOCUS_HELPER || rel in NON_MOUNT_HOST) continue;
      readFileSync(file, 'utf8')
        .split('\n')
        .forEach((line, index) => {
          if (DIRECT_FOCUS.some((pattern) => pattern.test(line))) {
            offenders.push(`src/${rel}:${index + 1}: ${line.trim()}`);
          }
        });
    }
    expect(offenders).toEqual([]);
  });
});
