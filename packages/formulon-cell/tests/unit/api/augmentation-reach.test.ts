import { readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';

const API = join(__dirname, '../../../src/api');

describe('published Sheet and Workbook types', () => {
  it('reach every module that augments them', () => {
    const augmenting = readdirSync(API)
      .filter((file) => file.endsWith('.ts') && file !== 'workbook.ts')
      .filter((file) =>
        readFileSync(join(API, file), 'utf8').includes("declare module './workbook.js'"),
      )
      .map((file) => file.replace(/\.ts$/, '.js'));
    const facade = readFileSync(join(API, 'workbook.ts'), 'utf8');
    const missing = augmenting.filter(
      (file) => !facade.includes(`export type * from './${file}';`),
    );
    expect(augmenting.length).toBeGreaterThan(0);
    expect(missing).toEqual([]);
  });
});
