import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');
const css = readFileSync(join(root, 'src/styles/core/app/dialog-modules/sort.css'), 'utf8').replace(
  /\/\*[\s\S]*?\*\//g,
  '',
);

const box = (button: string): { left: number; right: number } => {
  const selector = `.fc-overlay-portal[data-fc-platform="mac"] .fc-sortdlg .fc-sortdlg__${button}`;
  for (const m of css.matchAll(/([^{}]*)\{([^}]*)\}/g)) {
    if ((m[1] ?? '').trim() !== selector) continue;
    const left = Number(/left: (\d+)px;/.exec(m[2] ?? '')?.[1]);
    const width = Number(/width: (\d+)px;/.exec(m[2] ?? '')?.[1]);
    return { left, right: left + width };
  }
  throw new Error(`no rule for ${selector}`);
};

// The Mac Sort dialog's Add, Delete and Copy Level buttons sit side by side on the measured
// 23 px pitch: each meets the next without covering it.
describe('Mac Sort dialog level buttons', () => {
  it('abut without overlapping across the measured group span', () => {
    const add = box('add-level');
    const del = box('delete-level');
    const copy = box('copy-level');
    expect(add.right).toBe(del.left);
    expect(del.right).toBe(copy.left);
    expect([add.left, copy.right]).toEqual([19, 114]);
  });
});
