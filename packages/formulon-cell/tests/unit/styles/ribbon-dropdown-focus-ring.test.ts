import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');

/** Rules whose selector puts a ribbon dropdown button in :focus-visible. */
const focusRules = (file: string): { selector: string; body: string }[] =>
  [...readFileSync(join(root, file), 'utf8').matchAll(/([^{}]+)\{([^{}]*)\}/g)]
    .map((m) => ({ selector: (m[1] ?? '').trim(), body: m[2] ?? '' }))
    .filter((rule) => /rb-dd__btn:focus-visible/.test(rule.selector));

describe('ribbon dropdown focus ring', () => {
  it('draws the shared 2 px accent outline on a focused dropdown button', () => {
    const shared = focusRules('src/styles/toolbar/ribbon/dropdowns.css').find(
      (rule) => rule.selector === '.fc-tb__rb-dd__btn:focus-visible',
    );
    expect(shared?.body).toContain('outline: 2px solid var(--fc-tb-accent)');
  });

  it.each(['src/styles/toolbar/ribbon/dropdowns.css', 'src/styles/platform-mac.css'])(
    'is not removed for any dropdown variant in %s',
    (file) => {
      const removed = focusRules(file).filter((rule) => /outline:\s*(none|0)\b/.test(rule.body));
      expect(removed.map((rule) => rule.selector)).toEqual([]);
    },
  );
});
