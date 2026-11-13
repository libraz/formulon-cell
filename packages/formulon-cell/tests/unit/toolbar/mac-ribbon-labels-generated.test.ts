import { execFileSync } from 'node:child_process';
import { existsSync } from 'node:fs';
import { dirname, resolve as resolvePath } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  GROUP_LABELS,
  labels,
  macRibbonLabelForCommand,
} from '../../../src/toolbar/ribbon/mac/labels.js';
import { MAC_BUNDLE_LABELS } from '../../../src/toolbar/ribbon/mac/labels-bundle.js';
import { MAC_LOCALES } from '../../../src/toolbar/ribbon/mac/labels-data.js';

const MAC_APP = '/Applications/Microsoft Excel.app';
const HERE = dirname(fileURLToPath(import.meta.url));
const GENERATOR = resolvePath(HERE, '../../../../../scripts/gen-mac-ribbon-labels.py');
const BUNDLE_FILE = resolvePath(HERE, '../../../src/toolbar/ribbon/mac/labels-bundle.ts');

const resolve = (namespace: string, name: string, locale: (typeof MAC_LOCALES)[number]) => {
  if (namespace === 'labels') return labels[name]?.[locale];
  if (namespace === 'groups') return GROUP_LABELS[name]?.[locale];
  if (namespace === 'menu') return macRibbonLabelForCommand(name, locale);
  throw new Error(`Unknown bundle label namespace: ${namespace}`);
};

describe('generated Mac ribbon labels', () => {
  const entries = Object.entries(MAC_BUNDLE_LABELS);

  it('is not empty', () => {
    expect(entries.length).toBeGreaterThan(0);
  });

  it('never names the desktop product', () => {
    const named = entries.flatMap(([key, bundle]) =>
      Object.entries(bundle)
        .filter(([, text]) => /Excel/.test(text))
        .map(([locale]) => `${key} (${locale})`),
    );
    expect(named).toEqual([]);
  });

  it.each(entries)('%s resolves to the bundle label in every locale', (key, bundle) => {
    const split = key.indexOf(':');
    const namespace = key.slice(0, split);
    const name = key.slice(split + 1);
    for (const locale of MAC_LOCALES) {
      expect(resolve(namespace, name, locale), `${key} (${locale})`).toBe(bundle[locale]);
    }
  });

  it('cleans accelerator markers the way the generator documents', () => {
    expect(() =>
      execFileSync('python3', ['-I', '-m', 'doctest', GENERATOR], { stdio: 'pipe' }),
    ).not.toThrow();
  });

  it.skipIf(!existsSync(MAC_APP))('matches what the generator produces from the app bundle', () => {
    expect(() =>
      execFileSync('python3', ['-I', GENERATOR, '--check', BUNDLE_FILE], { stdio: 'pipe' }),
    ).not.toThrow();
  });
});
