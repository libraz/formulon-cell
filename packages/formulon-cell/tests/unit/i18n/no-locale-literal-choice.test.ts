import { readdirSync, readFileSync, statSync } from 'node:fs';
import { join, relative } from 'node:path';
import { describe, expect, it } from 'vitest';

// UI text comes from the locale dictionaries, never from a branch on the locale code: a
// two-way branch silently leaves every other locale on one language.
const SRC = join(__dirname, '../../../src');
const DICTIONARY_DIR = 'i18n/strings/';

/** Non-text locale branches, each with the reason it is not UI wording. */
const ALLOWED: Readonly<Record<string, string>> = {
  'components/color-palette.ts:list-separator':
    'list punctuation joining colour names, not a translatable label',
  'interact/format-dialog-model.ts:format-code-lists':
    'number-format codes listed per locale, data rather than wording',
};

const sourceFiles = (dir: string): string[] =>
  readdirSync(dir).flatMap((name) => {
    const path = join(dir, name);
    if (statSync(path).isDirectory()) return sourceFiles(path);
    return path.endsWith('.ts') ? [path] : [];
  });

const LOCALE_CODE = "'[a-z]{2}(?:-[A-Z]{2})?'";
const LOCALE_NAME = String.raw`\b\w*(?:lang|locale|Lang|Locale)\w*(?:\(\))?`;
const RULES: readonly { name: string; pattern: RegExp }[] = [
  {
    name: 'locale comparison choosing a string literal',
    pattern: new RegExp(String.raw`${LOCALE_NAME}\s*[!=]==\s*${LOCALE_CODE}\s*\?\s*['"\`]`, 'g'),
  },
  {
    name: 'locale comparison choosing a per-language field',
    pattern: new RegExp(
      String.raw`${LOCALE_NAME}\s*[!=]==\s*${LOCALE_CODE}\s*\?\s*[\w.\[\]]*\.(?:ja|en)\b`,
      'g',
    ),
  },
  {
    // A locale test in an `if` condition, with the wording set inside the guarded statement.
    name: 'locale-guarded statement setting a label literal',
    pattern: new RegExp(
      String.raw`\bif\s*\([^;{]*?${LOCALE_NAME}\s*[!=]==\s*${LOCALE_CODE}[^;{]*?\)\s*(?:\{[^}]*?|[^;{]*?)\b(?:label|title|text|caption|heading)\s*:\s*['"\`]`,
      'g',
    ),
  },
  {
    // A helper taking one Japanese and one English string fills every other locale from one of them.
    name: 'ja/en pair helper',
    pattern: /\(\s*(?:ja\s*:\s*string\s*,\s*en|en\s*:\s*string\s*,\s*ja)\s*:\s*string\b/g,
  },
  { name: 'per-language label field', pattern: /\b\w+Ja\b|\.ja\b|\b\w*_JA_\w*|\bJA_\w+/g },
];

const lineOf = (text: string, index: number): number => text.slice(0, index).split('\n').length;

const findings = (): string[] => {
  const found: string[] = [];
  for (const file of sourceFiles(SRC)) {
    const rel = relative(SRC, file).split('\\').join('/');
    if (rel.startsWith(DICTIONARY_DIR)) continue;
    const text = readFileSync(file, 'utf8');
    for (const rule of RULES) {
      for (const match of text.matchAll(rule.pattern)) {
        found.push(`${rel}:${lineOf(text, match.index ?? 0)} ${rule.name}: ${match[0]}`);
      }
    }
  }
  return found;
};

describe('locale-specific UI text', () => {
  it('comes from the dictionaries, not from a branch on the locale code', () => {
    const allowedFiles = Object.keys(ALLOWED).map((key) => key.split(':')[0]);
    const unexpected = findings().filter(
      (finding) => !allowedFiles.some((file) => finding.startsWith(`${file}:`)),
    );
    expect(unexpected).toEqual([]);
  });

  it('keeps each allowlisted branch present, so a stale entry is removed', () => {
    const all = findings();
    for (const key of Object.keys(ALLOWED)) {
      const file = key.split(':')[0];
      expect(all.some((finding) => finding.startsWith(`${file}:`))).toBe(true);
    }
  });
});
