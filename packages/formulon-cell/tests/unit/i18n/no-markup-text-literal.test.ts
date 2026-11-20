import { readdirSync, readFileSync, statSync } from 'node:fs';
import { join, relative } from 'node:path';
import { describe, expect, it } from 'vitest';

// Visible text in generated markup comes from the locale dictionaries; a capitalized English
// phrase written between a tag and its closing tag stays English in every locale.
const SRC = join(__dirname, '../../../src');
const DICTIONARY_DIR = 'i18n/';

const sourceFiles = (dir: string): string[] =>
  readdirSync(dir).flatMap((name) => {
    const path = join(dir, name);
    if (statSync(path).isDirectory()) return sourceFiles(path);
    return path.endsWith('.ts') ? [path] : [];
  });

const MARKUP_TEXT = />\s*[A-Z][a-z]+(?:[ -][A-Za-z]+)+\s*<\//g;

const lineOf = (text: string, index: number): number => text.slice(0, index).split('\n').length;

const findings = (): string[] => {
  const found: string[] = [];
  for (const file of sourceFiles(SRC)) {
    const rel = relative(SRC, file).split('\\').join('/');
    if (rel.startsWith(DICTIONARY_DIR)) continue;
    const text = readFileSync(file, 'utf8');
    for (const match of text.matchAll(MARKUP_TEXT)) {
      found.push(`${rel}:${lineOf(text, match.index ?? 0)} ${match[0]}`);
    }
  }
  return found;
};

describe('markup text', () => {
  it('is never a literal English phrase between a tag and its closing tag', () => {
    expect(findings()).toEqual([]);
  });
});
