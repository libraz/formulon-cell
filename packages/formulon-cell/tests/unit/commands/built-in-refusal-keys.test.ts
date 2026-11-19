import { readdirSync, readFileSync } from 'node:fs';
import { join, relative } from 'node:path';
import { describe, expect, it } from 'vitest';

const SRC = join(__dirname, '../../../src');

const sourceFiles = (dir: string): string[] =>
  readdirSync(dir, { withFileTypes: true }).flatMap((entry) => {
    const path = join(dir, entry.name);
    if (entry.isDirectory()) return sourceFiles(path);
    return entry.name.endsWith('.ts') ? [path] : [];
  });

/** The text from `open` to its matching bracket, or null when unbalanced. */
const balanced = (text: string, open: number): string | null => {
  const closer = text[open] === '{' ? '}' : ')';
  let depth = 0;
  for (let i = open; i < text.length; i++) {
    if (text[i] === text[open]) depth++;
    else if (text[i] === closer && --depth === 0) return text.slice(open, i + 1);
  }
  return null;
};

/** The innermost `{ ... }` literal around `index`. */
const enclosingLiteral = (text: string, index: number): string | null => {
  let depth = 0;
  for (let i = index; i >= 0; i--) {
    if (text[i] === '}') depth++;
    else if (text[i] === '{' && depth-- === 0) return balanced(text, i);
  }
  return null;
};

const norm = (literal: string): string => literal.replace(/\s+/g, ' ');

/**
 * Refusals that carry no `reasonKey`, by file and a fragment of the literal. Each is either text a
 * host or the user wrote (shown unchanged) or a condition with no built-in reason that says it.
 */
const ALLOWED: ReadonlyArray<readonly [file: string, fragment: string]> = [
  // Data validation text the workbook author wrote.
  ['commands/interaction-controller.ts', 'reason: outcome.message'],
];

describe('built-in refusals', () => {
  it('carry a reasonKey so a consumer never shows English-only text', () => {
    const unkeyed: string[] = [];
    for (const file of sourceFiles(SRC)) {
      const text = readFileSync(file, 'utf8');
      const name = relative(SRC, file);
      const literals = new Set<string>();
      for (const match of text.matchAll(/refusedBatch\(\s*(?=\{)/g)) {
        const call = balanced(text, match.index + match[0].length);
        if (call) literals.add(call);
      }
      for (const match of text.matchAll(
        /code: '(?:readOnly|operationDenied|cellIneligible|protected|outOfBounds|effectTooLarge|invalid)'/g,
      )) {
        const literal = enclosingLiteral(text, match.index);
        if (literal && /\breason:/.test(literal)) literals.add(literal);
      }
      for (const literal of literals) {
        if (/\breasonKey\b/.test(literal)) continue;
        const flat = norm(literal);
        const allowed = ALLOWED.some(([f, fragment]) => f === name && flat.includes(fragment));
        if (!allowed) unkeyed.push(`${name}: ${flat}`);
      }
    }
    expect(unkeyed).toEqual([]);
  });

  it('lists only allowances that still match a site', () => {
    const stale = ALLOWED.filter(([file, fragment]) => {
      const text = readFileSync(join(SRC, file), 'utf8').replace(/\s+/g, ' ');
      return !text.includes(fragment);
    });
    expect(stale).toEqual([]);
  });
});
