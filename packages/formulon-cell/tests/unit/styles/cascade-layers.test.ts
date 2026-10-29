import { readdirSync, readFileSync, statSync } from 'node:fs';
import { dirname, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';

const styleRoot = resolve(dirname(fileURLToPath(import.meta.url)), '../../../src/styles');

const walkCss = (dir: string): string[] => {
  const out: string[] = [];
  for (const name of readdirSync(dir)) {
    const full = resolve(dir, name);
    if (statSync(full).isDirectory()) out.push(...walkCss(full));
    else if (name.endsWith('.css')) out.push(full);
  }
  return out;
};

const stripCommentsAndStrings = (css: string): string =>
  css
    .replace(/\/\*[\s\S]*?\*\//g, '')
    .replace(/"(?:\\.|[^"\\])*"/g, '""')
    .replace(/'(?:\\.|[^'\\])*'/g, "''");

/** Top-level statements and block preludes, in source order. */
const topLevelStatements = (css: string): string[] => {
  const src = stripCommentsAndStrings(css);
  const out: string[] = [];
  let depth = 0;
  let start = 0;
  for (let i = 0; i < src.length; i += 1) {
    const ch = src[i];
    if (ch === '{') {
      if (depth === 0) out.push(src.slice(start, i).trim());
      depth += 1;
    } else if (ch === '}') {
      depth -= 1;
      if (depth === 0) start = i + 1;
    } else if (ch === ';' && depth === 0) {
      out.push(src.slice(start, i + 1).trim());
      start = i + 1;
    }
  }
  return out.filter((s) => s.length > 0);
};

const readTopLevel = (file: string): string[] => topLevelStatements(readFileSync(file, 'utf8'));

const declaredLayers = (readTopLevel(resolve(styleRoot, 'layers.css'))[0] ?? '')
  .replace(/^@layer\s+/, '')
  .replace(/;$/, '')
  .split(',')
  .map((name) => name.trim());

const files = walkCss(styleRoot);

describe('stylesheet cascade layers', () => {
  it('declares the layer order once', () => {
    expect(declaredLayers).toEqual([
      'fc.reset',
      'fc.base',
      'fc.theme',
      'fc.tokens',
      'fc.surface',
      'fc.app',
      'fc.chrome',
    ]);
  });

  // An unlayered rule outranks every layered one, so a stray rule both escapes
  // consumer overrides and defeats a host's `revert-layer` reset recipe.
  it.each(files.map((file) => [relative(styleRoot, file), file]))(
    '%s keeps every rule inside a declared layer',
    (_name, file) => {
      const isLayersFile = file === resolve(styleRoot, 'layers.css');
      for (const statement of readTopLevel(file)) {
        if (statement.startsWith('@import')) continue;
        if (isLayersFile && statement.endsWith(';')) continue;
        const layer = statement.match(/^@layer\s+([\w.-]+)$/)?.[1];
        expect(layer, `unlayered statement: ${statement.slice(0, 80)}`).toBeDefined();
        expect(declaredLayers).toContain(layer);
      }
    },
  );

  it.each(['index.css', 'toolbar.css'])('public entry %s fixes the layer order first', (entry) => {
    const css = readFileSync(resolve(styleRoot, entry), 'utf8').replace(/\/\*[\s\S]*?\*\//g, '');
    expect(css.trim().split('\n')[0]).toBe('@import "./layers.css";');
  });
});
