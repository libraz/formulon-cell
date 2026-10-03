import { existsSync, readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';

const CSS_IMPORT_PATTERN = /^@import\s+["']([^"']+)["'];[ \t]*\r?\n?/gm;

const expandCssFile = (filePath: string, importedFiles: Set<string>): string => {
  const absolutePath = resolve(filePath);
  if (importedFiles.has(absolutePath)) {
    throw new Error(`Circular CSS import: ${absolutePath}`);
  }

  const nextImportedFiles = new Set(importedFiles);
  nextImportedFiles.add(absolutePath);
  const source = readFileSync(absolutePath, 'utf8');

  return source.replace(CSS_IMPORT_PATTERN, (match, importPath: string) => {
    if (/^(?:[a-z][a-z\d+.-]*:|\/\/)/i.test(importPath)) return match;

    const importedFile = resolve(dirname(absolutePath), importPath);
    if (!existsSync(importedFile)) {
      throw new Error(`CSS import not found: ${importPath} (from ${absolutePath})`);
    }
    return expandCssFile(importedFile, nextImportedFiles);
  });
};

export const readCssWithImports = (filePath: string): string => expandCssFile(filePath, new Set());
