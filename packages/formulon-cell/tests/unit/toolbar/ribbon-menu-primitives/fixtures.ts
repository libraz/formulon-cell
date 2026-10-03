import { readdirSync, readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import ts from 'typescript';

export const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../../..');
export const menusDir = join(root, 'src/toolbar/ribbon/menus');
export const ribbonDir = join(root, 'src/toolbar/ribbon');
export const mountDir = join(root, 'src/mount');
export const disabledStateAuditDirs = [
  'src/interact',
  'src/mount',
  'src/toolbar',
  'src/components',
];

export const menuSources = (): { name: string; source: string }[] =>
  readdirSync(menusDir)
    .filter((name) => name.endsWith('.ts'))
    .map((name) => ({
      name,
      source: readFileSync(join(menusDir, name), 'utf8'),
    }));

export const menuConsumerSources = (): { name: string; source: string }[] => [
  ...menuSources(),
  {
    name: 'backstage-title.ts',
    source: readFileSync(join(ribbonDir, 'backstage-title.ts'), 'utf8'),
  },
];

export const sourceFilesUnder = (path: string): string[] => {
  const absolutePath = join(root, path);
  const files: string[] = [];
  for (const entry of readdirSync(absolutePath, { withFileTypes: true })) {
    const entryPath = `${path}/${entry.name}`;
    if (entry.isDirectory()) {
      files.push(...sourceFilesUnder(entryPath));
    } else if (entry.name.endsWith('.ts')) {
      files.push(entryPath);
    }
  }
  return files.sort();
};

export const collectStringLiteralArgs = (callName: string, argIndex: number): string[] => {
  const values = new Set<string>();
  for (const { name, source } of menuConsumerSources()) {
    if (name === 'general.ts') continue;
    const file = ts.createSourceFile(name, source, ts.ScriptTarget.Latest, true, ts.ScriptKind.TS);
    const visit = (node: ts.Node): void => {
      if (
        ts.isCallExpression(node) &&
        ts.isIdentifier(node.expression) &&
        node.expression.text === callName
      ) {
        const arg = node.arguments[argIndex];
        if (arg && ts.isStringLiteralLike(arg)) values.add(arg.text);
      }
      ts.forEachChild(node, visit);
    };
    visit(file);
  }
  return [...values].sort();
};

export const collectVisualMenuTileIcons = (): string[] => {
  const values = new Set<string>();
  const collectIconFromObject = (opts: ts.ObjectLiteralExpression): void => {
    const iconProp = opts.properties.find(
      (prop): prop is ts.PropertyAssignment =>
        ts.isPropertyAssignment(prop) &&
        ts.isIdentifier(prop.name) &&
        prop.name.text === 'icon' &&
        ts.isStringLiteralLike(prop.initializer),
    );
    if (iconProp && ts.isStringLiteralLike(iconProp.initializer)) {
      values.add(iconProp.initializer.text);
    }
  };

  for (const { name, source } of menuConsumerSources()) {
    if (name === 'general.ts') continue;
    const file = ts.createSourceFile(name, source, ts.ScriptTarget.Latest, true, ts.ScriptKind.TS);
    const visit = (node: ts.Node): void => {
      if (!ts.isCallExpression(node) || !ts.isIdentifier(node.expression)) {
        ts.forEachChild(node, visit);
        return;
      }

      if (node.expression.text === 'visualMenuTile') {
        const opts = node.arguments[0];
        if (opts && ts.isObjectLiteralExpression(opts)) {
          collectIconFromObject(opts);
        }
      }
      if (node.expression.text === 'visualMenuTileGrid') {
        const tiles = node.arguments[1];
        if (tiles && ts.isArrayLiteralExpression(tiles)) {
          for (const element of tiles.elements) {
            if (ts.isObjectLiteralExpression(element)) collectIconFromObject(element);
          }
        }
      }
      ts.forEachChild(node, visit);
    };
    visit(file);
  }
  return [...values].sort();
};

export const sourcesOutsidePrimitives = (): { name: string; source: string }[] =>
  menuSources().filter(({ name }) => name !== 'general.ts');
