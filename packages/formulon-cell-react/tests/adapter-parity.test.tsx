import path from 'node:path';
import { fileURLToPath } from 'node:url';
import ts from 'typescript';
import { describe, expect, it } from 'vitest';

// React and Vue expose the same options, events and handle. Props that only
// make sense in one framework are listed here and excluded from the comparison.
const REACT_ONLY = new Set([
  'className',
  'style',
  'children',
  'errorFallback',
  'signal',
  'onReady',
]);
const VUE_ONLY = new Set(['class', 'style', 'errorFallback', 'signal']);
const VUE_ONLY_EMITS = new Set(['ready']);

const here = path.dirname(fileURLToPath(import.meta.url));
const reactSrc = path.resolve(here, '../src/index.ts');
const vueSpreadsheet = path.resolve(here, '../../formulon-cell-vue/src/Spreadsheet.ts');
const vueIndex = path.resolve(here, '../../formulon-cell-vue/src/index.ts');

// Core's types come from its source through the adapters' type-check config, as `yarn typecheck`
// reads them, so the comparison never depends on a built core.
const typecheckConfig = path.resolve(here, '../tsconfig.typecheck.json');
const { options } = ts.parseJsonConfigFileContent(
  ts.readConfigFile(typecheckConfig, ts.sys.readFile).config,
  ts.sys,
  path.dirname(typecheckConfig),
);
const coreIndex = path.resolve(here, '../../formulon-cell/src/index.ts');
const program = ts.createProgram([reactSrc, vueSpreadsheet, vueIndex], options);
const checker = program.getTypeChecker();

const exportedTypeKeys = (file: string, name: string): string[] => {
  const sf = program.getSourceFile(file);
  const mod = sf ? checker.getSymbolAtLocation(sf) : undefined;
  let sym = mod ? checker.getExportsOfModule(mod).find((s) => s.getName() === name) : undefined;
  if (!sym) throw new Error(`${name} is not exported from ${file}`);
  if (sym.flags & ts.SymbolFlags.Alias) sym = checker.getAliasedSymbol(sym);
  return checker.getPropertiesOfType(checker.getDeclaredTypeOfSymbol(sym)).map((p) => p.getName());
};

// Keys of the `props` / `emits` object literals in the Vue component definition.
const vueOptionKeys = (option: 'props' | 'emits'): string[] => {
  const sf = program.getSourceFile(vueSpreadsheet);
  const keys: string[] = [];
  const unwrap = (expr: ts.Expression): ts.Expression =>
    ts.isSatisfiesExpression(expr) || ts.isAsExpression(expr) ? unwrap(expr.expression) : expr;
  // A spread `...name` contributes the keys of the module-level `const name = { ... }`.
  const spreadSource = (name: string): ts.ObjectLiteralExpression | undefined => {
    for (const stmt of sf?.statements ?? []) {
      if (!ts.isVariableStatement(stmt)) continue;
      for (const decl of stmt.declarationList.declarations) {
        if (!ts.isIdentifier(decl.name) || decl.name.text !== name || !decl.initializer) continue;
        const init = unwrap(decl.initializer);
        if (ts.isObjectLiteralExpression(init)) return init;
      }
    }
    return undefined;
  };
  const collect = (obj: ts.ObjectLiteralExpression): void => {
    for (const p of obj.properties) {
      if (ts.isSpreadAssignment(p) && ts.isIdentifier(p.expression)) {
        const source = spreadSource(p.expression.text);
        if (source) collect(source);
      } else if (p.name && (ts.isIdentifier(p.name) || ts.isStringLiteral(p.name))) {
        keys.push(p.name.text);
      }
    }
  };
  const visit = (node: ts.Node): void => {
    if (
      ts.isPropertyAssignment(node) &&
      ts.isIdentifier(node.name) &&
      node.name.text === option &&
      ts.isObjectLiteralExpression(node.initializer)
    ) {
      collect(node.initializer);
      return;
    }
    ts.forEachChild(node, visit);
  };
  if (sf) visit(sf);
  return keys;
};

const sorted = (xs: Iterable<string>): string[] => [...xs].sort();
const eventName = (prop: string): string => prop.charAt(2).toLowerCase() + prop.slice(3);

describe('React and Vue adapter parity', () => {
  const reactProps = exportedTypeKeys(reactSrc, 'SpreadsheetProps');
  const vueProps = vueOptionKeys('props');
  const vueEmits = vueOptionKeys('emits');

  it('read core from source', () => {
    expect(program.getSourceFile(coreIndex)).toBeDefined();
  });

  it('accept the same option props', () => {
    const react = reactProps.filter((p) => !/^on[A-Z]/.test(p) && !REACT_ONLY.has(p));
    const vue = vueProps.filter((p) => !VUE_ONLY.has(p));
    expect(sorted(react)).toEqual(sorted(vue));
  });

  it('report the same events', () => {
    const react = reactProps.filter((p) => /^on[A-Z]/.test(p) && !REACT_ONLY.has(p)).map(eventName);
    const vue = vueEmits.filter((e) => !VUE_ONLY_EMITS.has(e));
    expect(sorted(react)).toEqual(sorted(vue));
  });

  it('hand out the instance under one name', () => {
    expect(exportedTypeKeys(reactSrc, 'SpreadsheetRef')).toEqual(['spreadsheet']);
    expect(exportedTypeKeys(vueIndex, 'SpreadsheetExposed')).toEqual(['spreadsheet']);
  });
});
