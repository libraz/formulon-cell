#!/usr/bin/env node
/**
 * Public-surface gate. Reads the built declaration files reachable from the
 * `exports` maps of the three packages and compares them with manifest.json:
 * per entry point, names that must be exported (as a value or as a type) and
 * names that must not be, the allowed `./i18n/*` files, and for the root entry
 * the ceiling on runtime values plus the rule that every exported type is
 * reachable from some exported value's signature (no orphan types).
 *
 * After the checks it typechecks scripts/api-surface/augmentation against the built root d.ts, so a
 * Sheet/Workbook member that the declaration build dropped fails the gate.
 *
 * Usage: node scripts/api-surface/check.mjs [--absence]
 *   --absence  also reports public return types and required readonly fields
 *              that include `undefined` (the surface uses `null` for absence).
 * `yarn check:api` builds the libraries and runs it. Exit code 1 when any violation is found.
 */
import { spawnSync } from 'node:child_process';
import fs from 'node:fs';
import { createRequire } from 'node:module';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../..');
const require = createRequire(path.join(root, 'package.json'));
const ts = require('typescript');
const manifest = JSON.parse(
  fs.readFileSync(path.join(root, 'scripts/api-surface/manifest.json'), 'utf8'),
);
const checkAbsence = process.argv.includes('--absence');

const packages = [
  'packages/formulon-cell',
  'packages/formulon-cell-react',
  'packages/formulon-cell-vue',
];
const entries = new Map(); // entry name -> d.ts path
const i18nFiles = [];
for (const dir of packages) {
  const pkg = JSON.parse(fs.readFileSync(path.join(root, dir, 'package.json'), 'utf8'));
  for (const [sub, target] of Object.entries(pkg.exports ?? {})) {
    const types = typeof target === 'object' ? target.types : null;
    if (!types) continue;
    const name = sub === '.' ? pkg.name : `${pkg.name}/${sub.slice(2)}`;
    if (sub.includes('*')) {
      const [pre, post] = types.split('*');
      // `./dist/i18n/locales/*.d.ts` splits into a directory prefix and a file suffix.
      const base = path.join(root, dir, pre.endsWith('/') ? pre : path.dirname(pre));
      if (fs.existsSync(base))
        for (const f of fs.readdirSync(base))
          if (f.endsWith(post.replace(/^\//, ''))) i18nFiles.push(f.slice(0, -post.length));
      continue;
    }
    entries.set(name, path.join(root, dir, types));
  }
}

const violations = [];
const missingEntries = [...Object.keys(manifest.entries)].filter((e) => !entries.has(e));
for (const e of missingEntries) violations.push(`${e}: entry point not in any exports map`);

const files = [...entries.values()].filter((f) => fs.existsSync(f));
for (const [e, f] of entries)
  if (!fs.existsSync(f)) violations.push(`${e}: ${path.relative(root, f)} not built`);
const program = ts.createProgram(files, {
  target: ts.ScriptTarget.ES2022,
  module: ts.ModuleKind.ESNext,
  moduleResolution: ts.ModuleResolutionKind.Bundler,
  strict: true,
  skipLibCheck: true,
  noEmit: true,
  jsx: ts.JsxEmit.ReactJSX,
});
const checker = program.getTypeChecker();

const exportsOf = (file) => {
  const sf = program.getSourceFile(file);
  const sym = sf && checker.getSymbolAtLocation(sf);
  return sym ? checker.getExportsOfModule(sym) : [];
};

const ownFile = (decl) =>
  packages.some((dir) =>
    decl.getSourceFile().fileName.startsWith(path.join(root, dir) + path.sep),
  ) && !decl.getSourceFile().fileName.includes('node_modules');
const hasUndefined = (t) =>
  (t.flags & ts.TypeFlags.Undefined) !== 0 ||
  (t.isUnion() && t.types.some((u) => (u.flags & ts.TypeFlags.Undefined) !== 0));
const absence = new Set();
const scanAbsence = (sym, label) => {
  const target = sym.flags & ts.SymbolFlags.Alias ? checker.getAliasedSymbol(sym) : sym;
  const decl = target.declarations?.[0];
  if (!decl || !ownFile(decl)) return;
  const type = checker.getTypeOfSymbolAtLocation(target, decl);
  // A component's ReactNode return is the framework's type, not an absence signal.
  const isNode = (t) => /ReactNode|JSX\.Element|VNode/.test(checker.typeToString(t));
  for (const sig of type.getCallSignatures())
    if (hasUndefined(sig.getReturnType()) && !isNode(sig.getReturnType()))
      absence.add(`${label}()`);
  const host =
    target.flags & (ts.SymbolFlags.Interface | ts.SymbolFlags.Class | ts.SymbolFlags.TypeAlias)
      ? checker.getDeclaredTypeOfSymbol(target)
      : null;
  if (!host) return;
  for (const p of checker.getPropertiesOfType(host)) {
    const pd = p.declarations?.[0];
    if (!pd || !ownFile(pd) || p.flags & ts.SymbolFlags.Optional) continue;
    const pt = checker.getTypeOfSymbolAtLocation(p, pd);
    const sigs = pt.getCallSignatures();
    if (sigs.length) {
      if (sigs.some((s) => hasUndefined(s.getReturnType())))
        absence.add(`${label}.${p.getName()}()`);
    } else if (
      pd.modifiers?.some((m) => m.kind === ts.SyntaxKind.ReadonlyKeyword) &&
      hasUndefined(pt)
    )
      absence.add(`${label}.${p.getName()}`);
  }
};

const resolve = (sym) => (sym.flags & ts.SymbolFlags.Alias ? checker.getAliasedSymbol(sym) : sym);
// A type-only re-export of a class or const is a type on this entry point.
const isValueExport = (sym) =>
  (resolve(sym).flags & ts.SymbolFlags.Value) !== 0 &&
  !(sym.declarations ?? []).some((d) => d.isTypeOnly || d.parent?.parent?.isTypeOnly);

/** Symbols of every type reachable from the given values' types, members included. */
const reachableTypes = (values) => {
  const seenTypes = new Set();
  const seenAlias = new Set();
  const symbols = new Set();
  const note = (s) => {
    if (!s) return;
    symbols.add(s);
    if (s.flags & ts.SymbolFlags.Alias) symbols.add(resolve(s));
  };
  const typeOfSym = (s) => {
    const d = s.valueDeclaration ?? s.declarations?.[0];
    return d ? checker.getTypeOfSymbolAtLocation(s, d) : null;
  };
  // `T | undefined` and `T | null` drop T's alias symbol, so also follow the references as written.
  const refsIn = (node, stack) => {
    if (!node) return;
    if (
      ts.isTypeReferenceNode(node) ||
      ts.isExpressionWithTypeArguments(node) ||
      ts.isTypeQueryNode(node)
    ) {
      // Interned unions keep the alias of whichever declaration created them first, so record the name too.
      const name = ts.isTypeReferenceNode(node)
        ? node.typeName
        : ts.isTypeQueryNode(node)
          ? node.exprName
          : node.expression;
      const sym = checker.getSymbolAtLocation(name);
      if (sym) {
        note(sym);
        if (resolve(sym).flags & ts.SymbolFlags.TypeAlias)
          for (const d of resolve(sym).declarations ?? [])
            if (!seenAlias.has(d)) {
              seenAlias.add(d);
              refsIn(d.type, stack);
            }
      }
      stack.push(checker.getTypeFromTypeNode(node));
    }
    ts.forEachChild(node, (c) => refsIn(c, stack));
  };
  const visitSig = (sig, stack) => {
    refsIn(sig.declaration, stack);
    for (const p of sig.getParameters()) refsIn(p.valueDeclaration?.type, stack);
    for (const tp of sig.getTypeParameters() ?? []) {
      const c = tp.getConstraint();
      if (c) stack.push(c);
    }
    for (const p of sig.getParameters()) {
      const t = typeOfSym(p);
      if (t) stack.push(t);
    }
    stack.push(sig.getReturnType());
  };
  const stack = [];
  for (const v of values) {
    const target = resolve(v);
    const t = typeOfSym(target);
    if (t) stack.push(t);
    if (target.flags & ts.SymbolFlags.Class) stack.push(checker.getDeclaredTypeOfSymbol(target));
  }
  while (stack.length) {
    const t = stack.pop();
    if (!t || seenTypes.has(t)) continue;
    seenTypes.add(t);
    note(t.aliasSymbol);
    note(t.symbol);
    for (const a of t.aliasTypeArguments ?? []) stack.push(a);
    if (t.isUnionOrIntersection()) stack.push(...t.types);
    if (t.flags & ts.TypeFlags.Object && t.objectFlags & ts.ObjectFlags.Reference)
      stack.push(...checker.getTypeArguments(t));
    if (t.flags & ts.TypeFlags.Index) stack.push(t.type);
    if (t.flags & ts.TypeFlags.IndexedAccess) stack.push(t.objectType, t.indexType);
    if (t.flags & ts.TypeFlags.Conditional) stack.push(t.checkType, t.extendsType);
    // Walk members only of this package family's own types, not lib or DOM types.
    const decl = (t.aliasSymbol ?? t.symbol)?.declarations?.[0];
    if (!decl || !ownFile(decl)) continue;
    for (const sig of [...t.getCallSignatures(), ...t.getConstructSignatures()])
      visitSig(sig, stack);
    for (const p of checker.getPropertiesOfType(t)) {
      for (const d of p.declarations ?? []) refsIn(d.type, stack);
      const pt = typeOfSym(p);
      if (pt) stack.push(pt);
    }
    for (const info of checker.getIndexInfosOfType(t)) stack.push(info.type);
    for (const b of t.isClassOrInterface() ? checker.getBaseTypes(t) : []) stack.push(b);
    if (t.aliasSymbol) for (const d of t.aliasSymbol.declarations ?? []) refsIn(d.type, stack);
  }
  return symbols;
};

for (const [entry, spec] of Object.entries(manifest.entries)) {
  const file = entries.get(entry);
  if (!file || !fs.existsSync(file)) continue;
  const syms = exportsOf(file);
  const names = new Set(syms.map((s) => s.getName()));
  const byName = new Map(syms.map((s) => [s.getName(), s]));
  for (const [n, kind] of Object.entries(spec.mustExport)) {
    const sym = byName.get(n);
    if (!sym) violations.push(`${entry}: missing export ${n}`);
    else if ((kind === 'value') !== isValueExport(sym))
      violations.push(`${entry}: ${n} must be exported as a ${kind}`);
  }
  for (const n of spec.mustNotExport)
    if (names.has(n)) violations.push(`${entry}: must not export ${n}`);
  if (entry === '@libraz/formulon-cell') {
    const values = syms.filter(isValueExport);
    if (values.length > manifest.rootValueMax)
      violations.push(`${entry}: ${values.length} runtime values > ${manifest.rootValueMax}`);
    const reached = reachableTypes(values);
    for (const s of syms)
      if (!isValueExport(s) && !reached.has(s) && !reached.has(resolve(s)))
        violations.push(
          `${entry}: orphan type ${s.getName()} (no exported value's signature reaches it)`,
        );
  }
  if (checkAbsence) for (const s of syms) scanAbsence(s, `${entry}#${s.getName()}`);
}
/**
 * Configuration is data. A function anywhere under `SpreadsheetOptions` is either an event
 * callback (events go through `on()`, D6) or data passed as code (pass the value, D24); only
 * the keys `manifest.optionFunctionAllow` names, each a host service or behaviour hook with its
 * reason, may hold functions.
 */
const optionFunctions = (rootFile) => {
  const sym = exportsOf(rootFile).find((s) => s.getName() === 'SpreadsheetOptions');
  if (!sym) return ['SpreadsheetOptions not exported'];
  const found = [];
  const seen = new Set();
  const walk = (type, at) => {
    if (seen.has(type)) return;
    seen.add(type);
    if (type.isUnion()) {
      for (const t of type.types) walk(t, at);
      return;
    }
    if (type.getCallSignatures().length > 0) {
      found.push(at);
      return;
    }
    const elem = checker.isArrayType?.(type) ? checker.getTypeArguments(type)[0] : null;
    if (elem) {
      walk(elem, `${at}[]`);
      return;
    }
    const typeSym = type.getSymbol();
    const decl = typeSym?.declarations?.[0];
    // A class instance (a Workbook, an AbortSignal) is an object handed over, not configuration.
    if (!decl || !ownFile(decl) || (typeSym.flags & ts.SymbolFlags.Class) !== 0) return;
    for (const p of checker.getPropertiesOfType(type)) {
      const pd = p.valueDeclaration ?? p.declarations?.[0];
      if (pd) walk(checker.getTypeOfSymbolAtLocation(p, pd), `${at}.${p.getName()}`);
    }
    for (const info of checker.getIndexInfosOfType(type)) walk(info.type, `${at}[key]`);
  };
  const options = checker.getDeclaredTypeOfSymbol(resolve(sym));
  for (const p of checker.getPropertiesOfType(options)) {
    if (Object.hasOwn(manifest.optionFunctionAllow ?? {}, p.getName())) continue;
    const pd = p.valueDeclaration ?? p.declarations?.[0];
    if (pd) walk(checker.getTypeOfSymbolAtLocation(p, pd), `SpreadsheetOptions.${p.getName()}`);
  }
  return found;
};
{
  const rootFile = entries.get('@libraz/formulon-cell');
  if (rootFile && fs.existsSync(rootFile))
    for (const at of optionFunctions(rootFile))
      violations.push(`options: ${at} is a function (configuration is data; events use on())`);
}
for (const f of i18nFiles)
  if (!manifest.i18nLocales.includes(f))
    violations.push(`@libraz/formulon-cell/i18n/*: unexpected entry ${f}`);
if (checkAbsence) for (const a of absence) violations.push(`absence: ${a} can return undefined`);

// The Sheet/Workbook augmentations must survive into the built root declaration. `AUGMENTATION_DIST`
// points the root entry at another dist directory (used to prove the check red and green).
{
  const dist = process.env.AUGMENTATION_DIST;
  const dir = path.join(root, 'scripts/api-surface/augmentation');
  const project = dist
    ? path.join(dir, '.tsconfig.override.json')
    : path.join(dir, 'tsconfig.json');
  if (dist)
    fs.writeFileSync(
      project,
      JSON.stringify({
        extends: './tsconfig.json',
        compilerOptions: { paths: { '@libraz/formulon-cell': [path.join(dist, 'index.d.ts')] } },
      }),
    );
  const tsc = spawnSync(path.join(root, 'node_modules/.bin/tsc'), ['-p', project], {
    encoding: 'utf8',
  });
  if (dist) fs.rmSync(project);
  if (tsc.status !== 0)
    for (const line of (tsc.stdout + tsc.stderr).split('\n').filter(Boolean))
      violations.push(`augmentation: ${line}`);
}

const byKind = {};
for (const v of violations) {
  const k = v.split(':')[0];
  byKind[k] = (byKind[k] ?? 0) + 1;
}
if (violations.length) {
  console.log(violations.slice(0, Number(process.env.API_SURFACE_MAX ?? 200)).join('\n'));
  if (violations.length > Number(process.env.API_SURFACE_MAX ?? 200)) console.log(`… more`);
}
console.log(`api-surface: ${violations.length} violation(s)`, JSON.stringify(byKind));
process.exit(violations.length ? 1 : 0);
