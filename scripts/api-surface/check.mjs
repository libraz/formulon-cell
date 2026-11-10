#!/usr/bin/env node
/**
 * Public-surface gate. Reads the built declaration files reachable from the
 * `exports` maps of the three packages and compares them with manifest.json:
 * per entry point, names that must be exported and names that must not be,
 * the root-size ceiling, and the allowed `./i18n/*` files.
 *
 * Usage: node scripts/api-surface/check.mjs [--absence]
 *   --absence  also reports public return types and required readonly fields
 *              that include `undefined` (the surface uses `null` for absence).
 * Run after `yarn build`. Exit code 1 when any violation is found.
 */
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
      const base = path.join(root, dir, path.dirname(pre));
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

for (const [entry, spec] of Object.entries(manifest.entries)) {
  const file = entries.get(entry);
  if (!file || !fs.existsSync(file)) continue;
  const syms = exportsOf(file);
  const names = new Set(syms.map((s) => s.getName()));
  for (const n of spec.mustExport)
    if (!names.has(n)) violations.push(`${entry}: missing export ${n}`);
  for (const n of spec.mustNotExport)
    if (names.has(n)) violations.push(`${entry}: must not export ${n}`);
  if (entry === '@libraz/formulon-cell' && names.size > manifest.rootMax)
    violations.push(`${entry}: ${names.size} exports > ${manifest.rootMax}`);
  if (checkAbsence) for (const s of syms) scanAbsence(s, `${entry}#${s.getName()}`);
}
for (const f of i18nFiles)
  if (!manifest.i18nLocales.includes(f))
    violations.push(`@libraz/formulon-cell/i18n/*: unexpected entry ${f}`);
if (checkAbsence) for (const a of absence) violations.push(`absence: ${a} can return undefined`);

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
