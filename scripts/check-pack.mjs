#!/usr/bin/env node
/**
 * Packed-package guard. Reads each package tarball (.tgz) given on the command
 * line and fails when what would be uploaded is not what the manifest promises:
 *
 * - a dependency value that points at a local path or a workspace
 *   (`file:`, `link:`, `portal:`, `workspace:`);
 * - a declaration map (`*.d.ts.map`), or a `.d.ts` that still names one;
 * - declarations of which none carries a doc comment (a build that strips them);
 * - a file outside the manifest's `files` list and the always-included set;
 * - an `exports` (or `main` / `module` / `types`) target missing from the tarball.
 *
 * The checker does not produce tarballs. `yarn check:pack` packs the published
 * workspaces with yarn and runs it; the publish workflow runs it on `npm pack`
 * output, the same bytes `npm publish` uploads.
 *
 * Usage: node scripts/check-pack.mjs [--allow-engine-tarball] <package.tgz>...
 *   --allow-engine-tarball  reports a `file:` tarball for the calc engine dependency as a
 *                           warning instead of a violation. Only the local gate passes it,
 *                           while the engine is pinned to an unreleased build.
 * Exit code 1 on any violation.
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { gunzipSync } from 'node:zlib';

const DEPENDENCY_FIELDS = [
  'dependencies',
  'peerDependencies',
  'optionalDependencies',
  'devDependencies',
];
const LOCAL_PROTOCOL = /^(file|link|portal|workspace):/;
const ENGINE = '@libraz/formulon';
export const ALLOW_ENGINE_TARBALL = '--allow-engine-tarball';
const ROOT_ALWAYS_INCLUDED = /^(package\.json|readme(\..*)?|licen[cs]e(\..*)?)$/i;
// yarn packs a README at any depth and npm only at the root, so a nested one is allowed.
const NESTED_README = /(^|\/)readme(\.[^/]*)?$/i;

const cString = (field) => field.toString('utf8').replace(/\0.*$/s, '');
const octal = (field) => Number.parseInt(cString(field).trim() || '0', 8);

/** Reads a gzipped tarball into a map of path (without the leading `package/`) to contents. */
export function readTarball(file) {
  const tar = gunzipSync(fs.readFileSync(file));
  const entries = new Map();
  let longName = null;
  for (let at = 0; at + 512 <= tar.length; ) {
    const header = tar.subarray(at, at + 512);
    if (header.every((byte) => byte === 0)) break;
    const size = octal(header.subarray(124, 136));
    const type = String.fromCharCode(header[156] || 48);
    const body = tar.subarray(at + 512, at + 512 + size);
    at += 512 + Math.ceil(size / 512) * 512;
    if (type === 'x') {
      longName = body.toString('utf8').match(/\d+ path=([^\n]*)\n/)?.[1] ?? null;
      continue;
    }
    const prefix = cString(header.subarray(345, 500));
    const name = longName ?? (prefix ? `${prefix}/` : '') + cString(header.subarray(0, 100));
    longName = null;
    if (type === '0') entries.set(name.replace(/^[^/]+\//, ''), body);
  }
  return entries;
}

const exportTargets = (value) =>
  typeof value === 'string'
    ? [value]
    : value && typeof value === 'object'
      ? Object.values(value).flatMap(exportTargets)
      : [];

const escapeRegExp = (part) => part.replace(/[.+?^${}()|[\]\\]/g, '\\$&');

/**
 * Returns the rule violations for one unpacked package, and the warnings that
 * `allowEngineTarball` turns the engine's `file:` tarball dependency into.
 */
export function checkPackage(entries, { allowEngineTarball = false } = {}) {
  const raw = entries.get('package.json');
  if (!raw) return { problems: ['package.json is missing'], warnings: [] };
  const manifest = JSON.parse(raw.toString('utf8'));
  const problems = [];
  const warnings = [];

  for (const field of DEPENDENCY_FIELDS) {
    for (const [name, value] of Object.entries(manifest[field] ?? {})) {
      if (!LOCAL_PROTOCOL.test(value)) continue;
      const engineTarball =
        field === 'dependencies' && name === ENGINE && /^file:.+\.tgz$/.test(value);
      if (allowEngineTarball && engineTarball) {
        warnings.push(`${field}.${name} is "${value}" (allowed for this run only)`);
      } else problems.push(`${field}.${name} is "${value}"`);
    }
  }

  const listed = (manifest.files ?? []).map((entry) =>
    entry.replace(/^\.\//, '').replace(/\/$/, ''),
  );
  for (const [file, body] of entries) {
    if (file.endsWith('.d.ts.map')) problems.push(`${file} is a declaration map`);
    else if (file.endsWith('.d.ts') && body.includes('sourceMappingURL=')) {
      problems.push(`${file} names a source map`);
    }
    const inFiles = listed.some((entry) => file === entry || file.startsWith(`${entry}/`));
    if (!inFiles && !ROOT_ALWAYS_INCLUDED.test(file) && !NESTED_README.test(file)) {
      problems.push(`${file} is outside "files"`);
    }
  }

  const declarations = [...entries].filter(([file]) => file.endsWith('.d.ts'));
  if (declarations.length > 0 && !declarations.some(([, body]) => body.includes('/**'))) {
    problems.push('no declaration file carries a doc comment');
  }

  const targets = [
    ...exportTargets(manifest.exports),
    ...['main', 'module', 'types'].flatMap((key) => exportTargets(manifest[key])),
  ];
  for (const target of new Set(targets)) {
    const file = target.replace(/^\.\//, '');
    const pattern = new RegExp(`^${file.split('*').map(escapeRegExp).join('[^/]+')}$`);
    if (![...entries.keys()].some((entry) => pattern.test(entry))) {
      problems.push(`target ${target} is not in the package`);
    }
  }
  return { problems, warnings };
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const args = process.argv.slice(2);
  const allowEngineTarball = args.includes(ALLOW_ENGINE_TARBALL);
  const tarballs = args.filter((arg) => arg !== ALLOW_ENGINE_TARBALL);
  if (tarballs.length === 0) {
    console.error(`usage: check-pack.mjs [${ALLOW_ENGINE_TARBALL}] <package.tgz>...`);
    process.exit(2);
  }
  let failed = false;
  for (const tarball of tarballs) {
    const { problems, warnings } = checkPackage(readTarball(tarball), { allowEngineTarball });
    const status = problems.length === 0 ? 'ok' : `${problems.length} problem(s)`;
    console.log(`${path.basename(tarball)}: ${status}`);
    for (const warning of warnings) console.log(`  WARNING: ${warning}`);
    for (const problem of problems) console.log(`  ${problem}`);
    failed ||= problems.length > 0;
  }
  process.exit(failed ? 1 : 0);
}
