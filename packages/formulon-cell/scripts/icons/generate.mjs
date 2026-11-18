#!/usr/bin/env node
/**
 * Icon generator: gridsmith specs in `icons/specs` become the registry module and the CSS
 * data URIs marked in `src/styles`.
 *
 *   generate.mjs                          lint every spec, then write module, lock and CSS URIs
 *   generate.mjs --render <name…>         sheet, PNGs and a report per spec into scripts/icons/out/specs
 *   generate.mjs --lint --group <g…>      lint groups (`parts` too); --expect <names.json…>, --constraints <file>
 *   generate.mjs --guide [topic]          print the pinned gridsmith reference
 *   generate.mjs --check                  staleness check against the lock; needs no gridsmith
 *   generate.mjs --check --render         also compare the lock's pin with yarn.lock and re-render everything
 *
 * The expected-names files map a group to its names: `{ "<group>": ["name", …] }`.
 */

import { existsSync } from 'node:fs';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { createRequire } from 'node:module';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import {
  CSS_GROUP,
  checkTree,
  classViolations,
  constraintViolations,
  dataUriFor,
  emitModule,
  lintSpecRules,
  MODULE_PATH,
  parseCssUris,
  readTree,
  rewriteCssUris,
  sha256,
  specHash,
  toIconSegments,
  uriKey,
  validateTokens,
  yarnLockGridsmithCommit,
} from './emit.mjs';

const HERE = dirname(fileURLToPath(import.meta.url));
const PKG = resolve(HERE, '../..');
const REPO = resolve(PKG, '../..');
const OUT = join(HERE, 'out/specs');
const LOCK_PATH = join(PKG, 'icons/icons.lock.json');
const PNG_SIZES = [16, 18, 20, 32];

const argv = process.argv.slice(2);
const has = (flag) => argv.includes(flag);
const valuesOf = (flag) => {
  const at = argv.indexOf(flag);
  if (at < 0) return [];
  const out = [];
  for (let i = at + 1; i < argv.length && !argv[i].startsWith('--'); i += 1) out.push(argv[i]);
  return out;
};

const fail = (lines) => {
  for (const line of lines) console.error(line);
  process.exit(1);
};

const writeIfChanged = async (path, text) => {
  if (existsSync(path) && (await readFile(path, 'utf8')) === text) return false;
  await mkdir(dirname(path), { recursive: true });
  await writeFile(path, text);
  return true;
};

const stringifyLock = (lock) => `${JSON.stringify(lock, null, 2)}\n`;

const pinnedCommit = async () =>
  yarnLockGridsmithCommit(await readFile(join(REPO, 'yarn.lock'), 'utf8'));

/** gridsmith ships no build from a git install, so bundle its TypeScript sources on demand. */
const loadGridsmith = async () => {
  try {
    return await import('@libraz/gridsmith');
  } catch (error) {
    if (error.code !== 'ERR_MODULE_NOT_FOUND') throw error;
  }
  const require = createRequire(import.meta.url);
  const root = dirname(require.resolve('@libraz/gridsmith/package.json'));
  const commit = (await pinnedCommit()) ?? 'unpinned';
  const cached = join(PKG, 'node_modules/.cache/icons', `gridsmith-${commit.slice(0, 12)}.mjs`);
  if (!existsSync(cached)) {
    const { rolldown } = await import('rolldown');
    const bundle = await rolldown({
      input: join(root, 'src/index.ts'),
      platform: 'node',
      external: [/^[^./]/],
      resolve: { extensionAlias: { '.js': ['.ts', '.js'] } },
    });
    const { output } = await bundle.generate({ format: 'esm' });
    await bundle.close();
    await mkdir(dirname(cached), { recursive: true });
    await writeFile(cached, output[0].code);
  }
  return import(pathToFileURL(cached).href);
};

const readConstraints = async (path) => {
  if (!path) return {};
  const data = JSON.parse(await readFile(resolve(path), 'utf8'));
  const list = Array.isArray(data) ? data : (data.constraints ?? []);
  return Object.fromEntries(list.map((entry) => [entry.name, entry]));
};

const readGroupSpecs = async (tree, groups) => {
  const out = {};
  for (const group of groups) {
    if (group !== 'parts') continue;
    const { readdir } = await import('node:fs/promises');
    const dir = join(PKG, 'icons/parts');
    const files = existsSync(dir) ? (await readdir(dir)).filter((f) => f.endsWith('.json')) : [];
    for (const file of files.sort()) {
      out[file.slice(0, -5)] = {
        group,
        spec: JSON.parse(await readFile(join(dir, file), 'utf8')),
      };
    }
  }
  for (const [name, entry] of Object.entries(tree.specs)) {
    if (groups.includes(entry.group)) out[name] = entry;
  }
  return out;
};

/** gridsmith's own violations plus the house rules, per spec. */
const lintOne = (gs, tree, name, entry, constraint) => {
  const findings = [];
  let drawing = null;
  try {
    drawing = gs.render(entry.spec);
  } catch (error) {
    return { drawing, findings: [`${name}: ${error.message}`] };
  }
  for (const v of drawing.violations) findings.push(`${name}: ${v.rule} ${v.target}: ${v.message}`);
  const approved = tree.tokens.approvedExceptions?.[name] ?? [];
  for (const v of lintSpecRules({
    name,
    spec: entry.spec,
    tokens: tree.tokens,
    approved,
    constraint,
  })) {
    findings.push(`${name}: ${v.rule} ${v.target}: ${v.message}`);
  }
  if (!drawing.segments) findings.push(`${name}: spec is not an icon`);
  else if (constraint) {
    const segs = toIconSegments(drawing.segments.segments);
    for (const m of constraintViolations({ constraint, segments: segs, tokens: tree.tokens })) {
      findings.push(`${name}: constraint ${m}`);
    }
  }
  return { drawing, findings };
};

const measure = async (drawn) => {
  const { measureAll } = await import(pathToFileURL(join(REPO, 'scripts/icons/render.mjs')).href);
  return measureAll(
    drawn.map(([name, drawing]) => [name, toIconSegments(drawing.segments.segments)]),
  );
};

const expectedNames = async (files, groups) => {
  const names = new Set();
  for (const file of files) {
    const data = JSON.parse(await readFile(resolve(file), 'utf8'));
    for (const group of groups) for (const name of data[group] ?? []) names.add(name);
  }
  return names;
};

/** Everything the write mode produces, as text, from the specs and the CSS on disk. */
const buildOutputs = (gs, tree, commit) => {
  const drawn = {};
  const findings = [];
  for (const [name, entry] of Object.entries(tree.specs)) {
    const result = lintOne(gs, tree, name, entry, null);
    findings.push(...result.findings);
    drawn[name] = result.drawing;
  }
  if (findings.length > 0) return { findings };
  const segmentsOf = (name) => toIconSegments(drawn[name].segments.segments);
  const icons = {};
  for (const [name, entry] of Object.entries(tree.specs)) {
    if (entry.group !== CSS_GROUP) icons[name] = segmentsOf(name);
  }
  const moduleText = emitModule({ palette: tree.tokens.palette, icons });
  const css = {};
  const uris = {};
  for (const [file, text] of Object.entries(tree.css)) {
    const next = rewriteCssUris(text, tree.tokens, (name, _consumer, theme) => {
      if (!drawn[name]) throw new Error(`${file}: marker names unknown icon ${name}`);
      return dataUriFor(segmentsOf(name), tree.tokens, theme);
    });
    css[file] = next;
    for (const entry of parseCssUris(next, tree.tokens).uris) {
      if (entry.kind === 'icon') {
        uris[uriKey(entry.name, entry.consumer, entry.theme)] = sha256(entry.uri);
      }
    }
  }
  const lock = {
    gridsmith: { commit },
    specs: Object.fromEntries(
      Object.keys(tree.specs)
        .sort()
        .map((name) => [
          name,
          { group: tree.specs[name].group, sha256: specHash(tree.specs[name].spec) },
        ]),
    ),
    module: { sha256: sha256(moduleText) },
    css: Object.keys(tree.specs)
      .filter((name) => tree.specs[name].group === CSS_GROUP)
      .sort(),
    uris: Object.fromEntries(Object.entries(uris).sort(([a], [b]) => (a < b ? -1 : 1))),
  };
  return { findings, moduleText, css, lock };
};

const runCheck = async () => {
  const tree = await readTree(PKG);
  const { errors, unmarked } = checkTree(tree);
  errors.push(...validateTokens(tree.tokens));
  if (has('--render')) {
    const gs = await loadGridsmith();
    const pinned = await pinnedCommit();
    if (pinned !== tree.lock.gridsmith.commit) {
      errors.push(
        `yarn.lock pins gridsmith ${pinned}, the lock was written with ${tree.lock.gridsmith.commit}`,
      );
    }
    const built = buildOutputs(gs, tree, tree.lock.gridsmith.commit);
    errors.push(...built.findings);
    if (built.findings.length === 0) {
      if (built.moduleText !== tree.moduleText)
        errors.push(`${MODULE_PATH} differs from a fresh render`);
      for (const [file, text] of Object.entries(built.css)) {
        if (text !== tree.css[file]) errors.push(`${file}: CSS URIs differ from a fresh render`);
      }
    }
  }
  if (errors.length > 0) fail(errors);
  const note = tree.tokens.requireMarkers
    ? ''
    : `; ${unmarked} data:image/svg URIs carry no marker yet`;
  console.log(`icons: ${Object.keys(tree.specs).length} specs match the lock${note}`);
};

const runWrite = async () => {
  const gs = await loadGridsmith();
  const tree = await readTree(PKG);
  const tokenErrors = validateTokens(tree.tokens);
  if (tokenErrors.length > 0) fail(tokenErrors);
  const commit = await pinnedCommit();
  if (!commit) fail(['yarn.lock has no gridsmith resolution; run yarn install']);
  const built = buildOutputs(gs, tree, commit);
  if (built.findings.length > 0) fail(built.findings);
  const changed = [];
  if (await writeIfChanged(join(PKG, MODULE_PATH), built.moduleText)) changed.push(MODULE_PATH);
  for (const [file, text] of Object.entries(built.css)) {
    if (await writeIfChanged(join(PKG, file), text)) changed.push(file);
  }
  if (await writeIfChanged(LOCK_PATH, stringifyLock(built.lock)))
    changed.push('icons/icons.lock.json');
  console.log(changed.length > 0 ? `wrote ${changed.join(', ')}` : 'already up to date');
};

const runLint = async () => {
  const gs = await loadGridsmith();
  const tree = await readTree(PKG);
  const groups = valuesOf('--group');
  if (groups.length === 0) fail(['--lint needs --group <name…>']);
  const constraints = await readConstraints(valuesOf('--constraints')[0]);
  const specs = await readGroupSpecs(tree, groups);
  const findings = [];
  const drawn = [];
  for (const [name, entry] of Object.entries(specs)) {
    const result = lintOne(gs, tree, name, entry, constraints[name] ?? null);
    findings.push(...result.findings);
    if (result.drawing?.segments) drawn.push([name, result.drawing]);
  }
  const expectFiles = valuesOf('--expect');
  if (expectFiles.length > 0) {
    const expected = await expectedNames(expectFiles, groups);
    for (const name of expected) if (!specs[name]) findings.push(`${name}: expected but no spec`);
    for (const name of Object.keys(specs))
      if (!expected.has(name)) findings.push(`${name}: spec not in the expected names`);
  }
  if (drawn.length > 0) {
    for (const row of await measure(drawn)) {
      for (const message of classViolations({
        name: row.name,
        group: specs[row.name].group,
        metrics: row,
        tokens: tree.tokens,
      })) {
        findings.push(`${row.name}: ${message}`);
      }
    }
  }
  if (findings.length > 0) fail(findings);
  console.log(`lint: ${Object.keys(specs).length} specs in ${groups.join(', ')} clean`);
};

const runRender = async () => {
  const names = valuesOf('--render');
  if (names.length === 0) fail(['--render needs spec names']);
  const gs = await loadGridsmith();
  const tree = await readTree(PKG);
  const constraints = await readConstraints(valuesOf('--constraints')[0]);
  const specs = await readGroupSpecs(tree, [
    'parts',
    ...new Set(Object.values(tree.specs).map((e) => e.group)),
  ]);
  const commit = await pinnedCommit();
  for (const name of names) {
    const entry = specs[name];
    if (!entry) fail([`${name}: no such spec`]);
    const dir = join(OUT, name);
    await mkdir(dir, { recursive: true });
    const { drawing, findings } = lintOne(gs, tree, name, entry, constraints[name] ?? null);
    if (!drawing) fail(findings);
    await writeFile(join(dir, 'sheet.svg'), drawing.sheetSvg);
    const sizes = entry.group === CSS_GROUP ? [12, ...PNG_SIZES] : PNG_SIZES;
    for (const size of sizes) await writeFile(join(dir, `${size}.png`), drawing.png(size));
    const [row] = drawing.segments ? await measure([[name, drawing]]) : [null];
    const report = {
      name,
      group: entry.group,
      violations: findings,
      ink: row && { w: row.w, h: row.h, cx: row.cx, cy: row.cy, bleed: row.bleed },
      centreOffset: row && Math.max(Math.abs(row.cx - 12), Math.abs(row.cy - 12)),
      closestGap: row?.gap ?? null,
      classCheck: row
        ? classViolations({ name, group: entry.group, metrics: row, tokens: tree.tokens })
        : null,
    };
    await writeFile(join(dir, 'report.json'), `${JSON.stringify(report, null, 2)}\n`);
    await writeFile(
      join(dir, 'tool-log.txt'),
      [
        `command: generate.mjs ${argv.join(' ')}`,
        `gridsmith: ${commit}`,
        `spec sha256: ${specHash(entry.spec)}`,
        '',
      ].join('\n'),
    );
    const problems = [...findings, ...(report.classCheck ?? [])];
    console.log(`${name}: ${problems.length === 0 ? 'clean' : problems.join('; ')}  -> ${dir}`);
  }
};

const main = async () => {
  if (has('--guide')) {
    const topic = valuesOf('--guide')[0];
    const doc = (await loadGridsmith()).guide(topic);
    console.log(typeof doc.text === 'string' ? doc.text : JSON.stringify(doc, null, 2));
  } else if (has('--check')) await runCheck();
  else if (has('--lint')) await runLint();
  else if (has('--render')) await runRender();
  else await runWrite();
};

await main();
