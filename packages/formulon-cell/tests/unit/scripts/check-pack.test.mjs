import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { gzipSync } from 'node:zlib';
import { afterAll, describe, expect, it } from 'vitest';
import {
  ALLOW_ENGINE_TARBALL,
  checkPackage,
  readTarball,
} from '../../../../../scripts/check-pack.mjs';

const repo = path.resolve(import.meta.dirname, '../../../../..');
const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'check-pack-'));
afterAll(() => fs.rmSync(dir, { recursive: true, force: true }));

/** Writes a gzipped ustar archive with every file under `package/`, as npm and yarn pack it. */
function tarball(files) {
  const blocks = [];
  for (const [name, content] of Object.entries(files)) {
    const body = Buffer.from(content);
    const header = Buffer.alloc(512);
    header.write(`package/${name}`, 0);
    header.write('0000644\0', 100);
    header.write(`${body.length.toString(8).padStart(11, '0')}\0`, 124);
    header.write('0', 156);
    header.write('ustar\0', 257);
    blocks.push(header, body, Buffer.alloc((512 - (body.length % 512)) % 512));
  }
  const file = path.join(dir, `${blocks.length}-${Math.random().toString(36).slice(2)}.tgz`);
  fs.writeFileSync(file, gzipSync(Buffer.concat([...blocks, Buffer.alloc(1024)])));
  return file;
}

const manifest = (overrides = {}) =>
  JSON.stringify({
    name: '@scope/pkg',
    version: '1.0.0',
    exports: {
      '.': { types: './dist/index.d.ts', import: './dist/index.js' },
      './i18n/*': './dist/i18n/*.js',
      './package.json': './package.json',
    },
    files: ['dist', 'CHANGELOG.md'],
    dependencies: { engine: '1.2.3' },
    ...overrides,
  });

const clean = {
  'package.json': manifest(),
  'README.md': '#',
  LICENSE: 'license',
  'CHANGELOG.md': '#',
  'dist/index.js': 'export {};',
  'dist/index.d.ts': '/** The entry point. */\nexport {};',
  'dist/i18n/en.js': 'export {};',
  'dist/styles/README.md': '#',
};

const check = (files, options) => checkPackage(readTarball(tarball(files)), options).problems;

describe('packed-package guard', () => {
  it('passes a package that ships what its manifest lists', () => {
    expect(check(clean)).toEqual([]);
  });

  it.each(['file:/tmp/engine.tgz', 'link:../engine', 'portal:../engine', 'workspace:*'])(
    'rejects a %s dependency',
    (value) => {
      const files = { ...clean, 'package.json': manifest({ dependencies: { engine: value } }) };
      expect(check(files)).toEqual([`dependencies.engine is "${value}"`]);
    },
  );

  it('warns instead of failing on the engine tarball only when asked to', () => {
    const pinned = '@libraz/formulon';
    const value = 'file:/tmp/formulon/libraz-formulon-1.0.0.tgz';
    const entries = readTarball(
      tarball({ ...clean, 'package.json': manifest({ dependencies: { [pinned]: value } }) }),
    );
    expect(checkPackage(entries).problems).toEqual([`dependencies.${pinned} is "${value}"`]);
    expect(checkPackage(entries, { allowEngineTarball: true })).toEqual({
      problems: [],
      warnings: [`dependencies.${pinned} is "${value}" (allowed for this run only)`],
    });
  });

  it.each([
    ['another dependency', { dependencies: { other: 'file:/tmp/other-1.0.0.tgz' } }],
    ['the engine as a directory', { dependencies: { '@libraz/formulon': 'file:../formulon' } }],
    ['the engine through link:', { dependencies: { '@libraz/formulon': 'link:/tmp/x.tgz' } }],
    ['the engine as a peer', { peerDependencies: { '@libraz/formulon': 'file:/tmp/x.tgz' } }],
  ])('keeps failing %s when the engine tarball is allowed', (_, fields) => {
    const files = { ...clean, 'package.json': manifest(fields), 'dist/index.d.ts.map': '{}' };
    const problems = check(files, { allowEngineTarball: true });
    expect(problems).toHaveLength(2);
    expect(problems).toContain('dist/index.d.ts.map is a declaration map');
  });

  it.each(['.github/workflows/ci.yml', '.github/workflows/publish.yml'])(
    'is never relaxed by %s',
    (workflow) => {
      expect(fs.readFileSync(path.join(repo, workflow), 'utf8')).not.toContain(
        ALLOW_ENGINE_TARBALL,
      );
    },
  );

  it('rejects declaration maps and declarations that name one', () => {
    const files = {
      ...clean,
      'dist/index.d.ts': '/** The entry point. */\nexport {};\n//# sourceMappingURL=index.d.ts.map',
      'dist/index.d.ts.map': '{}',
    };
    expect(check(files)).toEqual([
      'dist/index.d.ts names a source map',
      'dist/index.d.ts.map is a declaration map',
    ]);
  });

  it('rejects declarations stripped of every doc comment', () => {
    expect(check({ ...clean, 'dist/index.d.ts': 'export {};' })).toEqual([
      'no declaration file carries a doc comment',
    ]);
  });

  it('rejects a file outside "files"', () => {
    expect(check({ ...clean, 'src/index.ts': 'export {};' })).toEqual([
      'src/index.ts is outside "files"',
    ]);
  });

  it('rejects an exports target the package does not contain', () => {
    const { 'dist/index.js': _, 'dist/i18n/en.js': __, ...files } = clean;
    expect(check(files)).toEqual([
      'target ./dist/index.js is not in the package',
      'target ./dist/i18n/*.js is not in the package',
    ]);
  });

  it('reads a path longer than the ustar name field from its pax header', () => {
    const long = `dist/${'nested/'.repeat(16)}index.js`;
    const field = ` path=package/${long}\n`;
    const digits = String(field.length + 3).length;
    const record = `${field.length + digits}${field}`;
    const header = (name, size, type) => {
      const block = Buffer.alloc(512);
      block.write(name, 0);
      block.write(`${size.toString(8).padStart(11, '0')}\0`, 124);
      block.write(type, 156);
      return block;
    };
    const pad = (text) => Buffer.concat([Buffer.from(text), Buffer.alloc(512 - text.length)]);
    const file = path.join(dir, 'pax.tgz');
    const archive = [header('PaxHeader', record.length, 'x'), pad(record)];
    archive.push(header('package/truncated', 2, '0'), pad('{}'), Buffer.alloc(1024));
    fs.writeFileSync(file, gzipSync(Buffer.concat(archive)));
    expect([...readTarball(file).keys()]).toEqual([long]);
  });
});
