import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { cp, mkdir, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const upstream = resolve(process.env.FORMULON_SOURCE_DIR ?? resolve(root, '../formulon'));
const single = 'build-wasm-cell-locales';
const threads = 'build-wasm-threads-cell-locales';
const output = join(root, '.yarn/local');
const archive = join(output, 'formulon-language-profiles.tgz');

function run(command, args, cwd = upstream, env = process.env) {
  const result = spawnSync(command, args, { cwd, env, stdio: 'inherit' });
  if (result.error) throw result.error;
  if (result.status !== 0) throw new Error(`${command} exited with ${result.status}`);
}

// Separate output directories keep this build independent of oracle capture.
run('make', [
  'wasm',
  'wasm-threads',
  `WASM_BUILD_DIR=${single}`,
  `WASM_THREADS_BUILD_DIR=${threads}`,
  `JOBS=${process.env.FORMULON_BUILD_JOBS ?? '4'}`,
]);
const staging = await mkdtemp(join(tmpdir(), 'formulon-cell-npm-'));
const pkg = join(staging, 'package');
try {
  await mkdir(pkg);
  for (const file of ['package.json', 'README.md', 'LICENSE', 'NOTICE', 'test']) {
    await cp(join(upstream, 'packages/npm', file), join(pkg, file), { recursive: true });
  }
  run(process.execPath, [
    'packages/npm/scripts/stage.mjs',
    '--build-dir',
    single,
    '--threads-build-dir',
    threads,
    '--out-dir',
    join(pkg, 'dist'),
  ]);
  for (const entry of ['.', './threads']) {
    run(process.execPath, ['--test', 'test/smoke.test.mjs'], pkg, {
      ...process.env,
      FORMULON_NPM_ENTRY: entry,
    });
    run(process.execPath, [join(root, 'scripts/local-formulon-locale-smoke.mjs'), pkg], pkg, {
      ...process.env,
      FORMULON_NPM_ENTRY: entry,
    });
  }
  // npm's archive format is a tarball containing a top-level package/ directory.
  await rm(join(pkg, 'test'), { recursive: true });
  await mkdir(output, { recursive: true });
  run('tar', ['-czf', archive, '-C', staging, 'package'], root);
  const sha256 = async (file) =>
    createHash('sha256')
      .update(await readFile(file))
      .digest('hex');
  await writeFile(
    join(output, 'formulon-language-profiles.json'),
    `${JSON.stringify(
      {
        builtAt: new Date().toISOString(),
        upstream,
        version: JSON.parse(await readFile(join(pkg, 'package.json'), 'utf8')).version,
        archiveSha256: await sha256(archive),
        wasmSha256: await sha256(join(pkg, 'dist/formulon.wasm')),
        threadsWasmSha256: await sha256(join(pkg, 'dist/formulon_threads.wasm')),
      },
      null,
      2,
    )}\n`,
  );
  console.log(`Local npm archive: ${archive}`);
} finally {
  await rm(staging, { recursive: true, force: true });
}
