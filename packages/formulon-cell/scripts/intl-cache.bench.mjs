import { resolve } from 'node:path';
import { performance } from 'node:perf_hooks';
import { fileURLToPath } from 'node:url';
import { createServer } from 'vite';

const packageRoot = resolve(fileURLToPath(new URL('..', import.meta.url)));
const server = await createServer({
  appType: 'custom',
  configFile: false,
  logLevel: 'error',
  root: packageRoot,
  server: { middlewareMode: true, hmr: false },
});
const { formatNumber } = await server.ssrLoadModule('/src/format/number-format.ts');

function benchmark(label, operation, iterations, warmup) {
  for (let i = 0; i < warmup; i += 1) operation();
  const start = performance.now();
  for (let i = 0; i < iterations; i += 1) operation();
  const elapsed = performance.now() - start;
  process.stdout.write(`${label}: ${elapsed.toFixed(1)} ms / ${iterations} calls\n`);
}

benchmark('General number', () => formatNumber(1234.5, undefined, 'de-DE'), 20_000, 1_000);
benchmark(
  'Fixed number',
  () => formatNumber(1234.5, { kind: 'fixed', decimals: 2, thousands: true }, 'de-DE'),
  20_000,
  1_000,
);
benchmark(
  'Numeric date',
  () => formatNumber(45292, { kind: 'date', pattern: 'yyyy-mm-dd' }, 'de-DE'),
  500,
  10,
);
benchmark(
  'Time only',
  () => formatNumber(0.5, { kind: 'time', pattern: 'hh:mm' }, 'de-DE'),
  500,
  10,
);
benchmark(
  'Localized long date',
  () => formatNumber(45651, { kind: 'custom', pattern: 'dddd, mmmm d, yyyy' }, 'de-DE'),
  500,
  10,
);

await server.close();
