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
const { wrapText } = await server.ssrLoadModule('/src/render/text-wrap.ts');

const text = 'The quick brown fox jumps over the lazy dog. '.repeat(6);
let constructions = 0;
let measureCalls = 0;
let checksum = 0;
const NativeSegmenter = Intl.Segmenter;
Intl.Segmenter = class extends NativeSegmenter {
  constructor(...args) {
    constructions += 1;
    super(...args);
  }
};

const wrap = () => {
  const lines = wrapText(text, {
    maxWidth: 130,
    measureText(value) {
      measureCalls += 1;
      return value.length * 7;
    },
    locale: 'en-US',
  });
  return lines;
};

const firstLineCount = wrap().length;
const firstConstructions = constructions;
const firstMeasureCalls = measureCalls;
for (let i = 0; i < 20; i += 1) wrap();
const repeatedConstructionsBefore = constructions;
const repeatedMeasureBefore = measureCalls;
const start = performance.now();
for (let i = 0; i < 500; i += 1) checksum += wrap().length;
const elapsed = performance.now() - start;

process.stdout.write(
  `First wrap: ${firstConstructions} segmenters, ${firstMeasureCalls} measureText calls, ${firstLineCount} lines\n`,
);
process.stdout.write(
  `Repeated wrap: ${constructions - repeatedConstructionsBefore} segmenters, ${(measureCalls - repeatedMeasureBefore) / 500} measureText calls/wrap\n`,
);
process.stdout.write(
  `500 wraps after 20 warmups: ${elapsed.toFixed(1)} ms; checksum ${checksum}\n`,
);

Intl.Segmenter = NativeSegmenter;
await server.close();
