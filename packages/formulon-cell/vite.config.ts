import { resolve } from 'node:path';
import dts from 'unplugin-dts/vite';
import { defineConfig } from 'vite';

// Library build for `@libraz/formulon-cell`.
//
// Output uses `preserveModules` so dist/ mirrors src/. That keeps subpath
// declarations stable while the calc engine remains an external dependency.
export default defineConfig({
  build: {
    target: 'es2022',
    sourcemap: false,
    assetsInlineLimit: 0,
    minify: 'oxc',
    reportCompressedSize: false,
    lib: {
      entry: {
        index: resolve(__dirname, 'src/index.ts'),
        'extensions/index': resolve(__dirname, 'src/extensions/index.ts'),
        'i18n/ja': resolve(__dirname, 'src/i18n/ja.ts'),
        'i18n/en': resolve(__dirname, 'src/i18n/en.ts'),
      },
      formats: ['es'],
    },
    rollupOptions: {
      // The Emscripten bundle uses TLA + new URL(...) for both the wasm and
      // worker scripts. Vite's lib worker plugin tries to bundle those as
      // iife workers, which breaks on TLA. Externalizing keeps the import
      // alive so the user's bundler handles the engine package assets.
      external: ['zustand', 'zustand/vanilla', '@libraz/formulon'],
      output: {
        preserveModules: true,
        preserveModulesRoot: 'src',
        entryFileNames: '[name].js',
      },
    },
  },
  esbuild: {
    legalComments: 'none',
  },
  plugins: [
    dts({
      // Use a build-only tsconfig that pins `rootDir: ./src` and excludes
      // tests/. TS 6 made `rootDir` default to the tsconfig directory, so
      // the regular tsconfig (which includes tests/) would otherwise emit
      // .d.ts under dist/src/, breaking the package's exports map.
      tsconfigPath: './tsconfig.build.json',
      // No rollupTypes — we want per-file .d.ts so subpath imports get
      // their own type declarations.
    }),
  ],
});
